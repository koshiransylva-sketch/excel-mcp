"""openpyxl で Excel ファイル（.xlsx / .xlsm）を操作する MCP サーバー。

"""

import fnmatch
import importlib.util
import logging
import os
import re
import shutil
import tempfile
import zipfile
from copy import copy
from datetime import date, datetime, time
from pathlib import Path
from typing import Annotated, Any, Literal

import openpyxl as xl
from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import ToolAnnotations
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import (
    column_index_from_string,
    coordinate_from_string,
    range_boundaries,
)
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import (
    AutoFilter,
    CustomFilter,
    CustomFilters,
    FilterColumn,
    Filters,
)
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
log = logging.getLogger("excel-mcp")

mcp = MCPServer("excel")

# ---------------------------------------------------------------------------
# 定数・ツールの性質
# ---------------------------------------------------------------------------

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
# 設定すると、このディレクトリの外にあるファイルは操作できない
ALLOWED_ROOT = os.environ.get("EXCEL_MCP_ROOT")

READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
DESTRUCTIVE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)

# ---------------------------------------------------------------------------
# 引数の型（説明と例は JSON Schema に載り、LLM に渡る）
# ---------------------------------------------------------------------------

PathArg = Annotated[
    str,
    Field(description="ブックの絶対パス（.xlsx / .xlsm）", examples=["/path/to/book.xlsx"]),
]
SheetArg = Annotated[
    str | None,
    Field(description="シート名。省略するとアクティブシート", examples=["Sheet1"]),
]
RangeArg = Annotated[
    str,
    Field(description="A1 形式のセル範囲", examples=["A1:D20", "B3"]),
]
ForceArg = Annotated[
    bool,
    Field(
        description=(
            "保存すると消える要素（図形、スライサーなど）を含むブックでも書き込む。"
            "通常は false のまま呼び、エラーで警告が返ったらユーザーに確認してから true にする"
        )
    ),
]

# ---------------------------------------------------------------------------
# 共通ヘルパー
# ---------------------------------------------------------------------------


def _resolve_path(path: str, *, must_exist: bool = True) -> Path:
    """パスを絶対パスにし、拡張子と許可ディレクトリを確かめる。"""
    p = Path(path).expanduser().resolve()
    if p.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ToolError(f"対応していない形式です: {p.suffix or '(拡張子なし)'}（.xlsx / .xlsm のみ）")
    if ALLOWED_ROOT:
        root = Path(ALLOWED_ROOT).expanduser().resolve()
        if not p.is_relative_to(root):
            raise ToolError(f"{root} の外にあるファイルは操作できません: {p}")
    if must_exist and not p.is_file():
        raise ToolError(f"ファイルが見つかりません: {p}")
    return p


def _load(p: Path, *, read_only: bool = False, data_only: bool = False) -> Workbook:
    """ブックを開く。.xlsm はマクロを残すため keep_vba=True で開く。"""
    try:
        return xl.load_workbook(
            p,
            read_only=read_only,
            data_only=data_only,
            keep_vba=p.suffix.lower() == ".xlsm",
        )
    except InvalidFileException as e:
        raise ToolError(f"Excel ファイルとして開けません: {e}") from e
    except (zipfile.BadZipFile, KeyError) as e:
        raise ToolError(f"ファイルが壊れているか、.xlsx 形式ではありません: {p}") from e


def _get_sheet(wb: Workbook, sheet: str | None) -> Worksheet:
    if sheet is None:
        return wb.active
    if sheet not in wb.sheetnames:
        raise ToolError(f"シート '{sheet}' がありません。シート: {wb.sheetnames}")
    return wb[sheet]


def _bounds(ws: Worksheet, ref: str) -> tuple[int, int, int, int]:
    """範囲を (min_col, min_row, max_col, max_row) にする。"A:C" や "2:5" も受け付ける。"""
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref.replace("$", "").upper())
    except (ValueError, TypeError) as e:
        raise ToolError(f"範囲の書き方が正しくありません: '{ref}'（例: \"A1:D20\"）") from e
    return (
        min_col or 1,
        min_row or 1,
        max_col or max(ws.max_column, 1),
        max_row or max(ws.max_row, 1),
    )


def _ref(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return start if start == end else f"{start}:{end}"


def _lossy_elements(p: Path) -> list[str]:
    """openpyxl で保存すると消える、または壊れる要素を、ファイルの中身（zip）から探す。"""
    found: set[str] = set()
    with zipfile.ZipFile(p) as z:
        names = z.namelist()
        if any(n.startswith("xl/ctrlProps/") for n in names):
            found.add("フォームコントロール（チェックボックス、ボタン）")
        if any(n.startswith(("xl/slicers/", "xl/slicerCaches/")) for n in names):
            found.add("スライサー")
        if any(n.startswith("xl/timelines/") for n in names):
            found.add("タイムライン")
        if any(n.startswith("xl/threadedComments/") for n in names):
            found.add("スレッド形式のコメント")
        if any(n.startswith("xl/media/") for n in names) and importlib.util.find_spec("PIL") is None:
            found.add("画像（Pillow が未インストールのため）")
        if "xl/vbaProject.bin" in names and p.suffix.lower() != ".xlsm":
            found.add("VBA マクロ（拡張子が .xlsx のため）")
        for n in names:
            if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                if b"<xdr:sp>" in z.read(n) or b"<xdr:sp " in z.read(n):
                    found.add("図形、テキストボックス、矢印")
            elif n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                xml = z.read(n)
                if b"sparklineGroup" in xml:
                    found.add("スパークライン")
                if b"x14:conditionalFormatting" in xml:
                    found.add("新しい形式の条件付き書式（データバー、アイコンセットなど）")
                if b"x14:dataValidation" in xml:
                    found.add("一部の入力規則（拡張形式）")
    return sorted(found)


CHART_NOTICE = (
    "このブックのグラフは保存後も残りますが、Excel で付けた細かい書式（スタイル、色など）は"
    "崩れることがあります"
)


def _notices(p: Path) -> list[str]:
    """保存しても消えないが、崩れることがある要素の注意。書き込みは止めない。"""
    with zipfile.ZipFile(p) as z:
        if any(n.startswith("xl/charts/chart") for n in z.namelist()):
            return [CHART_NOTICE]
    return []


def _lock_file(p: Path) -> Path | None:
    """Excel（~$）や LibreOffice（.~lock.#）がファイルを開いていると作るロックファイル。"""
    for name in (f"~${p.name}", f".~lock.{p.name}#"):
        lock = p.with_name(name)
        if lock.exists():
            return lock
    return None


def _guard_write(p: Path, force: bool) -> None:
    """書き込む前の確認。開いているファイルと、保存すると消える要素を調べる。"""
    if lock := _lock_file(p):
        raise ToolError(
            f"ファイルが Excel / LibreOffice で開かれています（{lock.name}）。閉じてから再実行してください"
        )
    if not force and (lossy := _lossy_elements(p)):
        raise ToolError(
            "このブックには、保存すると消える要素があります: "
            + "、".join(lossy)
            + "。ユーザーの了承を得てから force=true で再実行してください"
        )


def _save(wb: Workbook, p: Path) -> dict[str, Any]:
    """バックアップ（.bak）を作り、一時ファイルに保存してから置き換える。

    戻り値はツールの結果にそのまま混ぜる: {"backup": ..., "notices": [...]}
    """
    backup = None
    notices: list[str] = []
    if p.exists():
        notices = _notices(p)
        backup = p.with_name(p.name + ".bak")
        shutil.copy2(p, backup)
    fd, tmp = tempfile.mkstemp(dir=p.parent, prefix=f".{p.stem}.", suffix=p.suffix)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, p)
    except Exception as e:
        Path(tmp).unlink(missing_ok=True)
        raise ToolError(f"保存に失敗しました: {e}") from e
    return {"backup": str(backup) if backup else None, "notices": notices}


def _jsonable(v: Any) -> Any:
    """セルの値を JSON で返せる形にする。"""
    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    if isinstance(v, (datetime, date, time)):
        return v.isoformat()
    if hasattr(v, "text"):  # ArrayFormula など
        return v.text
    return str(v)


_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?$")


def _to_cell_value(v: Any, convert_dates: bool) -> Any:
    if convert_dates and isinstance(v, str):
        try:
            if _ISO_DATE.match(v):
                return date.fromisoformat(v)
            if _ISO_DATETIME.match(v):
                return datetime.fromisoformat(v)
        except ValueError:
            pass
    return v


def _merged_lookup(ws: Worksheet) -> dict[tuple[int, int], CellRange]:
    """(行, 列) → そのセルを含む結合範囲。"""
    lookup: dict[tuple[int, int], CellRange] = {}
    for rng in ws.merged_cells.ranges:
        for cell in rng.cells:
            lookup[cell] = rng
    return lookup


def _find_table(wb: Workbook, name: str) -> tuple[Worksheet, Table]:
    for ws in wb.worksheets:
        for table in ws.tables.values():  # tables.items() は (名前, 範囲の文字列) を返すので使わない
            if table.displayName.lower() == name.lower():
                return ws, table
    all_names = [t for ws in wb.worksheets for t in ws.tables]
    raise ToolError(f"テーブル '{name}' がありません。テーブル: {all_names}")


def _color(c: str) -> str:
    c = c.lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}([0-9A-F]{2})?", c):
        raise ToolError(f"色は 16 進数の RGB で指定してください（例: \"FFFF00\"）: {c}")
    return c


def _overlaps(a: str, b: str) -> bool:
    return not CellRange(a).isdisjoint(CellRange(b))


# ===========================================================================
# 1. ブック・シート管理
# ===========================================================================


@mcp.tool(annotations=READ_ONLY)
def get_workbook_info(path: PathArg) -> dict[str, Any]:
    """ブックの構造を返す。どの操作でも最初に呼ぶ。

    返す内容: シートごとの使用範囲、結合範囲、テーブル（名前、範囲、スタイル）、入力規則、
    オートフィルター、ウィンドウ枠の固定、保存すると消える要素（lossy_elements）、
    保存すると崩れることがある要素の注意（notices）。
    lossy_elements が空でなければ、書き込み系ツールの前にユーザーへ伝えること。
    """
    p = _resolve_path(path)
    wb = _load(p)
    try:
        sheets = []
        for ws in wb.worksheets:
            sheets.append(
                {
                    "name": ws.title,
                    "used_range": ws.dimensions,
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                    "merged_ranges": [str(r) for r in ws.merged_cells.ranges],
                    "tables": [
                        {
                            "name": t.displayName,
                            "range": t.ref,
                            "style": t.tableStyleInfo.name if t.tableStyleInfo else None,
                        }
                        for t in ws.tables.values()
                    ],
                    "data_validations": [
                        {"range": str(dv.sqref), "type": dv.type, "formula1": dv.formula1}
                        for dv in ws.data_validations.dataValidation
                    ],
                    "auto_filter": ws.auto_filter.ref,
                    "freeze_panes": ws.freeze_panes,
                    "state": ws.sheet_state,
                }
            )
        return {
            "path": str(p),
            "active_sheet": wb.active.title,
            "sheets": sheets,
            "lossy_elements": _lossy_elements(p),
            "notices": _notices(p),
        }
    finally:
        wb.close()


@mcp.tool(annotations=WRITE)
def create_workbook(
    path: Annotated[str, Field(description="作るブックの絶対パス（.xlsx）", examples=["/path/to/new.xlsx"])],
    sheet_names: Annotated[
        list[str] | None,
        Field(description="作るシート名の一覧。省略すると Sheet1 だけ", examples=[["売上", "マスタ"]]),
    ] = None,
    overwrite: Annotated[bool, Field(description="同じ名前のファイルがあれば上書きする")] = False,
) -> dict[str, Any]:
    """新しい空のブックを作る。親ディレクトリは存在している必要がある。"""
    p = _resolve_path(path, must_exist=False)
    if p.suffix.lower() != ".xlsx":
        raise ToolError("新規作成できるのは .xlsx だけです（.xlsm は VBA を含められないため）")
    if p.exists() and not overwrite:
        raise ToolError(f"ファイルがすでにあります: {p}（上書きするなら overwrite=true）")
    if not p.parent.is_dir():
        raise ToolError(f"ディレクトリがありません: {p.parent}")
    names = sheet_names or ["Sheet1"]
    if len(set(names)) != len(names):
        raise ToolError(f"シート名が重複しています: {names}")

    wb = Workbook()
    wb.active.title = names[0]
    for name in names[1:]:
        wb.create_sheet(name)
    saved = _save(wb, p)
    return {"path": str(p), "sheets": names, **saved}


@mcp.tool(annotations=DESTRUCTIVE)
def manage_sheet(
    path: PathArg,
    action: Annotated[
        Literal["add", "rename", "delete", "copy"],
        Field(description="add: 追加 / rename: 名前の変更 / delete: 削除 / copy: コピー"),
    ],
    sheet: Annotated[
        str | None,
        Field(description="対象のシート名（rename / delete / copy で必須）"),
    ] = None,
    new_name: Annotated[
        str | None,
        Field(description="新しいシート名（add / rename / copy で必須）"),
    ] = None,
    index: Annotated[
        int | None,
        Field(description="add / copy で、シートを置く位置（0 が先頭）。省略すると末尾"),
    ] = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """シートの追加、名前の変更、削除、コピー。

    注意: copy は値、書式、結合セルを複製するが、テーブル、入力規則、グラフ、画像は複製しない。
    rename しても、ほかのシートの数式にある古いシート名は書き換わらない。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    if action in ("rename", "delete", "copy") and not sheet:
        raise ToolError(f"{action} には sheet が必要です")
    if action in ("add", "rename", "copy") and not new_name:
        raise ToolError(f"{action} には new_name が必要です")
    if new_name and (len(new_name) > 31 or re.search(r"[\\/?*\[\]:]", new_name)):
        raise ToolError("シート名は 31 文字以内で、\\ / ? * [ ] : は使えません")

    wb = _load(p)
    try:
        if new_name and new_name in wb.sheetnames:
            raise ToolError(f"シート '{new_name}' はすでにあります")
        if action == "add":
            wb.create_sheet(new_name, index)
        else:
            ws = _get_sheet(wb, sheet)
            if action == "rename":
                ws.title = new_name
            elif action == "delete":
                if len(wb.worksheets) == 1:
                    raise ToolError("最後の 1 枚のシートは削除できません")
                wb.remove(ws)
            elif action == "copy":
                new_ws = wb.copy_worksheet(ws)
                new_ws.title = new_name
                if index is not None:
                    wb.move_sheet(new_ws, offset=index - wb.index(new_ws))
        saved = _save(wb, p)
        return {"sheets": wb.sheetnames, **saved}
    finally:
        wb.close()


# ===========================================================================
# 2. データの読み込み・検索
# ===========================================================================


@mcp.tool(annotations=READ_ONLY)
def read_range(
    path: PathArg,
    cell_range: Annotated[
        str | None,
        Field(description="A1 形式の範囲。省略すると使用範囲全体", examples=["A1:D20", "A:C"]),
    ] = None,
    sheet: SheetArg = None,
    mode: Annotated[
        Literal["formula", "value"],
        Field(
            description=(
                "formula: 数式をそのまま返す / value: 最後に Excel で計算された結果を返す。"
                "openpyxl で書いた数式は、Excel か LibreOffice で開き直すまで value が null になる"
            )
        ),
    ] = "formula",
    merged: Annotated[
        Literal["fill", "mark", "none"],
        Field(description="結合セルの左上以外のセル。fill: 左上の値で埋める / mark: \"<merged:B2>\" と印を付ける / none: null"),
    ] = "fill",
    max_rows: Annotated[int, Field(description="返す行数の上限", ge=1, le=5000)] = 200,
) -> dict[str, Any]:
    """範囲の値を 2 次元配列（行のリスト）で返す。

    行数が max_rows を超えたら truncated=true を返すので、範囲を分けて読み直す。
    日付は ISO 8601 の文字列（"2024-01-31T00:00:00"）で返す。
    """
    p = _resolve_path(path)
    wb = _load(p, data_only=(mode == "value"))
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _bounds(ws, cell_range or ws.dimensions)
        total_rows = max_row - min_row + 1
        truncated = total_rows > max_rows
        if truncated:
            max_row = min_row + max_rows - 1

        lookup = _merged_lookup(ws) if merged != "none" else {}
        values = []
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            out = []
            for cell in row:
                v = cell.value
                if isinstance(cell, MergedCell) and (rng := lookup.get((cell.row, cell.column))):
                    if merged == "fill":
                        v = ws.cell(rng.min_row, rng.min_col).value
                    else:
                        v = f"<merged:{get_column_letter(rng.min_col)}{rng.min_row}>"
                out.append(_jsonable(v))
            values.append(out)
        return {
            "sheet": ws.title,
            "range": _ref(min_col, min_row, max_col, max_row),
            "values": values,
            "truncated": truncated,
            "total_rows": total_rows,
        }
    finally:
        wb.close()


@mcp.tool(annotations=READ_ONLY)
def find(
    path: PathArg,
    query: Annotated[str, Field(description="探す値や文字列", examples=["売上", "^2024"])],
    sheet: Annotated[str | None, Field(description="シート名。省略するとすべてのシート")] = None,
    match: Annotated[
        Literal["contains", "exact", "regex"],
        Field(description="contains: 部分一致 / exact: 完全一致 / regex: 正規表現"),
    ] = "contains",
    case_sensitive: bool = False,
    mode: Annotated[
        Literal["formula", "value"],
        Field(description="formula: 数式の文字列も検索する / value: 計算結果を検索する"),
    ] = "formula",
    max_results: Annotated[int, Field(ge=1, le=1000)] = 100,
) -> dict[str, Any]:
    """値や文字列を検索して、見つかったセルの番地と値の一覧を返す。数値も文字列として比べる。"""
    p = _resolve_path(path)
    flags = 0 if case_sensitive else re.IGNORECASE
    try:
        pattern = re.compile(
            query if match == "regex" else (f"^{re.escape(query)}$" if match == "exact" else re.escape(query)),
            flags,
        )
    except re.error as e:
        raise ToolError(f"正規表現が正しくありません: {e}") from e

    wb = _load(p, read_only=True, data_only=(mode == "value"))
    try:
        targets = [_get_sheet(wb, sheet)] if sheet else wb.worksheets
        results: list[dict[str, Any]] = []
        truncated = False
        for ws in targets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    if pattern.search(str(_jsonable(cell.value))):
                        if len(results) >= max_results:
                            truncated = True
                            break
                        results.append({"sheet": ws.title, "cell": cell.coordinate, "value": _jsonable(cell.value)})
                if truncated:
                    break
            if truncated:
                break
        return {"matches": results, "count": len(results), "truncated": truncated}
    finally:
        wb.close()


# ===========================================================================
# 3. データの書き込み
# ===========================================================================


@mcp.tool(annotations=DESTRUCTIVE)
def write_range(
    path: PathArg,
    start_cell: Annotated[str, Field(description="書き込みを始める左上のセル", examples=["A1"])],
    values: Annotated[
        list[list[str | int | float | bool | None]],
        Field(
            description="行のリスト。\"=\" で始まる文字列は数式になる。null はセルを空にする",
            examples=[[["品名", "数量", "金額"], ["りんご", 3, "=B2*100"]]],
        ),
    ],
    sheet: SheetArg = None,
    convert_dates: Annotated[
        bool,
        Field(description="\"2024-01-31\" や \"2024-01-31T09:00\" の形の文字列を日付として書き込む"),
    ] = True,
    force: ForceArg = False,
) -> dict[str, Any]:
    """値と数式を書き込む。既存の値は上書きする。

    結合セルの左上以外に書くと、左上のセルに振り替える。
    同じ結合範囲に 2 つ以上の値が入るときは、後の値で上書きして warnings で知らせる。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    if not values or not any(values):
        raise ToolError("values が空です")

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        try:
            col_letter, row0 = coordinate_from_string(start_cell.replace("$", "").upper())
        except ValueError as e:
            raise ToolError(f"セルの番地が正しくありません: {start_cell}") from e
        col0 = column_index_from_string(col_letter)

        lookup = _merged_lookup(ws)
        filled: dict[str, str] = {}  # 結合範囲 → 最初に値を入れた元のセル
        warnings: list[str] = []
        count = 0
        width = max(len(r) for r in values)
        for i, row in enumerate(values):
            for j, v in enumerate(row):
                r, c = row0 + i, col0 + j
                origin = f"{get_column_letter(c)}{r}"
                if rng := lookup.get((r, c)):
                    key = str(rng)
                    if v is not None:
                        if key in filled:
                            warnings.append(f"結合範囲 {key} に {filled[key]} と {origin} の値が入りました。{origin} の値を残します")
                        filled[key] = origin
                    elif (r, c) != (rng.min_row, rng.min_col):
                        continue  # 左上以外への null は無視する（左上の値を消さない）
                    r, c = rng.min_row, rng.min_col
                ws.cell(row=r, column=c, value=_to_cell_value(v, convert_dates))
                count += 1
        saved = _save(wb, p)
        return {
            "sheet": ws.title,
            "range": _ref(col0, row0, col0 + width - 1, row0 + len(values) - 1),
            "cells_written": count,
            "warnings": warnings,
            **saved,
        }
    finally:
        wb.close()


def _last_row(ws: Worksheet, min_col: int, max_col: int) -> int:
    """指定した列の範囲で、値が入っている最後の行（書式だけの行は数えない）。"""
    for r in range(ws.max_row, 0, -1):
        for c in range(min_col, max_col + 1):
            if ws.cell(r, c).value is not None:
                return r
    return 0


@mcp.tool(annotations=WRITE)
def append_rows(
    path: PathArg,
    rows: Annotated[
        list[list[str | int | float | bool | None]],
        Field(description="追記する行のリスト", examples=[[["みかん", 5, "=B3*80"]]]),
    ],
    sheet: SheetArg = None,
    table: Annotated[
        str | None,
        Field(description="テーブル名。指定するとそのテーブルの直下に追記し、テーブルの範囲を広げる"),
    ] = None,
    convert_dates: bool = True,
    force: ForceArg = False,
) -> dict[str, Any]:
    """最終行の下に行を追記する。

    table を省略した場合でも、最終行がテーブルの最終行と同じならそのテーブルの範囲を広げる。
    テーブルの直下の書き込む行に別のデータがあるときは、上書きを防ぐためエラーを返す。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    if not rows:
        raise ToolError("rows が空です")

    wb = _load(p)
    try:
        warnings: list[str] = []
        target: Table | None = None
        if table:
            ws, target = _find_table(wb, table)
            if sheet and ws.title != sheet:
                raise ToolError(f"テーブル '{table}' はシート '{ws.title}' にあります")
            min_col, _, max_col, t_max_row = range_boundaries(target.ref)
            start_row = t_max_row + 1
            width = max(len(r) for r in rows)
            for r in range(start_row, start_row + len(rows)):
                for c in range(min_col, min_col + max(width, max_col - min_col + 1)):
                    if ws.cell(r, c).value is not None:
                        raise ToolError(
                            f"テーブル '{table}' の直下（{ws.cell(r, c).coordinate}）にデータがあるため追記できません。"
                            "insert_delete で行を挿入してから追記してください"
                        )
        else:
            ws = _get_sheet(wb, sheet)
            min_col = 1
            start_row = _last_row(ws, 1, max(ws.max_column, 1)) + 1
            for t in ws.tables.values():
                t_min_col, _, t_max_col, t_max_row = range_boundaries(t.ref)
                if t_max_row == start_row - 1:
                    target, min_col = t, t_min_col
                    break

        width = max(len(r) for r in rows)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                ws.cell(row=start_row + i, column=min_col + j, value=_to_cell_value(v, convert_dates))
        end_row = start_row + len(rows) - 1

        if target is not None:
            t_min_col, t_min_row, t_max_col, _ = range_boundaries(target.ref)
            if width > t_max_col - t_min_col + 1:
                warnings.append(f"テーブル '{target.displayName}' の列数より多い値はテーブルの外に書きました")
            target.ref = _ref(t_min_col, t_min_row, t_max_col, end_row)
            if target.autoFilter is not None:
                target.autoFilter.ref = target.ref

        saved = _save(wb, p)
        return {
            "sheet": ws.title,
            "range": _ref(min_col, start_row, min_col + width - 1, end_row),
            "table": target.displayName if target is not None else None,
            "table_range": target.ref if target is not None else None,
            "warnings": warnings,
            **saved,
        }
    finally:
        wb.close()


# ===========================================================================
# 4. 表の構造の変更
# ===========================================================================


def _shift_span(lo: int, hi: int, idx: int, amount: int, insert: bool) -> tuple[int, int] | None:
    """行（列）を挿入・削除したあとの [lo, hi] を返す。すべて消えたら None。"""
    if insert:
        if lo >= idx:
            return lo + amount, hi + amount
        if hi >= idx:
            return lo, hi + amount
        return lo, hi
    end = idx + amount - 1
    if hi < idx:
        return lo, hi
    if lo > end:
        return lo - amount, hi - amount
    removed = min(hi, end) - max(lo, idx) + 1
    new_lo = lo if lo < idx else idx
    new_hi = hi - removed
    return (new_lo, new_hi) if new_hi >= new_lo else None


@mcp.tool(annotations=DESTRUCTIVE)
def insert_delete(
    path: PathArg,
    action: Annotated[
        Literal["insert_rows", "delete_rows", "insert_cols", "delete_cols"],
        Field(description="行・列の挿入または削除"),
    ],
    index: Annotated[
        int | str,
        Field(description="挿入・削除を始める行番号（1 始まり）または列（\"C\" か 3）", examples=[5, "C"]),
    ],
    amount: Annotated[int, Field(description="挿入・削除する行数・列数", ge=1)] = 1,
    sheet: SheetArg = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """行・列を挿入または削除する。挿入は index の位置の前に入る。

    結合範囲とテーブルの範囲はずらす。
    数式の参照、条件付き書式、入力規則、シートのオートフィルター、グラフの参照はずれない。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    is_row = action.endswith("rows")
    insert = action.startswith("insert")
    if isinstance(index, str):
        index = int(index) if index.isdigit() else column_index_from_string(index.upper())
    if index < 1:
        raise ToolError("index は 1 以上です")

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        warnings = [
            "数式の参照、条件付き書式、入力規則、シートのオートフィルターはずれていません。必要なら確認してください"
        ]

        # 結合を一度外し、移動後にずらした位置で結合し直す
        merged = [CellRange(str(r)) for r in ws.merged_cells.ranges]
        for rng in merged:
            ws.unmerge_cells(str(rng))

        if is_row:
            (ws.insert_rows if insert else ws.delete_rows)(index, amount)
        else:
            (ws.insert_cols if insert else ws.delete_cols)(index, amount)

        for rng in merged:
            if is_row:
                span = _shift_span(rng.min_row, rng.max_row, index, amount, insert)
                new = None if span is None else (rng.min_col, span[0], rng.max_col, span[1])
            else:
                span = _shift_span(rng.min_col, rng.max_col, index, amount, insert)
                new = None if span is None else (span[0], rng.min_row, span[1], rng.max_row)
            if new is None or (new[0] == new[2] and new[1] == new[3]):
                warnings.append(f"結合範囲 {rng} は削除によりなくなりました")
                continue
            ws.merge_cells(_ref(*new))

        for t in list(ws.tables.values()):
            name = t.displayName
            min_col, min_row, max_col, max_row = range_boundaries(t.ref)
            if is_row:
                span = _shift_span(min_row, max_row, index, amount, insert)
                new = None if span is None else (min_col, span[0], max_col, span[1])
                header_deleted = not insert and index <= min_row < index + amount
            else:
                span = _shift_span(min_col, max_col, index, amount, insert)
                new = None if span is None else (span[0], min_row, span[1], max_row)
                header_deleted = False
                if span and (span[1] - span[0]) != (max_col - min_col):
                    warnings.append(f"テーブル '{name}' の列数が変わりました。列名を確認してください")
                    t.tableColumns = []
                    t._initialise_columns()
            if new is None or header_deleted or new[1] == new[3]:
                del ws.tables[name]
                warnings.append(f"テーブル '{name}' は見出し行かデータ行がなくなったため削除しました")
                continue
            t.ref = _ref(*new)
            if t.autoFilter is not None:
                t.autoFilter.ref = t.ref

        saved = _save(wb, p)
        return {"sheet": ws.title, "used_range": ws.dimensions, "warnings": warnings, **saved}
    finally:
        wb.close()


# ===========================================================================
# 5. 書式・レイアウト
# ===========================================================================


@mcp.tool(annotations=WRITE)
def format_range(
    path: PathArg,
    cell_range: RangeArg,
    sheet: SheetArg = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: Annotated[float | None, Field(gt=0, examples=[11])] = None,
    font_color: Annotated[str | None, Field(description="文字の色（RGB）", examples=["FF0000"])] = None,
    fill_color: Annotated[
        str | None,
        Field(description="背景色（RGB）。\"none\" で背景色を消す", examples=["FFFF00", "none"]),
    ] = None,
    number_format: Annotated[
        str | None,
        Field(description="表示形式", examples=["#,##0", "yyyy/mm/dd", "0.0%", "General"]),
    ] = None,
    border: Annotated[
        Literal["thin", "medium", "thick", "dashed", "double", "none"] | None,
        Field(description="範囲のすべてのセルの上下左右に付ける罫線。none で罫線を消す"),
    ] = None,
    horizontal: Literal["left", "center", "right", "general"] | None = None,
    vertical: Literal["top", "center", "bottom"] | None = None,
    wrap_text: bool | None = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """セルの書式を設定する。指定した項目だけを変え、ほかの書式は残す。"""
    p = _resolve_path(path)
    _guard_write(p, force)
    font_color = _color(font_color) if font_color else None
    fill = None
    if fill_color is not None:
        fill = PatternFill(fill_type=None) if fill_color.lower() == "none" else PatternFill(
            "solid", start_color=_color(fill_color), end_color=_color(fill_color)
        )
    side = None if border in (None, "none") else Side(style=border)

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _bounds(ws, cell_range)
        count = 0
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                if any(v is not None for v in (bold, italic, font_size, font_color)):
                    f: Font = copy(cell.font)
                    if bold is not None:
                        f.bold = bold
                    if italic is not None:
                        f.italic = italic
                    if font_size is not None:
                        f.size = font_size
                    if font_color is not None:
                        f.color = font_color
                    cell.font = f
                if fill is not None:
                    cell.fill = copy(fill)
                if number_format is not None:
                    cell.number_format = number_format
                if border is not None:
                    cell.border = Border() if side is None else Border(left=side, right=side, top=side, bottom=side)
                if any(v is not None for v in (horizontal, vertical, wrap_text)):
                    a: Alignment = copy(cell.alignment)
                    if horizontal is not None:
                        a.horizontal = horizontal
                    if vertical is not None:
                        a.vertical = vertical
                    if wrap_text is not None:
                        a.wrap_text = wrap_text
                    cell.alignment = a
                count += 1
        saved = _save(wb, p)
        return {"sheet": ws.title, "range": _ref(min_col, min_row, max_col, max_row), "cells": count, **saved}
    finally:
        wb.close()


@mcp.tool(annotations=WRITE)
def set_layout(
    path: PathArg,
    sheet: SheetArg = None,
    column_widths: Annotated[
        dict[str, float] | None,
        Field(description="列 → 幅（文字数）", examples=[{"A": 12, "B": 30}]),
    ] = None,
    row_heights: Annotated[
        dict[str, float] | None,
        Field(description="行番号 → 高さ（ポイント）", examples=[{"1": 24}]),
    ] = None,
    freeze_panes: Annotated[
        str | None,
        Field(description="このセルの左上で固定する（\"B2\" なら 1 行目と A 列）。\"none\" で解除", examples=["A2", "B2", "none"]),
    ] = None,
    auto_filter: Annotated[
        str | None,
        Field(description="シートのオートフィルターの範囲。\"none\" で解除。テーブルには filter_table を使う", examples=["A1:F100", "none"]),
    ] = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """列幅、行の高さ、ウィンドウ枠の固定、シートのオートフィルターを設定する。"""
    p = _resolve_path(path)
    _guard_write(p, force)

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        for col, width in (column_widths or {}).items():
            letter = get_column_letter(int(col)) if str(col).isdigit() else col.upper()
            ws.column_dimensions[letter].width = width
        for row, height in (row_heights or {}).items():
            ws.row_dimensions[int(row)].height = height
        if freeze_panes is not None:
            ws.freeze_panes = None if freeze_panes.lower() == "none" else freeze_panes.upper()
        if auto_filter is not None:
            if auto_filter.lower() == "none":
                ws.auto_filter.ref = None
                ws.auto_filter.filterColumn = []
            else:
                min_col, min_row, max_col, max_row = _bounds(ws, auto_filter)
                ref = _ref(min_col, min_row, max_col, max_row)
                for t in ws.tables.values():
                    if _overlaps(ref, t.ref):
                        raise ToolError(f"テーブル '{t.displayName}'（{t.ref}）と重なるため、シートのオートフィルターは付けられません")
                ws.auto_filter.ref = ref
        saved = _save(wb, p)
        return {
            "sheet": ws.title,
            "freeze_panes": ws.freeze_panes,
            "auto_filter": ws.auto_filter.ref,
            **saved,
        }
    finally:
        wb.close()


# ===========================================================================
# 6. テーブル・データの管理
# ===========================================================================

_TABLE_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$")
_CELL_LIKE = re.compile(r"^([A-Za-z]{1,3}\d+|[RrCc]\d*)$")


@mcp.tool(annotations=WRITE)
def create_table(
    path: PathArg,
    cell_range: Annotated[str, Field(description="テーブルにする範囲。1 行目が見出しになる", examples=["A1:D20"])],
    name: Annotated[
        str,
        Field(description="テーブル名。英字か _ で始め、英数字・_・. だけ。ブック内で一意", examples=["SalesTable"]),
    ],
    style: Annotated[
        str,
        Field(description="テーブルスタイル", examples=["TableStyleMedium2", "TableStyleLight9"]),
    ] = "TableStyleMedium2",
    sheet: SheetArg = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """範囲をテーブルにする。見出しの空欄や重複、範囲内の結合セル、テーブル名の重複はエラーにする。

    見出しが数値のときは文字列に変換する（Excel の仕様）。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    if not _TABLE_NAME.match(name) or _CELL_LIKE.match(name):
        raise ToolError(f"テーブル名が使えません: '{name}'（英字か _ で始め、セル番地と同じ形は不可）")

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _bounds(ws, cell_range)
        ref = _ref(min_col, min_row, max_col, max_row)
        if max_row <= min_row:
            raise ToolError("テーブルには見出し行とデータ行が 1 行以上必要です")

        existing = {t.lower() for w in wb.worksheets for t in w.tables} | {n.lower() for n in wb.defined_names}
        if name.lower() in existing:
            raise ToolError(f"テーブル名か名前 '{name}' はすでに使われています")
        for t in ws.tables.values():
            if _overlaps(ref, t.ref):
                raise ToolError(f"既存のテーブル '{t.displayName}'（{t.ref}）と重なっています")
        for m in ws.merged_cells.ranges:
            if _overlaps(ref, str(m)):
                raise ToolError(f"範囲内に結合セル {m} があります。結合を解除してから作ってください")
        if ws.auto_filter.ref and _overlaps(ref, ws.auto_filter.ref):
            raise ToolError("シートのオートフィルターと重なっています。set_layout で解除してください")

        headers = []
        for c in range(min_col, max_col + 1):
            cell = ws.cell(min_row, c)
            if cell.value is None or str(cell.value).strip() == "":
                raise ToolError(f"見出しが空のセルがあります: {cell.coordinate}")
            if not isinstance(cell.value, str):
                cell.value = str(_jsonable(cell.value))
            headers.append(cell.value)
        if len({h.lower() for h in headers}) != len(headers):
            raise ToolError(f"見出しが重複しています: {headers}")

        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
        ws.add_table(table)
        saved = _save(wb, p)
        return {"sheet": ws.title, "table": name, "range": ref, "headers": headers, **saved}
    finally:
        wb.close()


def _display(v: Any) -> str:
    """フィルターで比べるための、セルの値の文字列表現。"""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime):
        return v.date().isoformat() if v.time() == time() else v.isoformat(sep=" ")
    return "" if v is None else str(_jsonable(v))


def _num(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (datetime, date)):
        return None
    try:
        return float(str(v))
    except ValueError:
        return None


def _custom_match(cf: CustomFilter, v: Any) -> bool:
    op, val = cf.operator or "equal", cf.val or ""
    if op in ("equal", "notEqual") and ("*" in val or "?" in val):
        hit = fnmatch.fnmatch(_display(v).casefold(), val.casefold())
        return hit if op == "equal" else not hit
    a, b = _num(v), _num(val)
    if a is None or b is None:
        a, b = _display(v).casefold(), val.casefold()  # type: ignore[assignment]
    return {
        "equal": a == b,
        "notEqual": a != b,
        "greaterThan": a > b,
        "greaterThanOrEqual": a >= b,
        "lessThan": a < b,
        "lessThanOrEqual": a <= b,
    }[op]


def _passes(fc: FilterColumn, v: Any) -> bool:
    if fc.filters is not None:
        if v is None or _display(v) == "":
            return bool(fc.filters.blank)
        return _display(v) in set(fc.filters.filter)
    if fc.customFilters is not None:
        results = [_custom_match(cf, v) for cf in fc.customFilters.customFilter]
        return all(results) if fc.customFilters._and else any(results)
    return True


@mcp.tool(annotations=WRITE)
def filter_table(
    path: PathArg,
    table: Annotated[str, Field(description="テーブル名", examples=["SalesTable"])],
    condition: Annotated[
        Literal["values", "gte", "lte", "between", "contains", "clear"],
        Field(
            description=(
                "values: 値の一覧のどれかに一致 / gte: 以上 / lte: 以下 / between: values[0] 以上 values[1] 以下 / "
                "contains: 文字列を含む / clear: 解除（column を省略するとすべての列）"
            )
        ),
    ],
    column: Annotated[
        str | None,
        Field(description="条件を付ける列の見出し名（clear 以外では必須）", examples=["地域"]),
    ] = None,
    values: Annotated[
        list[str | int | float] | None,
        Field(description="条件の値。values は一覧、gte / lte / contains は 1 つ、between は 2 つ", examples=[["東京", "大阪"], [1000]]),
    ] = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """テーブルの列に条件を付けて絞り込み、条件に合わない行を非表示にする。

    ほかの列にすでに付いている条件は残り、すべての条件を満たす行だけが表示される。
    判定には、最後に Excel で計算された値を使う（openpyxl で書いた数式は値がないので空として扱う）。
    """
    p = _resolve_path(path)
    _guard_write(p, force)
    vals = [str(v) for v in (values or [])]
    need = {"values": (1, None), "gte": (1, 1), "lte": (1, 1), "between": (2, 2), "contains": (1, 1)}
    if condition != "clear":
        if column is None:
            raise ToolError("column が必要です")
        lo, hi = need[condition]
        if len(vals) < lo or (hi is not None and len(vals) > hi):
            raise ToolError(f"{condition} には values を {lo}{'' if hi == lo else ' 個以上'}個指定してください")

    wb = _load(p)
    cache = _load(p, read_only=True, data_only=True)
    try:
        ws, t = _find_table(wb, table)
        min_col, min_row, max_col, max_row = range_boundaries(t.ref)
        headers = [str(ws.cell(min_row, c).value) for c in range(min_col, max_col + 1)]
        if t.autoFilter is None:
            t.autoFilter = AutoFilter(ref=t.ref)

        col_id = None
        if column is not None:
            if column not in headers:
                raise ToolError(f"列 '{column}' がありません。見出し: {headers}")
            col_id = headers.index(column)

        # 同じ列の古い条件は置き換える。clear で column を省略したら全部消す
        if condition == "clear" and col_id is None:
            columns = []
        else:
            columns = [fc for fc in t.autoFilter.filterColumn if fc.colId != col_id]
        if condition != "clear":
            if condition == "values":
                fc = FilterColumn(colId=col_id, filters=Filters(filter=vals))
            elif condition == "contains":
                fc = FilterColumn(colId=col_id, customFilters=CustomFilters(customFilter=[CustomFilter("equal", f"*{vals[0]}*")]))
            elif condition == "gte":
                fc = FilterColumn(colId=col_id, customFilters=CustomFilters(customFilter=[CustomFilter("greaterThanOrEqual", vals[0])]))
            elif condition == "lte":
                fc = FilterColumn(colId=col_id, customFilters=CustomFilters(customFilter=[CustomFilter("lessThanOrEqual", vals[0])]))
            else:
                fc = FilterColumn(
                    colId=col_id,
                    customFilters=CustomFilters(
                        _and=True,
                        customFilter=[CustomFilter("greaterThanOrEqual", vals[0]), CustomFilter("lessThanOrEqual", vals[1])],
                    ),
                )
            columns.append(fc)
        t.autoFilter.filterColumn = sorted(columns, key=lambda f: f.colId)

        cws = cache[ws.title]
        visible = hidden = 0
        rows = cws.iter_rows(min_row=min_row + 1, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True)
        for r, row in enumerate(rows, start=min_row + 1):
            row = tuple(row) + (None,) * (len(headers) - len(row))
            ok = all(_passes(fc, row[fc.colId]) for fc in t.autoFilter.filterColumn)
            ws.row_dimensions[r].hidden = not ok
            visible, hidden = visible + ok, hidden + (not ok)

        saved = _save(wb, p)
        return {
            "table": t.displayName,
            "filters": [{"column": headers[fc.colId]} for fc in t.autoFilter.filterColumn],
            "visible_rows": visible,
            "hidden_rows": hidden,
            **saved,
        }
    finally:
        cache.close()
        wb.close()


def _validation_value(v: str | int | float, kind: str) -> str:
    if kind == "date" and isinstance(v, str) and _ISO_DATE.match(v):
        d = date.fromisoformat(v)
        return f"DATE({d.year},{d.month},{d.day})"
    return str(v)


@mcp.tool(annotations=WRITE)
def set_validation(
    path: PathArg,
    cell_range: RangeArg,
    type: Annotated[
        Literal["list", "whole", "decimal", "date", "textLength", "clear"],
        Field(description="list: プルダウン / whole: 整数 / decimal: 小数 / date: 日付 / textLength: 文字数 / clear: 範囲と重なる入力規則を削除"),
    ],
    sheet: SheetArg = None,
    items: Annotated[
        list[str] | None,
        Field(description="list で直接入力する選択肢（合計 255 文字以内、カンマは使えない）", examples=[["未着手", "進行中", "完了"]]),
    ] = None,
    source: Annotated[
        str | None,
        Field(description="list で選択肢を参照する範囲。別シートは 'シート名'!$A$2:$A$20 と書く", examples=["'マスタ'!$A$2:$A$20"]),
    ] = None,
    minimum: Annotated[
        str | int | float | None,
        Field(description="whole / decimal / date / textLength の最小値。日付は \"2024-01-01\"", examples=[0, "2024-01-01"]),
    ] = None,
    maximum: Annotated[str | int | float | None, Field(description="最大値")] = None,
    allow_blank: bool = True,
    error_title: str | None = None,
    error_message: Annotated[str | None, Field(examples=["一覧から選んでください"])] = None,
    prompt: Annotated[str | None, Field(description="セルを選んだときに出す説明")] = None,
    force: ForceArg = False,
) -> dict[str, Any]:
    """入力規則を付ける（プルダウンリスト、数値・日付・文字数の範囲、エラーメッセージ）。"""
    p = _resolve_path(path)
    _guard_write(p, force)

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _bounds(ws, cell_range)
        ref = _ref(min_col, min_row, max_col, max_row)

        if type == "clear":
            removed = []
            for dv in list(ws.data_validations.dataValidation):
                if any(_overlaps(ref, str(r)) for r in dv.sqref.ranges):
                    removed.append(str(dv.sqref))
                    ws.data_validations.dataValidation.remove(dv)
            saved = _save(wb, p)
            return {"sheet": ws.title, "removed": removed, **saved}

        operator = None
        formula2 = None
        if type == "list":
            if bool(items) == bool(source):
                raise ToolError("list には items か source のどちらか一方を指定してください")
            if items:
                if any("," in i or '"' in i for i in items):
                    raise ToolError("items にカンマやダブルクォートは使えません。source で範囲を参照してください")
                formula1 = '"' + ",".join(items) + '"'
                if len(formula1) > 257:
                    raise ToolError("items が長すぎます（合計 255 文字まで）。source で範囲を参照してください")
            else:
                formula1 = source.lstrip("=")
        else:
            if minimum is None and maximum is None:
                raise ToolError(f"{type} には minimum か maximum が必要です")
            if minimum is not None and maximum is not None:
                operator = "between"
                formula1 = _validation_value(minimum, type)
                formula2 = _validation_value(maximum, type)
            elif minimum is not None:
                operator, formula1 = "greaterThanOrEqual", _validation_value(minimum, type)
            else:
                operator, formula1 = "lessThanOrEqual", _validation_value(maximum, type)

        dv = DataValidation(
            type=type,
            operator=operator,
            formula1=formula1,
            formula2=formula2,
            allow_blank=allow_blank,
            showErrorMessage=True,
            errorStyle="stop",
            errorTitle=error_title,
            error=error_message,
            showInputMessage=prompt is not None,
            prompt=prompt,
        )
        dv.add(ref)
        ws.add_data_validation(dv)
        saved = _save(wb, p)
        return {"sheet": ws.title, "range": ref, "type": type, "formula1": formula1, **saved}
    finally:
        wb.close()


# ===========================================================================
# 7. グラフ
# ===========================================================================


@mcp.tool(annotations=WRITE)
def create_chart(
    path: PathArg,
    chart_type: Annotated[
        Literal["column", "bar", "line"],
        Field(description="column: 縦棒 / bar: 横棒 / line: 折れ線"),
    ],
    data_range: Annotated[
        str,
        Field(description="系列のデータ範囲（1 列が 1 系列）。titles_from_data=true なら 1 行目は系列名", examples=["B1:C13"]),
    ],
    categories_range: Annotated[
        str | None,
        Field(description="項目（横軸）の範囲。見出し行は含めない", examples=["A2:A13"]),
    ] = None,
    anchor: Annotated[str, Field(description="グラフの左上を置くセル", examples=["E2"])] = "E2",
    sheet: Annotated[str | None, Field(description="グラフを置くシート。省略するとアクティブシート")] = None,
    data_sheet: Annotated[str | None, Field(description="データがあるシート。省略すると sheet と同じ")] = None,
    titles_from_data: Annotated[bool, Field(description="data_range の 1 行目を系列名として使う")] = True,
    series_names: Annotated[
        list[str] | None,
        Field(description="系列名を上書きする（左の列から順に）", examples=[["2023年", "2024年"]]),
    ] = None,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    width_cm: Annotated[float, Field(gt=0)] = 15,
    height_cm: Annotated[float, Field(gt=0)] = 7.5,
    force: ForceArg = False,
) -> dict[str, Any]:
    """棒グラフ・折れ線グラフを作る。

    作ったグラフは、このあと同じブックを書き込み系ツールで保存しても残る。
    ただし Excel で付けた細かい書式（スタイル、色など）は、保存し直すと崩れることがある。
    """
    p = _resolve_path(path)
    _guard_write(p, force)

    wb = _load(p)
    try:
        ws = _get_sheet(wb, sheet)
        dws = _get_sheet(wb, data_sheet) if data_sheet else ws
        min_col, min_row, max_col, max_row = _bounds(dws, data_range)
        data = Reference(dws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)

        chart = LineChart() if chart_type == "line" else BarChart()
        if chart_type != "line":
            chart.type = "col" if chart_type == "column" else "bar"
        chart.add_data(data, titles_from_data=titles_from_data)
        if categories_range:
            c_min_col, c_min_row, c_max_col, c_max_row = _bounds(dws, categories_range)
            chart.set_categories(
                Reference(dws, min_col=c_min_col, min_row=c_min_row, max_col=c_max_col, max_row=c_max_row)
            )
        if series_names:
            for s, n in zip(chart.series, series_names):
                s.tx = SeriesLabel(v=n)
        chart.title = title
        chart.x_axis.title = x_axis_title
        chart.y_axis.title = y_axis_title
        # openpyxl 3.1 では既定で軸が非表示になるため、明示的に表示する
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.width, chart.height = width_cm, height_cm
        ws.add_chart(chart, anchor.upper())

        saved = _save(wb, p)
        return {
            "sheet": ws.title,
            "anchor": anchor.upper(),
            "series": len(chart.series),
            **saved,
        }
    finally:
        wb.close()


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
