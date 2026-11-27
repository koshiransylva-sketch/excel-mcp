# excel-mcp

An [MCP](https://modelcontextprotocol.io/) server for reading and editing Excel files (`.xlsx` / `.xlsm`).
Register it with an MCP client such as Claude Code or Claude Desktop, then just ask things like "read this sheet and total it by region" or "add a chart to the sales sheet". The AI works on your Excel files directly.

You don't need Excel installed. The server reads and writes files with [openpyxl](https://openpyxl.readthedocs.io/), so it runs on Linux, macOS, and Windows.

## Features

| Category | Tool | What it does |
| --- | --- | --- |
| Workbooks & sheets | `get_workbook_info` | Shows the workbook's structure: sheets, used ranges, merged cells, tables, data validation, and more |
| | `create_workbook` | Creates a new workbook |
| | `manage_sheet` | Adds, renames, deletes, or copies sheets |
| Reading | `read_range` | Reads values from a range (choose formulas or calculated results) |
| | `find` | Searches for values or text (contains, exact match, or regex) |
| Writing | `write_range` | Writes values and formulas (strings starting with `=` become formulas) |
| | `append_rows` | Appends rows below the last row (and extends the table, if there is one) |
| | `insert_delete` | Inserts or deletes rows and columns |
| Formatting & layout | `format_range` | Bold, italic, font size and color, fill color, number format, borders, alignment, text wrap |
| | `set_layout` | Column widths, row heights, freeze panes, AutoFilter |
| Tables | `create_table` | Turns a range into a styled table |
| | `filter_table` | Filters a table column by a condition |
| Data validation | `set_validation` | Drop-down lists and limits on numbers, dates, or text length |
| Charts | `create_chart` | Creates column, bar, or line charts |

### Keeping your files safe

- **Backups**: Before writing, the server copies the original file to `filename.xlsx.bak`.
- **Safe saves**: It saves to a temporary file first and then replaces the original, so a failed save leaves your file untouched.
- **Open files are protected**: Writes are refused while the file is open in Excel or LibreOffice.
- **Warns about lost content**: Some things are lost when openpyxl saves a workbook (shapes, slicers, form controls, sparklines, and so on). If a workbook has any of these, the server stops and tells you. It writes only when you agree and pass `force=true`.
- **Macros are kept**: `.xlsm` files are saved with their macros (VBA) intact.

### Limitations

- openpyxl does not calculate formulas. Formulas written by this server have no calculated value until you open and save the file in Excel or LibreOffice.
- Inserting or deleting rows and columns does not update formula references (unlike Excel).
- The old `.xls` format is not supported.
- Tool descriptions and messages are currently in Japanese. You can still make requests in any language.

## Installation

### 1. Requirements

- [uv](https://docs.astral.sh/uv/) (a Python package manager)

The server needs Python 3.15 or later. uv downloads it for you, so you don't need to install Python yourself.

If you don't have uv yet, install it:

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 2. Download

```bash
git clone https://github.com/koshiransylva-sketch/excel-mcp.git
cd excel-mcp
uv sync
```

`uv sync` installs the dependencies into `.venv`.

You'll need the **absolute path** of the cloned folder in the next step. Check it with `pwd` (or `cd` on Windows). The examples below use `/path/to/excel-mcp`.

### 3. Register the server with your MCP client

#### Claude Code (recommended: `claude mcp add`)

Run this command once, and the server is available in every folder where you open Claude Code:

```bash
claude mcp add excel --scope user \
  -e EXCEL_MCP_ROOT=~/Documents \
  -- uv --directory /path/to/excel-mcp run server.py
```

- `--scope user` registers the server for all of your projects (it is saved in `~/.claude.json`).
- `-e EXCEL_MCP_ROOT=...` sets the folder the server is allowed to work in (see [below](#restricting-which-folders-can-be-used)). It is optional.

To check that it worked, run `claude mcp list`, or type `/mcp` inside Claude Code.

<details>
<summary>Using it in one project only (<code>.mcp.json</code>)</summary>

Create a `.mcp.json` file in the project folder where you work with Excel files:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["--directory", "/path/to/excel-mcp", "run", "server.py"],
      "env": {
        "EXCEL_MCP_ROOT": "/path/to/your/project"
      }
    }
  }
}
```

If you commit this file to Git, others who use the repository get the same setup. They will need to change the paths to match their own machines.

The first time you open Claude Code in that folder, it asks whether to use this server. Allow it.

</details>

#### Claude Desktop

Add the server to Claude Desktop's config file, `claude_desktop_config.json`. You can open it from Settings → Developer → Edit Config.

| OS | Location |
| --- | --- |
| macOS | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Windows | `%APPDATA%\Claude\claude_desktop_config.json` |

```json
{
  "mcpServers": {
    "excel": {
      "command": "/full/path/to/uv",
      "args": ["--directory", "/path/to/excel-mcp", "run", "server.py"],
      "env": {
        "EXCEL_MCP_ROOT": "/Users/you/Documents"
      }
    }
  }
}
```

- Claude Desktop may not use your terminal's `PATH`, so set `command` to the absolute path of uv. Find it with `which uv` (macOS / Linux) or `where uv` (Windows).
- On Windows, write each `\` in a path as `\\` (for example, `"C:\\Users\\you\\excel-mcp"`).
- After saving, quit Claude Desktop completely and start it again.

#### Other MCP clients

This is a stdio MCP server. Any client can run it with this command:

```text
uv --directory /path/to/excel-mcp run server.py
```

### 4. Try it out

Once it's registered, just ask your client in plain language:

```text
Look at ~/Documents/sales.xlsx and put the totals for each region on a new sheet.
```

```text
Add a drop-down list with "To do", "In progress", and "Done" to column C of ~/Documents/tasks.xlsx.
```

Always give files as **absolute paths** (`~` works).

## Restricting which folders can be used

If you set the `EXCEL_MCP_ROOT` environment variable to a folder, the server only works on files inside that folder. Any file outside it is refused with an error.

```text
EXCEL_MCP_ROOT=~/Documents/excel
```

If it is not set, there is no restriction: the server can open any `.xlsx` or `.xlsm` file on your computer. Because you are letting an AI work on files on your computer, setting a dedicated working folder is recommended.

## Testing (for developers)

With [MCP Inspector](https://github.com/modelcontextprotocol/inspector), you can call the tools directly from your browser. It needs Node.js.

```bash
cd /path/to/excel-mcp
npx @modelcontextprotocol/inspector uv run server.py
```
