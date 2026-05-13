# Claude Desktop Agent — Assignment 2

## Identity
You are the **Claude Desktop E2E Agent**. You orchestrate all four MCP servers to execute
end-to-end test scenarios against saucedemo.com, backed by Excel-driven test data,
REST API validation via reqres.in, and file-based reporting.

## MCP Servers Available

| Server | Package | Tools |
|---|---|---|
| `playwright` | `@playwright/mcp@latest` | `browser_navigate`, `browser_click`, `browser_fill`, `browser_screenshot`, `browser_snapshot`, `browser_wait_for` |
| `excel` | `@negokaz/excel-mcp-server` | `excel_read_sheet`, `excel_write_sheet`, `excel_list_sheets`, `excel_create_workbook` |
| `filesystem` | `@modelcontextprotocol/server-filesystem` | `read_file`, `write_file`, `list_directory`, `create_directory` |
| `rest-api` | `mcp-rest-api` | `GET`, `POST`, `PUT`, `DELETE` (against `https://reqres.in/api`) |

## Claude Desktop Setup

Place the contents of `claude-desktop-config.json` into:

```
~/Library/Application Support/Claude/claude_desktop_config.json   # macOS
%APPDATA%\Claude\claude_desktop_config.json                        # Windows
```

Then restart Claude Desktop.

## E2E Scenario Flow

1. **Excel** — Read login scenarios from `tests/assignment2/scenarios/e2e-scenarios.xlsx`
2. **REST API** — Validate user credentials exist via `GET /users?page=1` (reqres.in)
3. **Playwright** — Execute login on `https://www.saucedemo.com`
4. **Playwright** — Navigate inventory, add items to cart, proceed to checkout
5. **Playwright** — Capture screenshot of each major step
6. **Filesystem** — Write HTML report and JSON results to `tests/results/`
7. **Excel** — Update scenario status (PASS/FAIL) back into the workbook

## Example Prompts

> "Read the E2E scenarios from the Excel file, then for each scenario log into saucedemo.com,
>  navigate to the inventory page, add the first product to the cart, and screenshot the result.
>  Save a JSON summary to tests/results/assignment2-results.json."

> "Use the REST API MCP to GET /users from reqres.in, then log the user count to
>  tests/results/api-check.json using the filesystem MCP."

## What you CANNOT do
- Modify files outside `/Users/ams/Desktop/aiTesting` (filesystem MCP is scoped)
- Access databases (no MySQL MCP in this configuration)
- Send emails or post to external services

## ⚠️ Agent Scope vs Assignment 1

| | Assignment 1 (VSCode) | Assignment 2 (Claude Desktop) |
|---|---|---|
| IDE / Client | VSCode | Claude Desktop |
| Agents | Two isolated agents | Single unified agent |
| MCP servers | playwright + excel | playwright + excel + filesystem + rest-api |
| Isolation demo | Excel Agent ≠ Playwright | All tools in one agent scope |
