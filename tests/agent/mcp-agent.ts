/**
 * MCP Agent — unified interface wrapping Playwright + Excel + Filesystem servers.
 *
 * In production Claude reads/writes through actual MCP tool calls.
 * In test code this class provides the same contract as programmatic helpers,
 * so specs stay identical whether driven by Claude or by the test runner.
 */
import { Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

// ─── Types ───────────────────────────────────────────────────────────────────

export interface LoginScenario {
  id: number;
  name: string;
  prompt: string;
  url: string;
  username: string;
  password: string;
  expected: string;
  expectedUrl: string;
  priority: string;
  status: string;
}

export interface NavScenario {
  id: number;
  name: string;
  prompt: string;
  precondition: string;
  steps: string;
  expected: string;
  priority: string;
  status: string;
}

export interface AgentResult {
  scenarioId: number | string;
  name: string;
  status: 'PASS' | 'FAIL' | 'SKIP';
  duration: number;
  actualUrl: string;
  error?: string;
  screenshot?: string;
  runAt: string;
}

export interface PromptTemplate {
  id: number;
  name: string;
  template: string;
  variables: string;
  example: string;
}

// ─── Playwright MCP Tools ────────────────────────────────────────────────────

export class PlaywrightMCP {
  constructor(private page: Page, private screenshotDir: string) {
    fs.mkdirSync(screenshotDir, { recursive: true });
  }

  /** Navigate to URL and wait for load */
  async navigate(url: string): Promise<void> {
    await this.page.goto(url, { waitUntil: 'domcontentloaded', timeout: 20000 });
  }

  /** Fill an input field */
  async fill(selector: string, value: string): Promise<void> {
    await this.page.locator(selector).fill(value);
  }

  /** Click an element */
  async click(selector: string): Promise<void> {
    await this.page.locator(selector).click();
  }

  /** Get text content of an element, null if missing */
  async getText(selector: string): Promise<string | null> {
    const el = this.page.locator(selector);
    return (await el.count()) > 0 ? el.textContent() : null;
  }

  /** Check if element is visible */
  async isVisible(selector: string): Promise<boolean> {
    return this.page.locator(selector).isVisible();
  }

  /** Get current page URL */
  currentUrl(): string {
    return this.page.url();
  }

  /** Take a screenshot and return the file path */
  async screenshot(name: string, fullPage = false): Promise<string> {
    const file = path.join(this.screenshotDir, `${name}-${Date.now()}.png`);
    await this.page.screenshot({ path: file, fullPage });
    return file;
  }

  /** Execute a scenario prompt — interprets the natural language steps */
  async executePrompt(scenario: LoginScenario): Promise<Partial<AgentResult>> {
    const start = Date.now();
    try {
      await this.navigate(scenario.url);
      await this.fill('#user-name', scenario.username);
      await this.fill('#password', scenario.password);
      await this.click('[data-test="login-button"]');

      // small wait for navigation/error
      await this.page.waitForTimeout(800);

      const actualUrl   = this.currentUrl();
      const errorEl     = await this.getText('[data-test="error"]');
      const screenshot  = await this.screenshot(`scenario-${scenario.id}-${scenario.name.replace(/\s+/g, '-').toLowerCase()}`);

      return { actualUrl, screenshot, duration: Date.now() - start, error: errorEl ?? undefined };
    } catch (err: any) {
      return { actualUrl: this.currentUrl(), duration: Date.now() - start, error: err.message };
    }
  }

  /** Select a dropdown option */
  async selectOption(selector: string, value: string): Promise<void> {
    await this.page.locator(selector).selectOption(value);
  }

  /** Get all text contents matching a selector */
  async getAllTexts(selector: string): Promise<string[]> {
    return this.page.locator(selector).allTextContents();
  }

  /** Count elements matching selector */
  async count(selector: string): Promise<number> {
    return this.page.locator(selector).count();
  }

  /** Reload the page */
  async reload(): Promise<void> {
    await this.page.reload({ waitUntil: 'domcontentloaded' });
  }
}

// ─── Excel MCP Tools ─────────────────────────────────────────────────────────

export class ExcelMCP {
  private wb: ExcelJS.Workbook;

  constructor(private filePath: string) {
    this.wb = new ExcelJS.Workbook();
  }

  async load(): Promise<void> {
    if (fs.existsSync(this.filePath)) {
      await this.wb.xlsx.readFile(this.filePath);
    }
  }

  /** Read all scenarios from a named sheet */
  readSheet<T>(sheetName: string): T[] {
    const sheet = this.wb.getWorksheet(sheetName);
    if (!sheet) return [];
    const headers: string[] = [];
    const rows: T[] = [];

    sheet.eachRow((row, idx) => {
      if (idx === 1) {
        row.eachCell(cell => headers.push(String(cell.value ?? '')));
        return;
      }
      const obj: any = {};
      headers.forEach((h, i) => {
        const val = row.getCell(i + 1).value;
        obj[h.replace(/\s+/g, '')] = val ?? '';
      });
      if (Object.values(obj).some(v => v !== '')) rows.push(obj);
    });
    return rows;
  }

  /** Read login scenarios (typed) */
  getLoginScenarios(): LoginScenario[] {
    const sheet = this.wb.getWorksheet('Login Scenarios');
    if (!sheet) return [];
    const out: LoginScenario[] = [];
    sheet.eachRow((row, i) => {
      if (i === 1) return;
      out.push({
        id:          row.getCell(1).value as number,
        name:        String(row.getCell(2).value ?? ''),
        prompt:      String(row.getCell(3).value ?? ''),
        url:         String(row.getCell(4).value ?? ''),
        username:    String(row.getCell(5).value ?? ''),
        password:    String(row.getCell(6).value ?? ''),
        expected:    String(row.getCell(7).value ?? ''),
        expectedUrl: String(row.getCell(8).value ?? ''),
        priority:    String(row.getCell(9).value ?? ''),
        status:      String(row.getCell(10).value ?? ''),
      });
    });
    return out.filter(s => s.name);
  }

  /** Read navigation scenarios (typed) */
  getNavScenarios(): NavScenario[] {
    const sheet = this.wb.getWorksheet('Navigation Scenarios');
    if (!sheet) return [];
    const out: NavScenario[] = [];
    sheet.eachRow((row, i) => {
      if (i === 1) return;
      out.push({
        id:           row.getCell(1).value as number,
        name:         String(row.getCell(2).value ?? ''),
        prompt:       String(row.getCell(3).value ?? ''),
        precondition: String(row.getCell(4).value ?? ''),
        steps:        String(row.getCell(5).value ?? ''),
        expected:     String(row.getCell(6).value ?? ''),
        priority:     String(row.getCell(7).value ?? ''),
        status:       String(row.getCell(8).value ?? ''),
      });
    });
    return out.filter(s => s.name);
  }

  /** Read prompt templates */
  getPromptTemplates(): PromptTemplate[] {
    const sheet = this.wb.getWorksheet('Prompt Templates');
    if (!sheet) return [];
    const out: PromptTemplate[] = [];
    sheet.eachRow((row, i) => {
      if (i === 1) return;
      out.push({
        id:        row.getCell(1).value as number,
        name:      String(row.getCell(2).value ?? ''),
        template:  String(row.getCell(3).value ?? ''),
        variables: String(row.getCell(4).value ?? ''),
        example:   String(row.getCell(5).value ?? ''),
      });
    });
    return out.filter(t => t.name);
  }

  /** Update the status of a scenario row in-place */
  async updateScenarioStatus(sheetName: string, id: number, status: 'PASS' | 'FAIL' | 'SKIP'): Promise<void> {
    const sheet = this.wb.getWorksheet(sheetName);
    if (!sheet) return;
    sheet.eachRow(row => {
      if (row.getCell(1).value === id) {
        const statusCol = sheetName === 'Login Scenarios' ? 10 : 8;
        const cell = row.getCell(statusCol);
        cell.value = status;
        const color = status === 'PASS' ? 'FF92D050' : status === 'FAIL' ? 'FFFF0000' : 'FFFFC000';
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
      }
    });
    await this.wb.xlsx.writeFile(this.filePath);
  }

  /** Write all agent results to the Results sheet */
  async writeResults(results: AgentResult[]): Promise<void> {
    let sheet = this.wb.getWorksheet('Results');
    if (!sheet) {
      sheet = this.wb.addWorksheet('Results');
    }
    // rebuild from row 2
    sheet.spliceRows(2, sheet.rowCount);

    results.forEach((r, i) => {
      const row = sheet!.getRow(i + 2);
      row.getCell(1).value = r.scenarioId;
      row.getCell(2).value = r.name;
      row.getCell(3).value = r.status;
      row.getCell(4).value = r.duration;
      row.getCell(5).value = r.actualUrl;
      row.getCell(6).value = r.error ?? '';
      row.getCell(7).value = r.screenshot ?? '';
      row.getCell(8).value = r.runAt;

      const color = r.status === 'PASS' ? 'FF92D050' : r.status === 'FAIL' ? 'FFFF0000' : 'FFFFC000';
      row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    });

    await this.wb.xlsx.writeFile(this.filePath);
  }

  async save(): Promise<void> {
    await this.wb.xlsx.writeFile(this.filePath);
  }
}

// ─── Filesystem MCP Tools ─────────────────────────────────────────────────────

export class FilesystemMCP {
  constructor(private baseDir: string) {
    fs.mkdirSync(baseDir, { recursive: true });
  }

  resolve(...parts: string[]): string {
    return path.join(this.baseDir, ...parts);
  }

  ensureDir(rel: string): void {
    fs.mkdirSync(this.resolve(rel), { recursive: true });
  }

  /** Write plain text */
  write(rel: string, content: string): string {
    const p = this.resolve(rel);
    fs.mkdirSync(path.dirname(p), { recursive: true });
    fs.writeFileSync(p, content, 'utf-8');
    return p;
  }

  /** Append a timestamped log line */
  log(rel: string, message: string): void {
    const p = this.resolve(rel);
    fs.mkdirSync(path.dirname(p), { recursive: true });
    fs.appendFileSync(p, `[${new Date().toISOString()}] ${message}\n`, 'utf-8');
  }

  /** Read file content */
  read(rel: string): string {
    return fs.readFileSync(this.resolve(rel), 'utf-8');
  }

  /** Write / read JSON */
  writeJson(rel: string, data: any): string {
    return this.write(rel, JSON.stringify(data, null, 2));
  }

  readJson<T>(rel: string): T {
    return JSON.parse(this.read(rel));
  }

  /** Check existence */
  exists(rel: string): boolean {
    return fs.existsSync(this.resolve(rel));
  }

  /** List files in a subdirectory */
  list(rel = ''): string[] {
    const p = this.resolve(rel);
    return fs.existsSync(p) ? fs.readdirSync(p) : [];
  }

  /** Write HTML report */
  writeHtml(rel: string, title: string, results: AgentResult[]): string {
    const pass   = results.filter(r => r.status === 'PASS').length;
    const fail   = results.filter(r => r.status === 'FAIL').length;
    const skip   = results.filter(r => r.status === 'SKIP').length;
    const rows   = results.map(r => {
      const bg = r.status === 'PASS' ? '#92d050' : r.status === 'FAIL' ? '#ff4444' : '#ffc000';
      const sc = r.screenshot ? `<a href="${r.screenshot}" target="_blank">📸</a>` : '';
      return `<tr>
        <td>${r.scenarioId}</td>
        <td>${r.name}</td>
        <td style="background:${bg};font-weight:bold;color:${r.status==='FAIL'?'#fff':'#000'}">${r.status}</td>
        <td>${r.duration}ms</td>
        <td>${r.actualUrl}</td>
        <td>${r.error ?? ''}</td>
        <td>${sc}</td>
        <td>${r.runAt}</td>
      </tr>`;
    }).join('\n');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>${title}</title>
  <style>
    body{font-family:Arial,sans-serif;margin:24px;background:#f5f5f5}
    h1{color:#4472c4;margin-bottom:4px}
    .meta{color:#666;font-size:13px;margin-bottom:16px}
    .badges{display:flex;gap:12px;margin-bottom:20px}
    .badge{padding:8px 18px;border-radius:6px;font-weight:bold;font-size:15px}
    .pass{background:#92d050} .fail{background:#ff4444;color:#fff}
    .skip{background:#ffc000} .total{background:#4472c4;color:#fff}
    table{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.1)}
    th{background:#4472c4;color:#fff;padding:10px 12px;text-align:left;font-size:13px}
    td{border:1px solid #ddd;padding:8px 10px;font-size:12px;vertical-align:top}
    tr:nth-child(even) td{background:#f9f9f9}
  </style>
</head>
<body>
<h1>🤖 MCP Agent — ${title}</h1>
<p class="meta">Generated: ${new Date().toISOString()} | Playwright + Excel + Filesystem MCP</p>
<div class="badges">
  <div class="badge pass">✅ PASS: ${pass}</div>
  <div class="badge fail">❌ FAIL: ${fail}</div>
  <div class="badge skip">⚠️ SKIP: ${skip}</div>
  <div class="badge total">🔢 TOTAL: ${results.length}</div>
</div>
<table>
  <thead><tr><th>ID</th><th>Scenario</th><th>Status</th><th>Duration</th><th>URL</th><th>Error</th><th>Shot</th><th>Run At</th></tr></thead>
  <tbody>${rows}</tbody>
</table>
</body>
</html>`;
    return this.write(rel, html);
  }
}

// ─── Unified MCPAgent ─────────────────────────────────────────────────────────

export class MCPAgent {
  readonly playwright: PlaywrightMCP;
  readonly excel:      ExcelMCP;
  readonly filesystem: FilesystemMCP;

  private runLog: string[] = [];

  constructor(
    page: Page,
    excelFile: string,
    resultsDir: string,
  ) {
    const screenshotDir = path.join(resultsDir, 'screenshots');
    this.playwright  = new PlaywrightMCP(page, screenshotDir);
    this.excel       = new ExcelMCP(excelFile);
    this.filesystem  = new FilesystemMCP(resultsDir);
  }

  /** Load the Excel workbook (must call before reading scenarios) */
  async init(): Promise<void> {
    await this.excel.load();
    this.log('MCPAgent initialised');
  }

  /** Emit a structured log entry to filesystem + in-memory */
  log(msg: string): void {
    const entry = `[${new Date().toISOString()}] ${msg}`;
    this.runLog.push(entry);
    this.filesystem.log('logs/agent.log', msg);
  }

  /** Save in-memory log to file */
  flushLog(): void {
    this.filesystem.write('logs/run-trace.log', this.runLog.join('\n'));
  }

  /** Build an AgentResult from a scenario + raw playwright output */
  buildResult(
    scenario: LoginScenario,
    playwrightOut: Partial<AgentResult>,
    pass: boolean,
    errorOverride?: string,
  ): AgentResult {
    return {
      scenarioId: scenario.id,
      name:       scenario.name,
      status:     pass ? 'PASS' : 'FAIL',
      duration:   playwrightOut.duration ?? 0,
      actualUrl:  playwrightOut.actualUrl ?? '',
      error:      errorOverride ?? playwrightOut.error,
      screenshot: playwrightOut.screenshot,
      runAt:      new Date().toISOString(),
    };
  }

  /** Finalise — write results to Excel + HTML, flush log */
  async finalise(results: AgentResult[], reportTitle: string): Promise<void> {
    await this.excel.writeResults(results);
    this.filesystem.writeJson('reports/agent-results.json', results);
    this.filesystem.writeHtml('reports/agent-report.html', reportTitle, results);

    const pass = results.filter(r => r.status === 'PASS').length;
    const fail = results.filter(r => r.status === 'FAIL').length;
    this.log(`Run complete — ${pass} PASS / ${fail} FAIL / ${results.length} TOTAL`);
    this.flushLog();
  }
}
