/**
 * Scenario Runner — orchestrates MCPAgent + PromptBuilder.
 *
 * Reads scenarios from Excel (Excel MCP),
 * builds action sequences (PromptBuilder),
 * executes via Playwright (Playwright MCP),
 * writes results back to Excel + filesystem (Filesystem MCP).
 */
import { Page } from '@playwright/test';
import * as path from 'path';
import { MCPAgent, AgentResult, LoginScenario, NavScenario } from './mcp-agent';
import { PromptBuilder, ActionExecutor, BuiltScenario } from './prompt-builder';

const SCENARIOS_XL = path.resolve(__dirname, '../scenarios/login-scenarios.xlsx');
const RESULTS_DIR  = path.resolve(__dirname, '../results');

export class ScenarioRunner {
  private agent:    MCPAgent;
  private builder:  PromptBuilder;
  private executor: ActionExecutor;
  private results:  AgentResult[] = [];

  constructor(page: Page) {
    this.agent    = new MCPAgent(page, SCENARIOS_XL, RESULTS_DIR);
    this.builder  = new PromptBuilder();
    this.executor = new ActionExecutor(page);
  }

  /** Bootstrap: load Excel, build PromptBuilder with templates */
  async init(): Promise<void> {
    await this.agent.init();
    const templates = this.agent.excel.getPromptTemplates();
    this.builder    = new PromptBuilder(templates);
    this.agent.log(`Loaded ${templates.length} prompt templates from Excel`);
  }

  // ─── Login Scenario Runner ──────────────────────────────────────────────────

  /** Run a single LoginScenario from Excel and return AgentResult */
  async runLoginScenario(scenario: LoginScenario): Promise<AgentResult> {
    this.agent.log(`▶  [${scenario.priority}] ${scenario.name}`);
    this.agent.log(`   Prompt: ${scenario.prompt}`);

    const built   = this.builder.fromLoginScenario(scenario);
    const start   = Date.now();
    let pass      = true;
    let firstError: string | undefined;
    let actualUrl = '';
    let screenshot = '';

    for (const action of built.actions) {
      this.agent.log(`   → ${action.description}`);
      const { ok, detail } = await this.executor.run(action);

      if (action.type === 'screenshot' && ok) {
        screenshot = `tests/results/screenshots/${action.value}-${Date.now()}.png`;
      }

      if (!ok) {
        this.agent.log(`   ✗ FAILED: ${detail}`);
        if (!firstError) firstError = detail;
        pass = false;
        // don't abort — record error but continue to next scenario
        break;
      }
    }

    // grab final URL from page
    try { actualUrl = this.agent.playwright.currentUrl(); } catch {}

    const result: AgentResult = {
      scenarioId: scenario.id,
      name:       scenario.name,
      status:     pass ? 'PASS' : 'FAIL',
      duration:   Date.now() - start,
      actualUrl,
      error:      firstError,
      screenshot,
      runAt:      new Date().toISOString(),
    };

    this.results.push(result);
    this.agent.log(`   ${result.status} (${result.duration}ms)\n`);
    await this.agent.excel.updateScenarioStatus('Login Scenarios', scenario.id, result.status);
    return result;
  }

  /** Run ALL login scenarios from the Excel workbook */
  async runAllLoginScenarios(): Promise<AgentResult[]> {
    const scenarios = this.agent.excel.getLoginScenarios();
    this.agent.log(`Excel MCP — loaded ${scenarios.length} login scenarios`);

    for (const s of scenarios) {
      await this.runLoginScenario(s);
    }
    return this.results;
  }

  /** Run a subset of scenarios by priority */
  async runByPriority(priority: 'P0' | 'P1'): Promise<AgentResult[]> {
    const all     = this.agent.excel.getLoginScenarios();
    const subset  = all.filter(s => s.priority === priority);
    this.agent.log(`Running ${subset.length} ${priority} scenarios`);

    for (const s of subset) {
      await this.runLoginScenario(s);
    }
    return this.results;
  }

  // ─── Navigation Scenario Runner ─────────────────────────────────────────────

  /** Run a single NavScenario (requires page to already be logged in) */
  async runNavScenario(scenario: NavScenario): Promise<AgentResult> {
    this.agent.log(`▶  NAV [${scenario.priority}] ${scenario.name}`);
    const built = this.builder.fromNavScenario(scenario);
    return this._executeBuilt(built, scenario.id, scenario.name);
  }

  // ─── Ad-hoc Prompt Runner ───────────────────────────────────────────────────

  /**
   * Accept a raw natural language prompt and run it immediately.
   * Mirrors what Claude does when using the Playwright MCP server.
   */
  async runPrompt(id: string, name: string, prompt: string): Promise<AgentResult> {
    this.agent.log(`▶  PROMPT "${name}": ${prompt}`);
    const built = this.builder.fromPrompt(id, name, prompt);
    return this._executeBuilt(built, id, name);
  }

  // ─── Shared executor ────────────────────────────────────────────────────────

  private async _executeBuilt(built: BuiltScenario, id: number | string, name: string): Promise<AgentResult> {
    const start = Date.now();
    let pass = true;
    let firstError: string | undefined;
    let screenshot = '';

    for (const action of built.actions) {
      this.agent.log(`   → ${action.description}`);
      const { ok, detail } = await this.executor.run(action);
      if (action.type === 'screenshot') screenshot = `tests/results/screenshots/${action.value}.png`;
      if (!ok) {
        if (!firstError) firstError = detail;
        pass = false;
        break;
      }
    }

    let actualUrl = '';
    try { actualUrl = this.agent.playwright.currentUrl(); } catch {}

    const result: AgentResult = {
      scenarioId: id,
      name,
      status: pass ? 'PASS' : 'FAIL',
      duration: Date.now() - start,
      actualUrl,
      error: firstError,
      screenshot,
      runAt: new Date().toISOString(),
    };

    this.results.push(result);
    this.agent.log(`   ${result.status} (${result.duration}ms)`);
    return result;
  }

  // ─── Finalise ────────────────────────────────────────────────────────────────

  async finalise(reportTitle = 'MCP Agent Login Test Report'): Promise<AgentResult[]> {
    await this.agent.finalise(this.results, reportTitle);

    // Also write a CSV via Filesystem MCP
    const csv = [
      'Scenario ID,Name,Status,Duration (ms),Actual URL,Error,Run At',
      ...this.results.map(r =>
        [r.scenarioId, `"${r.name}"`, r.status, r.duration, r.actualUrl, `"${r.error ?? ''}"`, r.runAt].join(',')
      ),
    ].join('\n');
    this.agent.filesystem.write('reports/agent-results.csv', csv);

    const pass = this.results.filter(r => r.status === 'PASS').length;
    const fail = this.results.filter(r => r.status === 'FAIL').length;
    console.log(`\n${'═'.repeat(50)}`);
    console.log('  MCP AGENT SCENARIO RUNNER — FINAL SUMMARY');
    console.log(`${'═'.repeat(50)}`);
    console.log(`  Total   : ${this.results.length}`);
    console.log(`  PASS    : ${pass} ✅`);
    console.log(`  FAIL    : ${fail} ❌`);
    console.log(`  Pass rate: ${this.results.length ? ((pass / this.results.length) * 100).toFixed(1) : 0}%`);
    console.log(`${'═'.repeat(50)}\n`);

    return this.results;
  }

  getResults(): AgentResult[] { return this.results; }
}
