/**
 * Prompt Builder — converts natural language scenario prompts into
 * structured Playwright action sequences.
 *
 * Works as an MCP bridge: Claude reads the "Prompt Templates" Excel sheet
 * (via Excel MCP), fills variables, and this module executes the result.
 */
import { Page } from '@playwright/test';
import { LoginScenario, NavScenario, PromptTemplate } from './mcp-agent';

// ─── Action types ─────────────────────────────────────────────────────────────

export type ActionType =
  | 'navigate'
  | 'fill'
  | 'click'
  | 'assert_url'
  | 'assert_visible'
  | 'assert_text'
  | 'assert_not_visible'
  | 'screenshot'
  | 'wait'
  | 'select'
  | 'reload'
  | 'assert_count';

export interface Action {
  type: ActionType;
  selector?: string;
  value?: string;
  url?: string;
  expected?: string | number;
  description: string;
}

export interface BuiltScenario {
  id: number | string;
  name: string;
  prompt: string;
  actions: Action[];
}

// ─── Prompt Interpreter ───────────────────────────────────────────────────────

export class PromptBuilder {
  private templates: PromptTemplate[];

  constructor(templates: PromptTemplate[] = []) {
    this.templates = templates;
  }

  /**
   * Convert a LoginScenario (read from Excel MCP) into executable actions.
   * Follows the natural language in scenario.prompt.
   */
  fromLoginScenario(scenario: LoginScenario): BuiltScenario {
    const actions: Action[] = [];

    // 1. Navigate
    actions.push({
      type: 'navigate',
      url: scenario.url,
      description: `Navigate to ${scenario.url}`,
    });

    // 2. Fill username (if provided)
    if (scenario.username !== '') {
      actions.push({
        type: 'fill',
        selector: '#user-name',
        value: scenario.username,
        description: `Fill username "${scenario.username}"`,
      });
    }

    // 3. Fill password (if provided)
    if (scenario.password !== '') {
      actions.push({
        type: 'fill',
        selector: '#password',
        value: scenario.password,
        description: `Fill password`,
      });
    }

    // 4. Click login
    actions.push({
      type: 'click',
      selector: '[data-test="login-button"]',
      description: 'Click Login button',
    });

    // 5. Assert based on expected outcome
    if (scenario.expectedUrl && scenario.expectedUrl !== scenario.url) {
      actions.push({
        type: 'assert_url',
        expected: scenario.expectedUrl,
        description: `Assert URL is "${scenario.expectedUrl}"`,
      });
    }

    // Check expected text for error scenarios
    if (scenario.expected.toLowerCase().includes('error')) {
      const errorText = scenario.expected.replace(/^error:\s*/i, '').replace(/^sorry,\s*/i, '');
      actions.push({
        type: 'assert_text',
        selector: '[data-test="error"]',
        expected: errorText,
        description: `Assert error message contains "${errorText}"`,
      });
    }

    // Logout scenario
    if (scenario.prompt.toLowerCase().includes('logout')) {
      actions.push(
        { type: 'click', selector: '#react-burger-menu-btn', description: 'Open sidebar menu' },
        { type: 'click', selector: '#logout_sidebar_link', description: 'Click Logout' },
        { type: 'assert_url', expected: scenario.url, description: 'Assert back on login page' },
      );
    }

    // Session persistence scenario
    if (scenario.prompt.toLowerCase().includes('refresh')) {
      actions.push(
        { type: 'reload', description: 'Reload page' },
        { type: 'assert_url', expected: scenario.expectedUrl, description: 'Assert URL after reload' },
      );
    }

    // Screenshot always at the end
    actions.push({
      type: 'screenshot',
      value: `scenario-${scenario.id}`,
      description: `Screenshot: ${scenario.name}`,
    });

    return { id: scenario.id, name: scenario.name, prompt: scenario.prompt, actions };
  }

  /**
   * Convert a NavScenario (from Excel MCP) into executable actions.
   * Parses the pipe-delimited Steps column.
   */
  fromNavScenario(scenario: NavScenario): BuiltScenario {
    const actions: Action[] = [];
    const stepLines = scenario.steps.split('|').map(s => s.trim()).filter(Boolean);

    for (const step of stepLines) {
      const lower = step.toLowerCase();

      if (lower.includes('navigate to') || lower.includes('go to')) {
        const url = step.match(/https?:\/\/[^\s]+/)?.[0];
        actions.push({ type: 'navigate', url: url ?? '', description: step });
      } else if (lower.includes('click') && lower.includes('sort')) {
        actions.push({ type: 'select', selector: '.product_sort_container', value: 'lohi', description: step });
      } else if (lower.includes('click') && lower.includes('add to cart')) {
        actions.push({ type: 'click', selector: '.btn_add_to_cart:first-of-type', description: step });
      } else if (lower.includes('check cart badge')) {
        actions.push({ type: 'assert_visible', selector: '.shopping_cart_badge', description: step });
      } else if (lower.includes('click') && lower.includes('product name')) {
        actions.push({ type: 'click', selector: '.inventory_item_name:first-of-type', description: step });
      } else if (lower.includes('click') && lower.includes('checkout')) {
        actions.push({ type: 'click', selector: '[data-test="checkout"]', description: step });
      } else if (lower.includes('fill') && lower.includes('name')) {
        actions.push(
          { type: 'fill', selector: '[data-test="firstName"]', value: 'MCP', description: 'Fill first name' },
          { type: 'fill', selector: '[data-test="lastName"]', value: 'Agent', description: 'Fill last name' },
          { type: 'fill', selector: '[data-test="postalCode"]', value: '12345', description: 'Fill postal code' },
        );
      } else if (lower.includes('click') && lower.includes('continue')) {
        actions.push({ type: 'click', selector: '[data-test="continue"]', description: step });
      } else if (lower.includes('finish')) {
        actions.push({ type: 'click', selector: '[data-test="finish"]', description: step });
      } else if (lower.includes('assert') || lower.includes('verify')) {
        const expectedText = scenario.expected;
        actions.push({ type: 'assert_text', selector: 'body', expected: expectedText, description: step });
      } else if (lower.includes('count') || lower.includes('read all prices')) {
        actions.push({ type: 'assert_count', selector: '.inventory_item', expected: 6, description: step });
      }
    }

    actions.push({ type: 'screenshot', value: `nav-${scenario.id}`, description: `Screenshot: ${scenario.name}` });
    return { id: scenario.id, name: scenario.name, prompt: scenario.prompt, actions };
  }

  /**
   * Fill a template from the Prompt Templates sheet with concrete variables.
   */
  fillTemplate(templateName: string, variables: Record<string, string>): string {
    const tpl = this.templates.find(t => t.name === templateName);
    if (!tpl) throw new Error(`Template "${templateName}" not found`);
    let result = tpl.template;
    for (const [k, v] of Object.entries(variables)) {
      result = result.replaceAll(`{${k}}`, v);
    }
    return result;
  }

  /** Parse a raw prompt string into a BuiltScenario (for ad-hoc use) */
  fromPrompt(id: string, name: string, prompt: string): BuiltScenario {
    const actions: Action[] = [];
    const sentences = prompt.split(',').map(s => s.trim());

    for (const s of sentences) {
      const lower = s.toLowerCase();
      if (lower.startsWith('navigate to')) {
        const url = s.match(/https?:\/\/[^\s,]+/)?.[0] ?? '';
        actions.push({ type: 'navigate', url, description: s });
      } else if (lower.startsWith('fill') || lower.startsWith('enter')) {
        const sel = s.match(/#([\w-]+)/)?.[0] ?? '';
        const val = s.match(/"([^"]+)"/)?.[1] ?? '';
        actions.push({ type: 'fill', selector: sel, value: val, description: s });
      } else if (lower.startsWith('click')) {
        const sel = s.match(/\[([^\]]+)\]|#([\w-]+)|"([^"]+)"/)?.[0] ?? '';
        actions.push({ type: 'click', selector: sel, description: s });
      } else if (lower.startsWith('assert') || lower.startsWith('verify')) {
        if (lower.includes('url')) {
          const url = s.match(/https?:\/\/[^\s,]+/)?.[0] ?? s.match(/"([^"]+)"/)?.[1] ?? '';
          actions.push({ type: 'assert_url', expected: url, description: s });
        } else {
          const text = s.match(/"([^"]+)"/)?.[1] ?? '';
          actions.push({ type: 'assert_text', selector: 'body', expected: text, description: s });
        }
      } else if (lower.includes('screenshot')) {
        actions.push({ type: 'screenshot', value: `prompt-${id}`, description: s });
      }
    }

    return { id, name, prompt, actions };
  }
}

// ─── Action Executor ──────────────────────────────────────────────────────────

export class ActionExecutor {
  constructor(private page: Page) {}

  async run(action: Action): Promise<{ ok: boolean; detail?: string }> {
    try {
      switch (action.type) {
        case 'navigate':
          await this.page.goto(action.url!, { waitUntil: 'domcontentloaded', timeout: 20000 });
          break;

        case 'fill':
          await this.page.locator(action.selector!).fill(action.value ?? '');
          break;

        case 'click':
          await this.page.locator(action.selector!).click({ timeout: 8000 });
          break;

        case 'select':
          await this.page.locator(action.selector!).selectOption(action.value ?? '');
          break;

        case 'reload':
          await this.page.reload({ waitUntil: 'domcontentloaded' });
          break;

        case 'assert_url': {
          await this.page.waitForTimeout(600);
          const cur = this.page.url();
          if (!cur.includes(String(action.expected))) {
            return { ok: false, detail: `URL "${cur}" does not contain "${action.expected}"` };
          }
          break;
        }

        case 'assert_visible': {
          const visible = await this.page.locator(action.selector!).isVisible({ timeout: 5000 });
          if (!visible) return { ok: false, detail: `"${action.selector}" not visible` };
          break;
        }

        case 'assert_not_visible': {
          const vis = await this.page.locator(action.selector!).isVisible();
          if (vis) return { ok: false, detail: `"${action.selector}" should not be visible` };
          break;
        }

        case 'assert_text': {
          await this.page.waitForTimeout(400);
          const el = this.page.locator(action.selector!);
          if ((await el.count()) === 0) return { ok: false, detail: `"${action.selector}" not found` };
          const text = (await el.textContent()) ?? '';
          if (!text.toLowerCase().includes(String(action.expected).toLowerCase())) {
            return { ok: false, detail: `Text "${text}" does not contain "${action.expected}"` };
          }
          break;
        }

        case 'assert_count': {
          const count = await this.page.locator(action.selector!).count();
          if (count !== action.expected) {
            return { ok: false, detail: `Count ${count} ≠ expected ${action.expected}` };
          }
          break;
        }

        case 'wait':
          await this.page.waitForTimeout(Number(action.value) || 500);
          break;

        case 'screenshot': {
          const dir = 'tests/results/screenshots';
          const p = `${dir}/${action.value}-${Date.now()}.png`;
          await this.page.screenshot({ path: p, fullPage: false });
          break;
        }
      }
      return { ok: true };
    } catch (err: any) {
      return { ok: false, detail: err.message };
    }
  }
}
