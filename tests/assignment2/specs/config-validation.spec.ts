/**
 * Assignment 2 — Claude Desktop MCP Configuration Validation
 *
 * Verifies that:
 *   1. claude-desktop-config.json has all 4 required MCP servers
 *   2. The Postman collection covers all CRUD operations + auth
 *   3. The OpenAPI spec defines the expected endpoints
 *   4. The Excel workbook (if built) has the required sheets
 */
import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

const BASE        = path.resolve(__dirname, '..');
const CD_CONFIG   = path.resolve(BASE, 'claude-desktop-config.json');
const POSTMAN     = path.resolve(BASE, 'api/postman-collection.json');
const OPENAPI     = path.resolve(BASE, 'api/openapi.yaml');
const WORKBOOK    = path.resolve(BASE, 'scenarios/e2e-scenarios.xlsx');
const AGENT_DEF   = path.resolve(BASE, 'agents/claude-desktop-agent.md');

// ─── Claude Desktop config ────────────────────────────────────────────────────

test.describe('Assignment 2 — Claude Desktop MCP Configuration', () => {

  test('claude-desktop-config.json exists', () => {
    expect(fs.existsSync(CD_CONFIG), 'claude-desktop-config.json missing').toBe(true);
    console.log('✅ claude-desktop-config.json found');
  });

  test('Config has exactly 4 MCP servers', () => {
    const config = JSON.parse(fs.readFileSync(CD_CONFIG, 'utf-8'));
    const servers = Object.keys(config.mcpServers);
    expect(servers).toContain('playwright');
    expect(servers).toContain('excel');
    expect(servers).toContain('filesystem');
    expect(servers).toContain('rest-api');
    expect(servers.length).toBe(4);
    console.log(`✅ 4 MCP servers configured: ${servers.join(', ')}`);
  });

  test('playwright server uses @playwright/mcp@latest', () => {
    const config = JSON.parse(fs.readFileSync(CD_CONFIG, 'utf-8'));
    const pw = config.mcpServers.playwright;
    expect(pw.command).toBe('npx');
    expect(pw.args).toContain('@playwright/mcp@latest');
    console.log('✅ playwright MCP configured correctly');
  });

  test('excel server uses @negokaz/excel-mcp-server', () => {
    const config = JSON.parse(fs.readFileSync(CD_CONFIG, 'utf-8'));
    const xl = config.mcpServers.excel;
    expect(xl.command).toBe('npx');
    expect(xl.args.some((a: string) => a.includes('excel-mcp-server'))).toBe(true);
    console.log('✅ excel MCP configured correctly');
  });

  test('filesystem server uses @modelcontextprotocol/server-filesystem', () => {
    const config = JSON.parse(fs.readFileSync(CD_CONFIG, 'utf-8'));
    const fs_ = config.mcpServers.filesystem;
    expect(fs_.command).toBe('npx');
    expect(fs_.args.some((a: string) => a.includes('server-filesystem'))).toBe(true);
    console.log('✅ filesystem MCP configured correctly');
  });

  test('rest-api server uses mcp-rest-api with reqres.in base URL', () => {
    const config = JSON.parse(fs.readFileSync(CD_CONFIG, 'utf-8'));
    const ra = config.mcpServers['rest-api'];
    expect(ra.command).toBe('npx');
    expect(ra.args.some((a: string) => a.includes('mcp-rest-api'))).toBe(true);
    expect(ra.env.API_BASE_URL).toContain('reqres.in');
    console.log(`✅ rest-api MCP → ${ra.env.API_BASE_URL}`);
  });

});

// ─── Postman collection ───────────────────────────────────────────────────────

test.describe('Assignment 2 — Postman Collection', () => {

  test('postman-collection.json exists', () => {
    expect(fs.existsSync(POSTMAN), 'postman-collection.json missing').toBe(true);
    console.log('✅ postman-collection.json found');
  });

  test('Collection has Users and Auth folders', () => {
    const collection = JSON.parse(fs.readFileSync(POSTMAN, 'utf-8'));
    const folderNames = collection.item.map((f: any) => f.name);
    expect(folderNames).toContain('Users');
    expect(folderNames).toContain('Auth');
    console.log(`✅ Folders: ${folderNames.join(', ')}`);
  });

  test('Users folder contains CRUD requests', () => {
    const collection = JSON.parse(fs.readFileSync(POSTMAN, 'utf-8'));
    const usersFolder = collection.item.find((f: any) => f.name === 'Users');
    expect(usersFolder).toBeTruthy();
    const methods = usersFolder.item.map((r: any) => r.request.method);
    expect(methods).toContain('GET');
    expect(methods).toContain('POST');
    expect(methods).toContain('PUT');
    expect(methods).toContain('DELETE');
    console.log(`✅ Users methods: ${[...new Set(methods)].join(', ')}`);
  });

  test('Auth folder contains login requests', () => {
    const collection = JSON.parse(fs.readFileSync(POSTMAN, 'utf-8'));
    const authFolder = collection.item.find((f: any) => f.name === 'Auth');
    expect(authFolder).toBeTruthy();
    const requestNames: string[] = authFolder.item.map((r: any) => r.name);
    expect(requestNames.some((n: string) => n.toLowerCase().includes('login'))).toBe(true);
    console.log(`✅ Auth requests: ${requestNames.join(', ')}`);
  });

  test('All requests have test scripts', () => {
    const collection = JSON.parse(fs.readFileSync(POSTMAN, 'utf-8'));
    let allHaveTests = true;
    collection.item.forEach((folder: any) => {
      folder.item.forEach((req: any) => {
        const hasTests = req.event?.some((e: any) => e.listen === 'test');
        if (!hasTests) {
          console.warn(`  ⚠️  No test script on: ${req.name}`);
          allHaveTests = false;
        }
      });
    });
    expect(allHaveTests).toBe(true);
    console.log('✅ All requests have test assertions');
  });

});

// ─── OpenAPI spec ─────────────────────────────────────────────────────────────

test.describe('Assignment 2 — OpenAPI Specification', () => {

  test('openapi.yaml exists', () => {
    expect(fs.existsSync(OPENAPI), 'openapi.yaml missing').toBe(true);
    console.log('✅ openapi.yaml found');
  });

  test('OpenAPI spec has required sections', () => {
    const content = fs.readFileSync(OPENAPI, 'utf-8');
    expect(content).toContain('openapi: 3.0');
    expect(content).toContain('info:');
    expect(content).toContain('servers:');
    expect(content).toContain('paths:');
    expect(content).toContain('components:');
    console.log('✅ OpenAPI spec has all required top-level sections');
  });

  test('OpenAPI spec defines /users and /login endpoints', () => {
    const content = fs.readFileSync(OPENAPI, 'utf-8');
    expect(content).toContain('/users');
    expect(content).toContain('/login');
    console.log('✅ /users and /login endpoints defined');
  });

  test('OpenAPI spec targets reqres.in', () => {
    const content = fs.readFileSync(OPENAPI, 'utf-8');
    expect(content).toContain('reqres.in');
    console.log('✅ OpenAPI server URL targets reqres.in');
  });

});

// ─── Agent definition ─────────────────────────────────────────────────────────

test.describe('Assignment 2 — Claude Desktop Agent Definition', () => {

  test('claude-desktop-agent.md exists', () => {
    expect(fs.existsSync(AGENT_DEF), 'claude-desktop-agent.md missing').toBe(true);
    console.log('✅ claude-desktop-agent.md found');
  });

  test('Agent definition documents all 4 MCP servers', () => {
    const content = fs.readFileSync(AGENT_DEF, 'utf-8');
    expect(content).toContain('playwright');
    expect(content).toContain('excel');
    expect(content).toContain('filesystem');
    expect(content).toContain('rest-api');
    console.log('✅ All 4 MCP servers documented in agent definition');
  });

  test('Agent definition includes setup instructions', () => {
    const content = fs.readFileSync(AGENT_DEF, 'utf-8');
    expect(content).toContain('claude_desktop_config.json');
    console.log('✅ Setup instructions present');
  });

});

// ─── Excel workbook (conditional) ────────────────────────────────────────────

test.describe('Assignment 2 — Excel Workbook (if built)', () => {

  test('e2e-scenarios.xlsx has 4 sheets when present', async () => {
    if (!fs.existsSync(WORKBOOK)) {
      console.log('⏭️  Workbook not built — skipping (run build-e2e-scenarios.ts first)');
      test.skip();
      return;
    }
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(WORKBOOK);
    const sheetNames: string[] = wb.worksheets.map((ws: any) => ws.name);
    expect(sheetNames).toContain('Login Scenarios');
    expect(sheetNames).toContain('E2E Scenarios');
    expect(sheetNames).toContain('REST API Scenarios');
    expect(sheetNames).toContain('Results');
    console.log(`✅ Sheets found: ${sheetNames.join(', ')}`);
  });

});
