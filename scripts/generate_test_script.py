#!/usr/bin/env python3
"""
QA Ultra Tester - TypeScript Test Script Generator
Generates Playwright TypeScript .spec.ts files from test-results.json
"""

import json
import sys
from pathlib import Path
from datetime import datetime


class TestScriptGenerator:
    def __init__(self, output_dir="hasil-test"):
        self.output_dir = Path(output_dir)
        self.tests_dir = self.output_dir / "tests"
        self.tests_dir.mkdir(parents=True, exist_ok=True)
        self.data = {}

    def load_from_json(self, filename="test-results.json"):
        filepath = self.output_dir / filename
        if not filepath.exists():
            print(f"Error: {filepath} not found.")
            return False
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            self.data = json.load(f)
        return True

    def set_data(self, data: dict):
        self.data = data

    def _target_url(self):
        return self.data.get("environment") or self.data.get("target_url") or "https://example.com"

    def _safe_name(self, name: str) -> str:
        return name.replace("'", "\\'").replace('"', '\\"')

    # ── Generate all spec files ──────────────────────────────
    def generate_all(self):
        tests = self.data.get("test_results", [])
        bugs = self.data.get("bugs", [])
        perf = self.data.get("performance", {})
        a11y = self.data.get("accessibility", [])
        url = self._target_url()

        # Group tests by layer
        layers = {}
        for t in tests:
            layer = t.get("layer", "UI").upper()
            layers.setdefault(layer, []).append(t)

        generated = []

        # Smoke / Happy Path (UI layer PASS tests)
        ui_tests = layers.get("UI", [])
        if ui_tests:
            generated.append(self._gen_smoke(ui_tests, url))

        # Negative tests (FAIL tests from any layer)
        fail_tests = [t for t in tests if t.get("status") == "FAIL"]
        if fail_tests:
            generated.append(self._gen_negative(fail_tests, url))

        # Security tests
        security_tests = layers.get("SECURITY", [])
        if security_tests or bugs:
            generated.append(self._gen_security(security_tests, bugs, url))

        # API Validation
        api_tests = layers.get("API", [])
        if api_tests:
            generated.append(self._gen_api(api_tests, url))

        # Accessibility
        if a11y or layers.get("A11Y"):
            generated.append(self._gen_accessibility(layers.get("A11Y", []), a11y, url))

        # Performance
        if perf or layers.get("PERF"):
            generated.append(self._gen_performance(layers.get("PERF", []), perf, url))

        # Chaos
        chaos_tests = layers.get("CHAOS", [])
        if chaos_tests:
            generated.append(self._gen_chaos(chaos_tests, url))

        # Playwright config
        generated.append(self._gen_playwright_config())

        return [g for g in generated if g]

    # ── Smoke (Happy Path) ───────────────────────────────────
    def _gen_smoke(self, tests, url):
        cases = ""
        for t in tests:
            name = self._safe_name(t.get("name", "Unnamed"))
            details = t.get("details", "")
            expected = "PASS" if t.get("status") == "PASS" else "FAIL"
            cases += f"""
  test('{name}', async ({{ page }}) => {{
    await page.goto('{url}');
    // TODO: implement test logic for "{name}"
    // Expected: {expected}
    // Details: {details}
    await expect(page).not.toHaveTitle(/error/i);
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

test.describe('Smoke Tests (Happy Path)', () => {{
  test.beforeEach(async ({{ page }}) => {{
    await page.goto('{url}');
  }});
{cases}}});
"""
        filepath = self.tests_dir / "smoke.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Negative Tests ───────────────────────────────────────
    def _gen_negative(self, tests, url):
        cases = ""
        for t in tests:
            name = self._safe_name(t.get("name", "Unnamed"))
            details = t.get("details", "Should handle invalid input gracefully")
            cases += f"""
  test('{name}', async ({{ page }}) => {{
    await page.goto('{url}');
    // TODO: implement negative test for "{name}"
    // Details: {details}
    // Verify error handling, validation messages, boundary conditions
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

test.describe('Negative Tests', () => {{
{cases}}});
"""
        filepath = self.tests_dir / "negative.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Security Tests ───────────────────────────────────────
    def _gen_security(self, tests, bugs, url):
        bug_cases = ""
        for bug in bugs:
            bid = bug.get("id", "BUG-XXX")
            title = self._safe_name(bug.get("title", ""))
            steps = bug.get("steps", [])
            steps_comment = "\n".join(f"    // {s}" for s in steps)
            expected = self._safe_name(bug.get("expected", ""))
            bug_cases += f"""
  test('[{bid}] {title}', async ({{ page }}) => {{
    await page.goto('{url}');
    // Steps to reproduce:
{steps_comment}
    // Expected: {expected}
    // WARNING: Run ONLY in non-production environment
  }});
"""
        test_cases = ""
        for t in tests:
            name = self._safe_name(t.get("name", "Unnamed"))
            test_cases += f"""
  test('{name}', async ({{ page }}) => {{
    await page.goto('{url}');
    // TODO: implement security test
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

/**
 * SECURITY TESTS
 * WARNING: Only run in non-production/staging environments.
 * These tests may attempt injection payloads.
 */
test.describe('Security Tests', () => {{
{bug_cases}{test_cases}}});
"""
        filepath = self.tests_dir / "security.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── API Validation ───────────────────────────────────────
    def _gen_api(self, tests, url):
        cases = ""
        for t in tests:
            name = self._safe_name(t.get("name", "Unnamed"))
            cases += f"""
  test('{name}', async ({{ page, request }}) => {{
    // Intercept and validate API calls
    const apiLogs: any[] = [];
    await page.route('**/api/**', async (route) => {{
      const response = await route.fetch();
      apiLogs.push({{
        url: route.request().url(),
        method: route.request().method(),
        status: response.status(),
      }});
      await route.fulfill({{ response }});
    }});

    await page.goto('{url}');
    // TODO: trigger API call and validate response
    // Verify: status codes, response schema, no sensitive data leak
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

test.describe('API Validation', () => {{
{cases}}});
"""
        filepath = self.tests_dir / "api-validation.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Accessibility ────────────────────────────────────────
    def _gen_accessibility(self, tests, issues, url):
        issue_checks = ""
        for issue in issues:
            desc = self._safe_name(issue.get("description", ""))
            element = self._safe_name(issue.get("element", ""))
            wcag = issue.get("wcag", "")
            issue_checks += f"""
  test('WCAG {wcag}: {desc}', async ({{ page }}) => {{
    await page.goto('{url}');
    // Element: {element}
    // Check: {desc}
    // TODO: validate fix for this accessibility issue
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

test.describe('Accessibility (WCAG 2.1 AA)', () => {{

  test('Full page axe-core scan', async ({{ page }}) => {{
    await page.goto('{url}');
    const results = await new AxeBuilder({{ page }}).withTags(['wcag2a', 'wcag2aa']).analyze();
    expect(results.violations).toEqual([]);
  }});
{issue_checks}}});
"""
        filepath = self.tests_dir / "accessibility.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Performance ──────────────────────────────────────────
    def _gen_performance(self, tests, perf_metrics, url):
        metric_checks = ""
        for metric, val in perf_metrics.items():
            if isinstance(val, dict):
                threshold = {"LCP": 2500, "FID": 100, "CLS": 0.1, "TTFB": 800, "FCP": 1800}.get(metric, 5000)
                unit = val.get("unit", "ms")
                metric_checks += f"""
  test('Core Web Vital: {metric} < {threshold}{unit}', async ({{ page }}) => {{
    await page.goto('{url}');
    const timing = await page.evaluate(() => {{
      const nav = performance.getEntriesByType('navigation')[0] as PerformanceNavigationTiming;
      const paint = performance.getEntriesByType('paint');
      return {{
        ttfb: nav?.responseStart ?? 0,
        fcp: paint.find(p => p.name === 'first-contentful-paint')?.startTime ?? 0,
        domComplete: nav?.domComplete ?? 0,
      }};
    }});
    // TODO: measure {metric} and assert < {threshold}
    console.log('Timing:', timing);
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

test.describe('Performance (Core Web Vitals)', () => {{
{metric_checks}}});
"""
        filepath = self.tests_dir / "performance.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Chaos ────────────────────────────────────────────────
    def _gen_chaos(self, tests, url):
        cases = ""
        for t in tests:
            name = self._safe_name(t.get("name", "Unnamed"))
            cases += f"""
  test('{name}', async ({{ page, context }}) => {{
    await page.goto('{url}');
    // TODO: implement chaos scenario
    // Examples: network throttle, CPU throttle, clear storage, rapid interactions
  }});
"""
        content = f"""import {{ test, expect }} from '@playwright/test';

/**
 * CHAOS ENGINEERING TESTS
 * Tests app resilience under adverse conditions.
 * Safety: max CPU throttle 6x, max offline 30s, no production storage clear.
 */
test.describe('Chaos Engineering', () => {{
{cases}}});
"""
        filepath = self.tests_dir / "chaos.spec.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath

    # ── Playwright Config ────────────────────────────────────
    def _gen_playwright_config(self):
        content = """import { defineConfig, devices } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

// Clean old artifacts before test execution
const screenshotsDir = path.resolve(__dirname, '../screenshots');
if (fs.existsSync(screenshotsDir)) {
  fs.rmSync(screenshotsDir, { recursive: true, force: true });
}
fs.mkdirSync(screenshotsDir, { recursive: true });

export default defineConfig({
  testDir: './tests',
  timeout: 30_000,
  retries: 2,
  reporter: [['list'], ['json', { outputFile: 'playwright-results.json' }]],
  use: {
    baseURL: process.env.BASE_URL || 'https://example.com',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    trace: 'retain-on-failure',
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
    { name: 'mobile-chrome', use: { ...devices['Pixel 5'] } },
    { name: 'mobile-safari', use: { ...devices['iPhone 14'] } },
  ],
});
"""
        filepath = self.tests_dir / "playwright.config.ts"
        filepath.write_text(content, encoding='utf-8')
        print(f"  Generated: {filepath}")
        return filepath


# ── CLI ──────────────────────────────────────────────────────
if __name__ == "__main__":
    output_dir = sys.argv[1] if len(sys.argv) > 1 else "hasil-test"
    gen = TestScriptGenerator(output_dir)

    if gen.load_from_json():
        print("Generating TypeScript test scripts from test-results.json...")
        files = gen.generate_all()
        print(f"\nGenerated {len(files)} files in {gen.tests_dir}")
    else:
        print("No test-results.json found. Creating demo scripts...")
        demo = {
            "target_url": "https://example.com",
            "test_results": [
                {"name": "Login Page Load", "layer": "UI", "status": "PASS", "duration_ms": 1200, "browser": "Chromium", "details": ""},
                {"name": "Invalid Login", "layer": "UI", "status": "FAIL", "duration_ms": 800, "browser": "Chromium", "details": "No error message shown"},
                {"name": "API Auth Check", "layer": "API", "status": "PASS", "duration_ms": 200, "browser": "Chromium", "details": ""},
            ],
            "bugs": [
                {"id": "BUG-001", "severity": "HIGH", "title": "XSS in search", "location": "/search",
                 "steps": ["1. Go to search", "2. Input <script>alert(1)</script>", "3. Submit"],
                 "actual": "Script executes", "expected": "Input sanitized", "recommendation": "Sanitize input"}
            ],
            "performance": {"LCP": {"value": 2300, "unit": "ms", "rating": "good"}, "TTFB": {"value": 500, "unit": "ms", "rating": "good"}},
            "accessibility": [{"type": "contrast", "element": ".btn", "description": "Low contrast ratio", "wcag": "1.4.3", "severity": "serious"}]
        }
        gen.set_data(demo)
        files = gen.generate_all()
        print(f"\nGenerated {len(files)} demo files in {gen.tests_dir}")
