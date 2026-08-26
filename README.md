# QA Ultra Tester

**Enterprise-Grade E2E Testing Skill for OpenCode**

Skill pengujian kualitas perangkat lunak tingkat ultra yang menggabungkan UI testing, API validation, accessibility compliance, cross-browser testing, database validation, video recording, auto bug reporting, test data generation, performance budgeting, flaky test detection, dan chaos engineering lanjutan.

---

## Features

### 10 Testing Layers

| Layer | Description |
|-------|-------------|
| **UI/UX Testing** | 100% element interaction coverage |
| **API Validation** | Intercept & validate every network request |
| **Accessibility (WCAG)** | WCAG 2.1 AA compliance checking |
| **Cross-Browser** | Chromium, Firefox, WebKit testing |
| **Database Validation** | Data integrity verification |
| **Video Recording** | Full test session recording |
| **Auto Bug Reporting** | GitHub/Jira integration |
| **Test Data Generator** | Synthetic data with boundary cases |
| **Performance Budget** | Core Web Vitals tracking |
| **Flaky Test Detection** | Retry & reliability analysis |
| **Chaos Engineering** | Network, CPU, browser chaos |

---

## Installation

### Prerequisites

- OpenCode CLI
- Node.js 20+
- Python 3.10+
- Git

### Install via Git

```bash
# Clone skill ke folder OpenCode
git clone https://github.com/your-repo/qa-ultra-tester.git ~/.config/opencode/skills/qa-ultra-tester

# Install dependencies
pip install playwright openpyxl python-docx pillow
npx playwright install --with-deps chromium firefox webkit
```

### Install via npm (Global)

```bash
npm install -g qa-ultra-tester
```

---

## Usage

### Basic Commands

| Command | Description |
|---------|-------------|
| `/qa-test <url>` | Full E2E testing (all layers) |
| `/qa-api <url>` | API-focused testing |
| `/qa-a11y <url>` | Accessibility audit |
| `/qa-perf <url>` | Performance testing |
| `/qa-security <url>` | Security testing |
| `/qa-regression <url>` | Regression testing |
| `/qa-chaos <url>` | Chaos engineering |
| `/qa-report` | Generate reports |

### Natural Language

```
"Testing E2E untuk https://example.com"
"QA testing https://example.com berdasarkan FSD ini"
"Audit accessibility https://example.com"
"Security test https://example.com/login"
"Performance test https://example.com"
```

---

## Multi-Mode Input

| Mode | Input | Description |
|------|-------|-------------|
| **Script-Driven** | Test script + FSD + URL | Execute existing tests |
| **Spec-Driven** | FSD/PRD + URL | Generate tests from specs |
| **Whitebox** | Source code + URL | Test based on code analysis |
| **Blackbox** | URL only | Explore and test blindly |
| **API-First** | API docs + URL | Test all endpoints |
| **Regression** | Previous results + URL | Compare with baseline |

---

## Output Artifacts

```
hasil-test/
├── QA_TEST_REPORT.docx        # Word report with screenshots
├── QA_TEST_REPORT.html        # Interactive HTML report
├── test-results.json          # Machine-readable results
├── screenshots/               # Bug evidence
├── videos/                    # Test recordings
├── api-logs/                  # Network logs
├── tests/                     # Playwright scripts
├── ci-cd/                     # Pipeline templates
├── test-data/                 # Generated fixtures
├── performance/               # Core Web Vitals
├── accessibility/             # WCAG reports
└── flaky-analysis/            # Reliability scores
```

---

## Quality Gates

| Severity | Criteria | Action |
|----------|----------|--------|
| 🔴 CRITICAL | Security, data loss, auth bypass | Block release |
| 🟠 HIGH | Validation bypass, broken features | Fix before release |
| 🟡 MEDIUM | UX issues, accessibility | Fix after release |
| 🟢 LOW | Cosmetic, documentation | Backlog |

---

## Scripts

### Generate Report

```bash
python scripts/generate_report.py
```

### Generate Test Data

```bash
python scripts/generate_test_data.py
```

### Analyze Performance

```bash
python scripts/performance_analyzer.py
```

### Check Accessibility

```bash
python scripts/accessibility_checker.py
```

---

## CI/CD Integration

### GitHub Actions

Copy `ci-cd/playwright-e2e.yml` ke `.github/workflows/`:

```yaml
# Add secrets
# BASE_URL: Your target URL
# SLACK_WEBHOOK_URL: Slack notification webhook
```

### GitLab CI

Copy `ci-cd/.gitlab-ci.yml` ke root project:

```yaml
# Add variables
# BASE_URL: Your target URL
# SLACK_WEBHOOK_URL: Slack notification webhook
```

---

## Configuration

### Thresholds

Edit `scripts/performance_analyzer.py` untuk customize thresholds:

```python
THRESHOLDS = {
    "LCP": {"good": 2500, "needs_improvement": 4000},
    "FID": {"good": 100, "needs_improvement": 300},
    "CLS": {"good": 0.1, "needs_improvement": 0.25}
}
```

### WCAG Criteria

Edit `scripts/accessibility_checker.py` untuk customize WCAG checks:

```python
WCAG_CRITERIA = {
    "color_contrast": {"level": "AA", "criterion": "1.4.3"},
    "keyboard_navigation": {"level": "A", "criterion": "2.1.1"}
}
```

---

## Troubleshooting

### Playwright not found

```bash
npx playwright install --with-deps chromium
```

### Python dependencies

```bash
pip install --upgrade playwright openpyxl python-docx pillow
```

### Permission denied

```bash
chmod +x scripts/*.py
```

---

## License

MIT

---

## Contributing

1. Fork repository
2. Create feature branch
3. Commit changes
4. Push to branch
5. Create Pull Request

---

## Support

- GitHub Issues: [Report Bug](https://github.com/your-repo/qa-ultra-tester/issues)
- Documentation: [Wiki](https://github.com/your-repo/qa-ultra-tester/wiki)
