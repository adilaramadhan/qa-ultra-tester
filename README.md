# QA Ultra Tester

**Enterprise-Grade E2E Testing Skill for OpenCode**

Skill pengujian kualitas perangkat lunak tingkat ultra yang menggabungkan UI testing, API validation, accessibility compliance, cross-browser testing, database validation, video recording, auto bug reporting, test data generation, performance budgeting, flaky test detection, dan chaos engineering lanjutan.

---

## 3 Trigger Modes

Skill ini otomatis aktif dan mengarahkan scope testing sesuai 3 trigger mode:

| Mode | Trigger Phrase / Command | Testing Scope & Layers | Output |
|---|---|---|---|
| **Mode 1: E2E Full Testing** | `"Testing E2E Menu [Nama Menu]..."`<br>`/qa-test` | **Semua 11 Layer Testing**<br>(UI/UX, API, a11y, Cross-Browser, DB, Video, Auto Bug, Data Gen, Perf, Flaky, Chaos) | Laporan Lengkap (HTML + DOCX), Playwright Specs lengkap (`smoke`, `negative`, `security`, `api`, `a11y`), Evidence, Logs |
| **Mode 2: Bug Hunter** | `"Bug hunter menu [Nama Menu]..."`<br>`/qa-bug-hunter` | **5 Focus Layers**<br>(UI/UX, Security Injection, Chaos/Boundary, Video Recording, Flaky Analysis) | Laporan Temuan Defect/Bug (HTML + DOCX), Security Repro Specs (`security.spec.ts`), Screenshots & Video |
| **Mode 3: API Testing** | `"API Testing menu [Nama Menu]..."`<br>`/qa-api` | **4 Backend Focus Layers**<br>(API/Network, DB Integrity, Perf Budget, Test Data Gen) | Laporan Validasi Endpoint (HTML + DOCX), API Test Specs (`api-validation.spec.ts`), Network Logs |

---

## Features & 11 Testing Layers

| Layer | Description |
|---|---|
| **Layer 1: UI/UX Testing** | 100% element interaction coverage (buttons, inputs, modals, navigation) |
| **Layer 2: API Validation** | Intercept & validate every network request, status, and payload |
| **Layer 3: Accessibility (WCAG)** | WCAG 2.1 AA compliance checking via axe-core & contrast scans |
| **Layer 4: Cross-Browser** | Chromium, Firefox, WebKit rendering & responsive viewports |
| **Layer 5: Database Validation** | CRUD data integrity & soft/hard delete verification |
| **Layer 6: Video Recording** | Full test session recording (`.webm` / `.mp4`) |
| **Layer 7: Test Data Generator** | Synthetic data with boundary, XSS, SQLi cases |
| **Layer 8: Performance Budget** | Google Core Web Vitals (LCP, FID, CLS, TTFB) tracking |
| **Layer 9: Flaky Test Detection** | Retry analysis (up to 3x) & reliability scoring |
| **Layer 10: Chaos Engineering** | Network throttling (3G/offline), CPU slowdown, storage clear |
| **Layer 11: Safety Boundaries** | Staging/sandbox safe execution constraints |

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
git clone https://github.com/adilaramadhan/qa-ultra-tester.git ~/.config/opencode/skills/qa-ultra-tester

# Install Python dependencies
pip install playwright openpyxl python-docx pillow

# Install npm dependencies & Playwright browsers
cd ~/.config/opencode/skills/qa-ultra-tester
npm install
npx playwright install --with-deps chromium firefox webkit
```

---

## Strict Reporting Contract & Artifacts

Semua hasil pengujian disimpan terpusat di folder target aktif (**Strict Co-Location**):

```
[Target Folder]/hasil-test/
├── QA_TEST_REPORT.docx           # Laporan Word bergambar (Times New Roman standar perusahaan)
├── QA_TEST_REPORT.html           # Laporan HTML interaktif (Dark/Light toggle, Lightbox Modal Popup)
├── test-results.json             # Single source of truth (JSON Schema validated)
├── screenshots/                  # Bukti visual pengujian & bug reproduction
├── videos/                       # Video rekaman sesi testing (.webm / .mp4)
├── api-logs/                     # Network/API intercept logs
└── tests/                        # Generated Playwright TypeScript test scripts
    ├── playwright.config.ts
    ├── smoke.spec.ts
    ├── negative.spec.ts
    ├── security.spec.ts
    ├── api-validation.spec.ts
    └── accessibility.spec.ts
```

---

## Scripts & CLI

### 1. Generate Report (HTML + DOCX)

Menghasilkan laporan interaktif HTML (dengan Lightbox modal & Core Web Vitals) serta DOCX dari `test-results.json`:

```bash
python scripts/generate_report.py [output_dir]
```

### 2. Generate TypeScript Test Scripts

Menghasilkan file test Playwright TypeScript (`.spec.ts`) otomatis dari `test-results.json`:

```bash
python scripts/generate_test_script.py [output_dir]
```

### 3. Generate Test Data

```bash
python scripts/generate_test_data.py
```

### 4. Analyze Performance

```bash
python scripts/performance_analyzer.py
```

### 5. Check Accessibility

```bash
python scripts/accessibility_checker.py
```

---

## Quality Gates

| Severity | Criteria | Action |
|---|---|---|
| 🔴 **CRITICAL** | Security (XSS, SQLi, Auth Bypass), Data Loss, Server 500 | Block release |
| 🟠 **HIGH** | Validation bypass, Broken non-critical flow, Perf > 5s | Fix before release |
| 🟡 **MEDIUM** | UX inconsistencies, Minor visual bugs, a11y violations | Fix after release |
| 🟢 **LOW** | Cosmetic issues, Typo, Documentation gaps | Backlog |

---

## CI/CD Integration

Templates CI/CD siap pakai tersedia untuk:
- GitHub Actions: `ci-cd/playwright-e2e.yml`
- GitLab CI: `ci-cd/.gitlab-ci.yml`
- Jenkins: `ci-cd/jenkins-pipeline.groovy`

---

## License

MIT © [Adila Ramadhan](https://github.com/adilaramadhan)
