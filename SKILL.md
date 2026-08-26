---
name: qa-ultra-tester
description: Ultimate E2E QA testing skill with UI testing, API validation, accessibility (WCAG), cross-browser, database validation, video recording, auto bug reporting, test data generation, performance budget, flaky test detection, and advanced chaos engineering. Generates professional reports with embedded screenshots.
---

# QA ULTRA TESTER - Enterprise-Grade E2E Testing Skill

Skill pengujian kualitas perangkat lunak **TINGKAT ULTRA** yang menggabungkan UI testing, API validation, accessibility compliance, cross-browser testing, database validation, video recording, auto bug reporting, test data generation, performance budgeting, flaky test detection, dan chaos engineering lanjutan.

---

## 1. CORE TESTING LAYERS

### Layer 1: UI/UX Testing (Interactive Coverage)

**MANDATORY: 100% Element Interaction Coverage**

```
SETIAP Tombol:
  - Klik semua tombol (primary, secondary, icon buttons)
  - Hover states, focus states, disabled states
  - Double-click actions
  - Right-click context menu (jika ada)

SETIAP Input:
  - Text: kosong, spasi, 1 char, max char, special chars, XSS, SQLi
  - Number: negatif, 0, desimal, max int, overflow
  - Email: valid, invalid, tanpa @, tanpa domain, duplicate
  - Password: weak, medium, strong, match/mismatch confirmation
  - Date: past, today, future, min/max range, invalid format
  - File: valid type, invalid type, oversized, empty file, multiple files
  - Dropdown: single select, multi select, search/filter, clear selection

SETIAP Navigation:
  - Buka semua tab dan sub-tab
  - Expand/collapse semua accordion
  - Buka semua modal dan tutup (X, button, ESC, backdrop click)
  - Navigate semua breadcrumb links
```

### Layer 2: API/Network Validation

**INTERCEPT & VALIDATE SETIAP REQUEST**

```javascript
// Intercept semua network request
await page.route('**/*', async route => {
  const request = route.request();
  const response = await route.fetch();

  // Log untuk analysis
  apiLogs.push({
    url: request.url(),
    method: request.method(),
    status: response.status(),
    duration: Date.now() - startTime,
    payload: request.postData(),
    response: await response.text()
  });

  await route.continue();
});

// VALIDASI WAJIB:
// 1. Response status (200, 201, 400, 401, 403, 404, 500)
// 2. Response time < 2000ms (warning) / < 500ms (optimal)
// 3. Response payload structure (schema validation)
// 4. CORS headers present
// 5. No sensitive data leaked in response
// 6. Content-Type correct
// 7. No broken API endpoints (404)
// 8. No server errors (5xx)
```

**API Testing Matrix:**

| Endpoint | Method | Test Case | Expected |
|----------|--------|-----------|----------|
| /api/login | POST | Valid credentials | 200 + token |
| /api/login | POST | Invalid password | 401 |
| /api/login | POST | Empty body | 400 |
| /api/users | GET | Authenticated | 200 + data |
| /api/users | GET | No token | 401 |
| /api/users/:id | DELETE | Non-admin | 403 |

### Layer 3: Accessibility Testing (WCAG 2.1 AA)

**SCAN SETIAP HALAMAN**

```markdown
## Color Contrast
- Text contrast ratio >= 4.5:1 (normal text)
- Text contrast ratio >= 3:1 (large text)
- Non-text contrast >= 3:1 (UI components)

## Keyboard Navigation
- Semua interactive elements reachable via Tab
- Focus order logical (top→bottom, left→right)
- Focus visible (tidak invisible)
- Skip navigation link present
- No keyboard traps

## ARIA & Semantics
- Semua images punya alt text (atau alt="" untuk decorative)
- Form inputs punya associated labels
- Error messages linked to inputs (aria-describedby)
- Buttons punya accessible names
- Headings hierarchy correct (h1 → h2 → h3)

## Screen Reader
- Landmarks defined (main, nav, banner, contentinfo)
- Dynamic content announced (aria-live)
- Tables punya th and scope attributes
```

**Tool: axe-core / Lighthouse Accessibility Audit**

### Layer 4: Cross-Browser & Device Testing

**TEST DI SEMUA BROWSER**

```markdown
## Desktop Browsers
- Chromium (Playwright default)
- Firefox (playwright install firefox)
- WebKit/Safari (playwright install webkit)

## Mobile Viewports
- iPhone 14: 390x844
- iPhone 14 Pro Max: 430x932
- Samsung Galaxy S23: 360x780
- iPad: 768x1024
- iPad Pro: 1024x1366

## Responsive Breakpoints
- Mobile: < 768px
- Tablet: 768px - 1024px
- Desktop: > 1024px
- Wide: > 1440px

## Test Per Browser
- Layout tidak broken
- Tidak ada horizontal scroll
- Touch targets >= 44x44px (mobile)
- Font readable tanpa zoom
```

### Layer 5: Database Validation

**VERIFY DATA INTEGRITY**

```markdown
## Setelah Operasi CRUD
1. CREATE: Query database, pastikan data tersimpan dengan benar
2. READ: Bandingkan data UI vs database
3. UPDATE: Verify semua field terupdate
4. DELETE: Cek soft delete (flag) atau hard delete (hapus)

## Query Patterns
-- Verify data exists
SELECT * FROM users WHERE email = 'test@example.com';

-- Verify data integrity
SELECT COUNT(*) FROM orders WHERE user_id = ? AND status = 'active';

-- Check for orphaned records
SELECT * FROM order_items WHERE order_id NOT IN (SELECT id FROM orders);

-- Verify cascade delete
SELECT * FROM comments WHERE post_id = ?; -- Should be empty after post delete
```

### Layer 6: Video Recording

**RECORD FULL TEST SESSION**

```javascript
// Mulai recording
const video = await page.video();

// ... jalankan semua test ...

// Simpan video
await page.close();
const path = await video.path();
// Video tersimpan otomatis di test-results/

// Atau record manual dengan screenshot berurutan
const frames = [];
for (let i = 0; i < testSteps.length; i++) {
  await testSteps[i]();
  frames.push(await page.screenshot());
}
// Gabungkan frames jadi video (gunakan ffmpeg)
```

### Layer 7: Test Data Generator

**SYNTHETIC DATA UNTUK TESTING**

```markdown
## Data Types
-姓名: Faker.name.findName()
- Email: Faker.internet.email()
- Phone: Faker.phone.phoneNumber()
- Address: Faker.address.streetAddress()
- Company: Faker.company.companyName()
- Text: Faker.lorem.paragraphs(3)
- Number: Faker.random.number({ min: 1, max: 1000 })
- Date: Faker.date.between('2020-01-01', '2025-12-31')
- Avatar: Faker.image.avatar()

## Boundary Data
- Empty string: ""
- Single char: "a"
- Max length: string repeated 10000 times
- Special chars: "!@#$%^&*()_+-=[]{}|;':\",./<>?"
- Unicode: "测试中文", "العربية", "🔴🟡🟢"
- SQL injection: "' OR '1'='1'; DROP TABLE users;--"
- XSS: "<script>alert('XSS')</script>"
- Path traversal: "../../../etc/passwd"

## Invalid Patterns
- Email tanpa @: "userexample.com"
- Email tanpa domain: "user@"
- Phone berisi huruf: "abc123"
- Date format salah: "2025/13/45"
- Number negatif: "-123"
- Overflow: "99999999999999999"
```

### Layer 8: Performance Budget

**CORE WEB VITALS TRACKING**

```markdown
## Metrics Threshold
- LCP (Largest Contentful Paint): < 2.5s (Good), < 4.0s (Needs Improvement)
- FID (First Input Delay): < 100ms (Good), < 300ms (Needs Improvement)
- CLS (Cumulative Layout Shift): < 0.1 (Good), < 0.25 (Needs Improvement)
- TTFB (Time to First Byte): < 800ms
- Total Page Weight: < 3MB
- JavaScript Bundle: < 500KB
- CSS Bundle: < 200KB

## Measurement
await page.evaluate(() => {
  const paint = performance.getEntriesByType('paint');
  const navigation = performance.getEntriesByType('navigation')[0];

  return {
    ttfb: navigation.responseStart,
    domComplete: navigation.domComplete,
    loadComplete: navigation.loadEventEnd,
    firstPaint: paint.find(p => p.name === 'first-paint')?.startTime,
    firstContentfulPaint: paint.find(p => p.name === 'first-contentful-paint')?.startTime
  };
});
```

### Layer 9: Flaky Test Detection

**RETRY & TRACK RELIABILITY**

```markdown
## Retry Protocol
1. Test gagal → Retry max 3x dengan delay 2 detik
2. Jika pass di retry → Mark sebagai FLAKY
3. Jika fail 3x → Mark sebagai GENUINE FAILURE

## Reliability Score
- Pass on first run: 100% reliability
- Pass on 2nd retry: 66% reliability
- Pass on 3rd retry: 33% reliability
- Never pass: 0% reliability (genuine bug)

## Flaky Patterns to Detect
- Timing-dependent assertions
- Network-dependent tests
- Non-deterministic animations
- Shared state between tests
- Race conditions

## Report Format
| Test Name | Runs | First Pass | Flaky Rate | Status |
|-----------|------|------------|------------|--------|
| Login Test | 10 | 8 | 20% | ⚠️ FLAKY |
| Dashboard Load | 10 | 10 | 0% | ✅ STABLE |
```

### Layer 10: Advanced Chaos Engineering

**BREAK THINGS ON PURPOSE**

```markdown
## Network Chaos
- Throttle ke 3G (400kbps, 2000ms latency)
- Simulate offline mode
- Random packet loss (5%, 10%, 25%)
- DNS timeout simulation
- Slow 3G (100kbps, 5000ms latency)

## Browser Chaos
- CPU throttle 4x slowdown
- Memory pressure simulation
- Clear cookies mid-session
- Block specific domains (analytics, CDN)
- Disable JavaScript randomly

## Application Chaos
- Refresh halaman saat loading
- Back button saat form terisi
- Multiple tabs same session
- Rapid modal open/close
- Concurrent form submissions

## Implementation
// Network throttling
const client = await page.context().newCDPSession(page);
await client.send('Network.emulateNetworkConditions', {
  offline: false,
  downloadThroughput: (400 * 1024) / 8,  // 400kbps
  uploadThroughput: (400 * 1024) / 8,
  latency: 2000  // 2000ms
});

// CPU throttle
await client.send('Emulation.setCPUThrottlingRate', { rate: 4 });

// Clear storage mid-test
await page.evaluate(() => {
  localStorage.clear();
  sessionStorage.clear();
});
await page.reload();
```

---

## 2. MULTI-MODE INPUT PROCESSING

| Mode | Input | Workflow |
|------|-------|----------|
| **Mode 1: Script + Spec + URL** | Test script + FSD/PRD + URL | Eksekusi script, validasi vs spesifikasi, tambah negative tests |
| **Mode 2: Spec-Driven** | FSD/PRD/User Guide + URL | Bedah business rules, buat test matrix, gap analysis |
| **Mode 3: Whitebox** | Source code + URL | Analisis kode, trace API, test live |
| **Mode 4: Blackbox** | URL saja | Eksplorasi mandiri, identifikasi semua fitur |
| **Mode 5: API-First** | API docs/Swagger + URL | Test semua endpoint, validasi schema, performance |
| **Mode 6: Regression** | Previous test results + URL | Bandingkan dengan baseline, deteksi regression |

---

## 3. BUG REPORTING FORMAT

```markdown
### [🔴 CRITICAL] BUG-001: SQL Injection Vulnerable on Login Field

**Lokasi:** https://example.com/login - Email Input Field
**Severity:** CRITICAL
**Type:** Security / Injection
**CVSS Score:** 9.8

**Langkah Reproduksi:**
1. Buka https://example.com/login
2. Input `' OR '1'='1'; DROP TABLE users;--` di field email
3. Input apa saja di field password
4. Klik tombol Login

**Kondisi Aktual:**
- Login berhasil tanpa valid credential
- Database table users terhapus
- Error message menampilkan stack trace

**Kondisi Harapan:**
- Login gagal dengan message "Invalid credentials"
- Input di-sanitize dan di-escape
- Tidak ada error message yang leak implementation detail

**Bukti:**
- Screenshot: `screenshots/bug-001-login.png`
- Video: `videos/bug-001-login.mp4`
- API Log: `api-logs/bug-001-login.json`

**Rekomendasi Solusi:**
1. Implement parameterized queries / prepared statements
2. Add input validation dan sanitization
3. Implement rate limiting pada login endpoint
4. Generic error messages saja (jangan display stack trace)

**WCAG Impact:** None (Security issue)
**Cross-browser:** Reproducible di semua browser
```

---

## 4. OUTPUT ARTIFACTS

```
hasil-test/
├── QA_TEST_REPORT.docx           # Laporan Word bergambar
├── QA_TEST_REPORT.html           # Laporan HTML interaktif
├── test-results.json             # Machine-readable results
├── screenshots/                  # Bukti visual
│   ├── bug-001-login.png
│   ├── bug-002-dashboard.png
│   └── baseline-homepage.png
├── videos/                       # Recording test sessions
│   ├── test-session-001.mp4
│   └── bug-replay-001.mp4
├── api-logs/                     # Network/API logs
│   ├── all-requests.json
│   └── errors-only.json
├── tests/                        # Playwright scripts
│   ├── smoke.spec.ts
│   ├── regression.spec.ts
│   ├── api-validation.spec.ts
│   └── accessibility.spec.ts
├── ci-cd/                        # Pipeline templates
│   ├── playwright-e2e.yml
│   ├── gitlab-ci.yml
│   └── jenkins-pipeline.groovy
├── test-data/                    # Generated test data
│   ├── users.json
│   └── fixtures.json
├── performance/                  # Performance reports
│   ├── lighthouse-report.json
│   └── core-web-vitals.json
├── accessibility/                # WCAG reports
│   └── axe-report.json
└── flaky-analysis/               # Flaky test tracking
    └── reliability-scores.json
```

---

## 5. EXECUTION PROTOCOL

```markdown
## Phase 1: Discovery (5-10 menit)
1. Buka target URL
2. Snapshot semua halaman
3. Identify semua interactive elements
4. Map navigation structure
5. Identify forms, modals, tabs

## Phase 2: Happy Path Testing (30-40%)
1. Test semua happy path scenarios
2. Verify core user flows
3. Capture baseline screenshots
4. Record network traffic

## Phase 3: Negative Testing (25-30%)
1. Invalid inputs semua field
2. Boundary testing
3. Error message validation
4. Form validation testing

## Phase 4: Security Testing (10-15%)
1. XSS injection semua input
2. SQL injection semua input
3. Path traversal testing
4. Authentication/authorization testing

## Phase 5: Accessibility Testing (10%)
1. axe-core scan
2. Keyboard navigation test
3. Color contrast check
4. ARIA validation

## Phase 6: Performance Testing (5-10%)
1. Core Web Vitals measurement
2. Page load timing
3. Resource size analysis
4. API response time

## Phase 7: Chaos Testing (5-10%)
1. Network throttling
2. Rapid interactions
3. Session disruption
4. Browser stress test

## Phase 8: Reporting (10%)
1. Aggregate all results
2. Generate test reports
3. Create CI/CD templates
4. Document recommendations
```

---

## 6. QUALITY GATES

```markdown
## BLOCKER (Tidak boleh rilis)
- Security vulnerabilities (XSS, SQLi, CSRF)
- Data loss scenarios
- Authentication bypass
- Server errors (500)
- Broken critical user flows

## HIGH (Harus fix sebelum rilis)
- Form validation bypass
- Broken non-critical features
- Missing error handling
- Performance > 5 detik

## MEDIUM (Fix setelah rilis)
- UX inconsistencies
- Minor visual bugs
- Accessibility violations (WCAG AA)
- Flaky tests > 20%

## LOW (Backlog)
- Cosmetic issues
- Documentation gaps
- Code improvements
```

---

## 7. AUTOMATION COMMANDS

| Command | Description |
|---------|-------------|
| `/qa-test <url>` | Full E2E testing |
| `/qa-api <url>` | API-focused testing |
| `/qa-a11y <url>` | Accessibility audit |
| `/qa-perf <url>` | Performance testing |
| `/qa-security <url>` | Security testing |
| `/qa-regression <url>` | Regression testing |
| `/qa-chaos <url>` | Chaos engineering |
| `/qa-report` | Generate reports |
| `/qa-fix` | Auto-fix suggestions |

---

## 8. TOOLS INTEGRATION

```markdown
## Required
- Playwright MCP (browser automation)
- axe-core (accessibility scanning)
- Lighthouse (performance audit)

## Optional
- OWASP ZAP (security scanning)
- k6 / Artillery (load testing)
- Sentry (error tracking)
- Percy / Applitools (visual regression)

## Database
- MySQL/PostgreSQL client for data validation
- MongoDB client for NoSQL validation
```

---

## 9. REPORT TEMPLATE

```markdown
# QA TEST REPORT - [Project Name]

## Executive Summary
- Total Tests: XXX
- Passed: XXX (XX%)
- Failed: XXX (XX%)
- Flaky: XXX (XX%)
- Duration: XX minutes

## Quality Score: X/100

### Breakdown
- Functionality: XX/100
- Security: XX/100
- Accessibility: XX/100
- Performance: XX/100
- Cross-browser: XX/100

## Critical Issues
1. [BUG-001] SQL Injection on Login
2. [BUG-002] XSS on Search Field

## Recommendations
1. Fix all CRITICAL issues before release
2. Address HIGH issues within 1 week
3. Schedule MEDIUM issues for next sprint

## Production Readiness: ❌ NOT READY / ✅ READY
```

---

## 10. SKILL ACTIVATION

Skill ini aktif secara otomatis saat user meminta:
- "Testing E2E untuk [URL]"
- "QA testing [URL]"
- "Audit website [URL]"
- "Check accessibility [URL]"
- "Security test [URL]"
- "Performance test [URL]"
- Atu command `/qa-test`, `/qa-api`, `/qa-a11y`, dll.
