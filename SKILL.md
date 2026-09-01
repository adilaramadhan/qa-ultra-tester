---
name: qa-ultra-tester
description: Ultimate E2E QA testing skill with 3 trigger modes (Testing E2E, Bug Hunter, API Testing). Includes UI testing, API validation, accessibility (WCAG), cross-browser, database validation, video recording, auto bug reporting, test data generation, performance budget, flaky test detection, and advanced chaos engineering. Generates professional reports with embedded screenshots.
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

### Layer 11: Safety Boundaries

**WAJIB DIPATUHI UNTUK CHAOS & SECURITY TESTING**

```markdown
## Batasan Keamanan
- HANYA jalankan security test (SQLi, XSS, path traversal) di environment NON-PRODUCTION
- WAJIB konfirmasi ke user sebelum menjalankan destructive test
- Jangan pernah execute DROP/DELETE/TRUNCATE query di database production
- Rate limit semua brute-force test (max 10 req/detik)
- Jangan capture/log credential yang valid — hanya test payload

## Batasan Chaos
- CPU throttle max 6x (jangan freeze browser)
- Network throttle boleh offline, tapi max 30 detik per sesi
- Jangan clear storage di production tanpa konfirmasi user
- CDP session: SATU context per page, jangan share antar tab
```

---

## 2. TRIGGER MODES (3 MODE UTAMA)

Skill ini diaktifkan oleh **3 trigger utama**. Setiap trigger menentukan layer mana yang dieksekusi agar hemat token dan fokus.

### Mode 1: Testing E2E Menu [nama_menu]

**Trigger:** `Testing E2E menu ...` / `Testing E2E fitur ...`

**Deskripsi:** Full end-to-end testing. Semua layer diaktifkan.

| Layer | Status |
|-------|--------|
| Layer 1: UI/UX Testing | AKTIF |
| Layer 2: API/Network Validation | AKTIF |
| Layer 3: Accessibility (WCAG) | AKTIF |
| Layer 4: Cross-Browser | AKTIF |
| Layer 5: Database Validation | AKTIF |
| Layer 6: Video Recording | AKTIF |
| Layer 7: Test Data Generator | AKTIF |
| Layer 8: Performance Budget | AKTIF |
| Layer 9: Flaky Test Detection | AKTIF |
| Layer 10: Chaos Engineering | AKTIF |
| Layer 11: Safety Boundaries | AKTIF |

**Execution Phases:** Phase 1-9 (semua)

---

### Mode 2: Bug Hunter Menu [nama_menu]

**Trigger:** `Bug hunter menu ...` / `Bug hunting ...` / `Cari bug menu ...`

**Deskripsi:** Fokus eksplorasi mencari bug. Interaksi UI intensif, inject security payload, stress test. TIDAK melakukan API schema validation, accessibility audit, cross-browser, atau database validation.

| Layer | Status |
|-------|--------|
| Layer 1: UI/UX Testing | AKTIF |
| Layer 2: API/Network Validation | SKIP |
| Layer 3: Accessibility (WCAG) | SKIP |
| Layer 4: Cross-Browser | SKIP |
| Layer 5: Database Validation | SKIP |
| Layer 6: Video Recording | AKTIF |
| Layer 7: Test Data Generator | AKTIF (boundary data only) |
| Layer 8: Performance Budget | SKIP |
| Layer 9: Flaky Test Detection | AKTIF |
| Layer 10: Chaos Engineering | AKTIF |
| Layer 11: Safety Boundaries | AKTIF |

**Execution Phases:**
- Phase 1: Discovery
- Phase 2: Happy Path (singkat, hanya core flow)
- Phase 3: Negative Testing (INTENSIF — fokus cari bug)
- Phase 4: Security Testing (INTENSIF)
- Phase 7: Chaos Testing (INTENSIF)
- Phase 8: Generate Test Scripts
- Phase 9: Reporting

---

### Mode 3: API Testing [nama_endpoint/menu]

**Trigger:** `API Testing ...` / `Test API ...`

**Deskripsi:** Fokus pengujian API/network. Intercept semua request, validasi schema, response time, error handling. TIDAK melakukan UI interaction, accessibility, cross-browser, video recording, atau chaos engineering.

| Layer | Status |
|-------|--------|
| Layer 1: UI/UX Testing | SKIP |
| Layer 2: API/Network Validation | AKTIF (INTENSIF) |
| Layer 3: Accessibility (WCAG) | SKIP |
| Layer 4: Cross-Browser | SKIP |
| Layer 5: Database Validation | AKTIF |
| Layer 6: Video Recording | SKIP |
| Layer 7: Test Data Generator | AKTIF |
| Layer 8: Performance Budget | AKTIF (API response time focus) |
| Layer 9: Flaky Test Detection | AKTIF |
| Layer 10: Chaos Engineering | SKIP |
| Layer 11: Safety Boundaries | AKTIF |

**Execution Phases:**
- Phase 1: Discovery (API endpoint mapping)
- Phase 2: Happy Path (valid request/response)
- Phase 3: Negative Testing (invalid payload, missing fields, wrong types)
- Phase 4: Security Testing (auth bypass, token manipulation)
- Phase 6: Performance Testing (response time, throughput)
- Phase 8: Generate Test Scripts
- Phase 9: Reporting

---

## 3. BUG REPORTING FORMAT

**PENTING:** Format ini WAJIB konsisten dengan JSON schema di Section 4. Setiap bug ditulis ke `test-results.json` sebagai objek JSON, lalu script `generate_report.py` yang mencetak ke HTML/DOCX. Agent TIDAK BOLEH membuat tabel markdown ad-hoc.

```json
{
  "id": "BUG-001",
  "severity": "CRITICAL",
  "title": "SQL Injection Vulnerable on Login Field",
  "location": "https://example.com/login - Email Input Field",
  "steps": [
    "1. Buka https://example.com/login",
    "2. Input `' OR '1'='1'; DROP TABLE users;--` di field email",
    "3. Input apa saja di field password",
    "4. Klik tombol Login"
  ],
  "actual": "Login berhasil tanpa valid credential. Database table users terhapus. Error message menampilkan stack trace.",
  "expected": "Login gagal dengan message 'Invalid credentials'. Input di-sanitize. Tidak ada error leak.",
  "evidence": {
    "screenshot": "screenshots/bug-001-login.png",
    "video": "videos/bug-001-login.mp4",
    "network_log": "api-logs/bug-001-login.json"
  },
  "recommendation": "1. Implement parameterized queries. 2. Add input validation. 3. Implement rate limiting. 4. Generic error messages."
}
```

---

## 4. STRICT REPORTING CONTRACT (SINGLE SOURCE OF TRUTH)

**ATURAN MUTLAK:** Semua hasil testing WAJIB ditulis ke `hasil-test/test-results.json` dengan skema JSON ketat di bawah. Laporan HTML/DOCX/Markdown HANYA boleh di-generate dari file JSON ini via `python scripts/generate_report.py`. Agent DILARANG membuat format tabel/laporan kustom sendiri.

### 4.1. JSON Schema — `test-results.json`

```json
{
  "project_name": "string",
  "testing_mode": "e2e|bug_hunter|api",
  "target_url": "string",
  "test_date": "YYYY-MM-DD",
  "tester": "string",
  "generated_at": "ISO-8601",
  "quality_score": 0,
  "video_recording": "videos/test-session.webm",
  "summary": {
    "total": 0,
    "passed": 0,
    "failed": 0,
    "flaky": 0,
    "duration_ms": 0
  },
  "bugs": [
    {
      "id": "BUG-001",
      "severity": "CRITICAL|HIGH|MEDIUM|LOW",
      "priority": "P1|P2|P3|P4",
      "type": "Functional|Security|UI/UX|Validation",
      "title": "string",
      "location": "string",
      "steps": ["step 1", "step 2"],
      "actual": "string",
      "expected": "string",
      "evidence": {
        "screenshot": "screenshots/bug-001.png",
        "video": "videos/bug-001.webm",
        "network_log": "api-logs/bug-001.json"
      },
      "recommendation": "string"
    }
  ],
  "test_results": [
    {
      "name": "string",
      "category": "happy_path|negative|boundary|security|api|a11y",
      "layer": "UI|API|A11Y|PERF|SECURITY|CHAOS",
      "status": "PASS|FAIL",
      "duration_ms": 0,
      "browser": "Chromium|Firefox|WebKit",
      "steps": [
        "1. Buka halaman target",
        "2. Input data valid pada form",
        "3. Klik tombol Submit"
      ],
      "expected": "Form berhasil disimpan dan muncul notifikasi sukses",
      "actual": "Form tersimpan dan dialihkan ke halaman list",
      "retries": 0,
      "flaky": false,
      "details": "string",
      "screenshot": "screenshots/test-failed-name.png"
    }
  ],
  "performance": {
    "LCP": { "value": 0, "unit": "ms", "rating": "good|needs_improvement|poor" },
    "FID": { "value": 0, "unit": "ms", "rating": "good|needs_improvement|poor" },
    "CLS": { "value": 0, "unit": "", "rating": "good|needs_improvement|poor" },
    "TTFB": { "value": 0, "unit": "ms", "rating": "good|needs_improvement|poor" }
  },
  "accessibility": [
    {
      "type": "string",
      "element": "string",
      "description": "string",
      "wcag": "string",
      "severity": "critical|serious|moderate|minor"
    }
  ]
}
```

### 4.2. Workflow Pelaporan & Lokasi File

**ATURAN LOKASI FILE (STRICT CO-LOCATION):**
Semua output, file, folder, screenshot, video, script test, dan report **WAJIB dibuat dan disimpan HANYA di dalam folder kerja aktif** (working directory saat ini, misal `Testing menu Banner/hasil-test/`). DILARANG membuat file tercecer di luar direktori pengujian yang sedang aktif.

```
1. Jalankan semua test di dalam folder target aktif
2. Buat folder `hasil-test/` di DALAM folder target tersebut:
   [Current Target Folder]/hasil-test/
3. Semua screenshot masuk ke: [Current Target Folder]/hasil-test/screenshots/
4. Semua video masuk ke: [Current Target Folder]/hasil-test/videos/
5. Tulis `test-results.json` di: [Current Target Folder]/hasil-test/test-results.json
6. Jalankan: python [path_skill]/scripts/generate_report.py "[Current Target Folder]/hasil-test"
   → Output: QA_TEST_REPORT.html
   → Output: QA_TEST_REPORT.docx
7. Jalankan: python [path_skill]/scripts/generate_test_script.py "[Current Target Folder]/hasil-test"
   → Output: [Current Target Folder]/hasil-test/tests/*.spec.ts
```

### 4.3. Output Artifacts (Terpusat di 1 Folder)

```
[Current Target Folder]/hasil-test/
├── test-results.json             # SATU-SATUNYA sumber data (JSON ketat)
├── QA_TEST_REPORT.html           # HTML report kustom premium (gambar & video embedded)
├── QA_TEST_REPORT.docx           # Laporan Word bergambar profesional
├── screenshots/                  # Semua bukti screenshot (bug, error, failure)
│   ├── bug-001-*.png
│   ├── bug-002-*.png
│   └── fail-*.png
├── videos/                       # Video rekaman sesi testing
│   └── test-session.webm
├── tests/                        # Playwright TypeScript test scripts
│   ├── smoke.spec.ts
│   ├── negative.spec.ts
│   ├── security.spec.ts
│   ├── accessibility.spec.ts
│   ├── performance.spec.ts
│   ├── chaos.spec.ts
│   ├── api-validation.spec.ts
│   └── playwright.config.ts
└── api-logs/                     # Network request & response logs (jika ada)
    └── all-requests.json
```

---

## 5. EXECUTION PROTOCOL

**Phase yang dijalankan tergantung trigger mode (lihat Section 2).** Berikut daftar lengkap phase:

```markdown
## Phase 1: Discovery (5-10 menit) — ALL MODES
1. Buka target URL
2. Snapshot semua halaman
3. Identify semua interactive elements (E2E & Bug Hunter) ATAU API endpoints (API Testing)
4. Map navigation structure
5. Identify forms, modals, tabs

## Phase 2: Happy Path Testing — E2E: 30-40% | Bug Hunter: 10% (core flow saja) | API: 20%
1. Test semua happy path scenarios
2. Verify core user flows
3. Capture baseline screenshots
4. Record network traffic

## Phase 3: Negative Testing — E2E: 25-30% | Bug Hunter: 35% (INTENSIF) | API: 30%
1. Invalid inputs semua field
2. Boundary testing
3. Error message validation
4. Form validation testing

## Phase 4: Security Testing — E2E: 10-15% | Bug Hunter: 25% (INTENSIF) | API: 15%
1. XSS injection semua input (NON-PRODUCTION ONLY)
2. SQL injection semua input (NON-PRODUCTION ONLY)
3. Path traversal testing
4. Authentication/authorization testing

## Phase 5: Accessibility Testing — E2E: 10% | Bug Hunter: SKIP | API: SKIP
1. axe-core scan
2. Keyboard navigation test
3. Color contrast check
4. ARIA validation

## Phase 6: Performance Testing — E2E: 5-10% | Bug Hunter: SKIP | API: 15%
1. Core Web Vitals measurement
2. Page load timing
3. Resource size analysis
4. API response time

## Phase 7: Chaos Testing — E2E: 5-10% | Bug Hunter: 20% (INTENSIF) | API: SKIP
1. Network throttling
2. Rapid interactions
3. Session disruption
4. Browser stress test

## Phase 8: Generate TypeScript Test Scripts (WAJIB — ALL MODES)
1. Generate Playwright TypeScript spec files di `hasil-test/tests/`
2. Setiap layer testing WAJIB punya file `.spec.ts` sendiri
3. Script WAJIB menggunakan TypeScript (bukan JavaScript)
4. Gunakan `python scripts/generate_test_script.py` untuk scaffolding
5. Struktur file:
   - `smoke.spec.ts` — happy path
   - `negative.spec.ts` — negative/boundary
   - `security.spec.ts` — security injection
   - `accessibility.spec.ts` — axe-core WCAG (E2E only)
   - `performance.spec.ts` — Core Web Vitals (E2E & API only)
   - `chaos.spec.ts` — chaos engineering (E2E & Bug Hunter only)
   - `api-validation.spec.ts` — API intercept & validation (E2E & API only)

## Phase 9: Reporting (WAJIB via Pipeline — ALL MODES)
1. WAJIB simpan video rekaman ke `hasil-test/videos/` atau path root `hasil-test/`
2. WAJIB ambil screenshot untuk:
   - Setiap bug temuan (`hasil-test/screenshots/bug-xxx.png`)
   - Setiap test case yang FAILED / ERROR (`hasil-test/screenshots/fail-xxx.png`)
   - Ketidaksesuaian UI/UX / anomaly visual
3. Tulis SEMUA data ke `hasil-test/test-results.json` (JSON ketat sesuai Section 4)
4. Jalankan `python scripts/generate_report.py`
5. Output: `QA_TEST_REPORT.html` (HTML kustom premium) + `QA_TEST_REPORT.docx`
6. DILARANG menulis laporan dalam format lain di luar pipeline ini
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

| Command | Mode | Description |
|---------|------|-------------|
| `Testing E2E menu [X]` | E2E | Full testing semua layer |
| `Bug hunter menu [X]` | Bug Hunter | Fokus cari bug (UI, Security, Chaos) |
| `API Testing [X]` | API | Fokus API validation & performance |
| `/qa-report` | — | Generate reports dari test-results.json |
| `/qa-generate-scripts` | — | Generate TypeScript spec files dari test-results.json |

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

Laporan di-generate HANYA oleh `python scripts/generate_report.py` dari `test-results.json`.

**HTML Report** menggunakan desain kustom premium (BUKAN bawaan Playwright reporter):
- Dashboard dengan quality score gauge, summary cards, severity breakdown chart
- Bug details lengkap: steps reproduksi, evidence (screenshot embedded base64), recommendation
- Test matrix tabel per-layer (UI/API/A11Y/PERF/SECURITY/CHAOS)
- Performance Core Web Vitals dengan bar chart warna (good/needs_improvement/poor)
- Accessibility violations dengan WCAG criterion
- Flaky test analysis
- Dark/light mode toggle, responsive, printable

**DOCX Report** menggunakan template profesional terstandarisasi:
- Cover page: Judul (Times New Roman 22pt bold), Subtitle (16pt), Tabel info dengan background (Target URL, Fitur Uji, Tanggal, Metodologi, Status Kesiapan)
- Section headings: Times New Roman 14pt bold
- Body text & tabel: Times New Roman 12pt
- Header tabel: Dark Blue (#1B3A5C) dengan teks putih
- Zebra rows (abu-abu muda) pada baris genap tabel
- Severity color-coded: Critical (Merah), High (Orange), Medium (Kuning Tua), Low (Hijau)
- Screenshot bukti visual di-embed langsung di dokumen dengan caption bernomor
- Struktur: Cover -> Ringkasan Eksekutif -> Statistik Bug -> Rincian Bug -> Hasil Pengujian Fitur -> Performa -> Aksesibilitas -> Bukti Visual -> Rekomendasi Tindak Lanjut

**JSON Report** (`test-results.json`) adalah satu-satunya sumber data — semua format lain dibaca dari sini.

---

## 10. SKILL ACTIVATION

Skill ini **HANYA** diaktifkan oleh 3 trigger berikut:

### Trigger 1: Testing E2E
```
"Testing E2E menu [nama_menu]"
"Testing E2E fitur [nama_fitur]"
"Testing E2E [URL]"
```
→ Jalankan **SEMUA layer** (full coverage). Cocok untuk pre-release testing.

### Trigger 2: Bug Hunter
```
"Bug hunter menu [nama_menu]"
"Bug hunting [nama_menu/URL]"
"Cari bug menu [nama_menu]"
"Hunt bug [nama_menu]"
```
→ Jalankan layer **UI/UX, Security, Chaos, Video, Flaky, Test Data** saja. Fokus menemukan bug sebanyak mungkin.

### Trigger 3: API Testing
```
"API Testing [nama_endpoint/menu]"
"Test API [nama_endpoint/menu]"
"API test [URL]"
```
→ Jalankan layer **API Validation, Database, Performance, Test Data, Flaky** saja. Fokus endpoint, schema, response time.

### Contoh Prompt Lengkap (Optimal 1 Prompt)

```
Testing E2E menu Banner di https://example.com

Login:
- Inputer: NIK 12345 / Password: pass123
- Approver: NIK 67890 / Password: pass456

Lokasi: Menu Utama > Master Data > Banner
Scope: Input data baru, Edit, Delete, Approve, Reject, Unapprove
```

```
Bug hunter menu Banner Input di https://example.com
Login: NIK 12345 / Password: pass123
Fokus: form input, validasi file upload, boundary testing
```

```
API Testing menu User Management di https://example.com
Login: NIK admin01 / Password: admin123
Fokus: CRUD endpoint, auth token, response schema
```
