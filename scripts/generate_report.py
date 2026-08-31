#!/usr/bin/env python3
"""
QA Ultra Tester - Report Generator v2.1
Generates premium HTML and Word (.docx) reports from test-results.json
Single source of truth: hasil-test/test-results.json

Changes v2.1:
- Fixed dark/light toggle
- Fixed double numbering in Steps to Reproduce
- Added video embed (test recording)
- Added screenshot on failed/bug test results
- Performance metrics now show full names + descriptions
- Duration converted from ms to human-readable (min:sec)
"""

import os
import sys
import json
import base64
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ── Performance metric descriptions ──────────────────────────
METRIC_INFO = {
    "LCP": {"full": "Largest Contentful Paint", "desc": "Waktu render elemen terbesar yang terlihat di viewport. Target: < 2.5s"},
    "FID": {"full": "First Input Delay", "desc": "Waktu dari interaksi pertama user sampai browser mulai proses. Target: < 100ms"},
    "CLS": {"full": "Cumulative Layout Shift", "desc": "Skor stabilitas visual layout halaman. Target: < 0.1"},
    "TTFB": {"full": "Time to First Byte", "desc": "Waktu dari request sampai byte pertama response diterima. Target: < 800ms"},
    "FCP": {"full": "First Contentful Paint", "desc": "Waktu render konten pertama (teks/gambar) di layar. Target: < 1.8s"},
    "INP": {"full": "Interaction to Next Paint", "desc": "Latensi terburuk dari semua interaksi user selama halaman terbuka. Target: < 200ms"},
    "TBT": {"full": "Total Blocking Time", "desc": "Total waktu main thread ter-blok antara FCP dan Time to Interactive. Target: < 200ms"},
    "SI": {"full": "Speed Index", "desc": "Seberapa cepat konten terlihat selama page load. Target: < 3.4s"},
    "TTI": {"full": "Time to Interactive", "desc": "Waktu sampai halaman fully interactive. Target: < 3.8s"},
}


def format_duration(ms):
    """Convert milliseconds to human-readable format."""
    if ms is None or ms == 0:
        return "0s"
    total_seconds = ms / 1000
    if total_seconds < 1:
        return f"{ms}ms"
    minutes = int(total_seconds // 60)
    seconds = total_seconds % 60
    if minutes == 0:
        return f"{seconds:.1f}s"
    return f"{minutes}m {seconds:.1f}s"


def strip_step_number(step_text):
    """Remove leading number prefix like '1. ' or '1) ' from step text."""
    import re
    return re.sub(r'^\d+[\.\)\-]\s*', '', str(step_text))


class QAReportGenerator:
    def __init__(self, output_dir="hasil-test"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.data = {}

    def load_from_json(self, filename="test-results.json"):
        """Load all data from the single source of truth JSON file."""
        filepath = self.output_dir / filename
        if not filepath.exists():
            print(f"Error: {filepath} not found. Run tests first.")
            return False
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            self.data = json.load(f)
        return True

    def set_data(self, data: dict):
        """Set data directly (for programmatic use)."""
        self.data = data

    # ── helpers ──────────────────────────────────────────────
    def _bugs(self):
        return self.data.get("bugs", [])

    def _tests(self):
        return self.data.get("test_results", [])

    def _perf(self):
        return self.data.get("performance_metrics", self.data.get("performance", {}))

    def _a11y(self):
        return self.data.get("accessibility", [])

    def _summary(self):
        return self.data.get("summary", {})

    def _quality_score(self):
        return self.data.get("quality_score", 0)

    def _embed_file(self, path_str, file_type="image"):
        """Embed file as base64 for HTML report. Supports image and video with multi-path lookup."""
        if not path_str:
            return None
        
        candidates = []
        p = Path(path_str)
        if p.is_absolute():
            candidates.append(p)
        else:
            candidates.append(self.output_dir / p)
            candidates.append(self.output_dir / "screenshots" / p.name)
            candidates.append(self.output_dir / "videos" / p.name)
            candidates.append(self.output_dir.parent / p)
            candidates.append(Path.cwd() / p)

        target_file = None
        for c in candidates:
            if c.exists() and c.is_file():
                target_file = c
                break

        if not target_file:
            return None

        ext = target_file.suffix.lower().lstrip(".")
        if file_type == "video":
            mime = {"mp4": "video/mp4", "webm": "video/webm", "ogg": "video/ogg",
                    "avi": "video/x-msvideo", "mov": "video/quicktime"}.get(ext, "video/mp4")
        else:
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                    "webp": "image/webp", "gif": "image/gif"}.get(ext, "image/png")
        try:
            with open(target_file, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            return f"data:{mime};base64,{b64}"
        except Exception:
            return None

    def _find_video(self):
        """Find test recording video file in multiple possible locations and verify existence."""
        video_path = self.data.get("video_recording") or self.data.get("video")
        if video_path:
            # Check if direct path or relative path exists
            candidates = [
                Path(video_path),
                self.output_dir / video_path,
                self.output_dir / Path(video_path).name,
                self.output_dir / "videos" / Path(video_path).name,
                self.output_dir.parent / video_path,
                self.output_dir.parent / "videos" / Path(video_path).name,
            ]
            for cand in candidates:
                if cand.exists() and cand.is_file() and cand.stat().st_size > 0:
                    return str(cand)
        
        search_dirs = [
            self.output_dir,
            self.output_dir / "videos",
            self.output_dir / "test-results",
            self.output_dir.parent / "test-results",
            self.output_dir.parent / "videos",
            Path.cwd() / "test-results",
        ]
        for sdir in search_dirs:
            if sdir.exists():
                for ext in ["*.webm", "*.mp4", "*.mov"]:
                    videos = [v for v in sdir.glob(ext) if v.is_file() and v.stat().st_size > 0]
                    if videos:
                        return str(videos[0])
                    # Recursive search inside subfolders
                    videos_rec = [v for v in sdir.rglob(ext) if v.is_file() and v.stat().st_size > 0]
                    if videos_rec:
                        return str(videos_rec[0])
        return None

    # ── DOCX (Professional Template) ────────────────────────
    def generate_docx(self, filename="QA_TEST_REPORT.docx"):
        if not HAS_DOCX:
            print("Warning: python-docx not installed. Skipping DOCX generation.")
            return None

        from docx.shared import Cm, Emu
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import qn, nsdecls
        from docx.oxml import parse_xml

        doc = Document()

        # ── Style setup ──
        BODY_FONT = 'Times New Roman'
        TITLE_FONT = 'Times New Roman'
        BODY_SIZE = 12
        TITLE_SIZE = 14

        style = doc.styles['Normal']
        style.font.name = BODY_FONT
        style.font.size = Pt(BODY_SIZE)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.space_before = Pt(2)

        DARK_BLUE = RGBColor(0x1B, 0x3A, 0x5C)
        HEADER_BG = "1B3A5C"
        HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
        LIGHT_GRAY_BG = "F2F2F2"
        GREEN = RGBColor(0x22, 0x8B, 0x22)
        RED = RGBColor(0xDC, 0x35, 0x45)
        ORANGE = RGBColor(0xFF, 0x8C, 0x00)
        YELLOW_DARK = RGBColor(0xCC, 0x99, 0x00)

        severity_colors = {
            "critical": RED, "high": ORANGE,
            "medium": YELLOW_DARK, "low": GREEN
        }

        def set_cell_shading(cell, color_hex):
            """Set background color for a table cell."""
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
            cell._tc.get_or_add_tcPr().append(shading)

        def set_cell_text(cell, text, bold=False, color=None, size=None, align=None, font=None):
            """Set formatted text in a cell."""
            cell.text = ""
            p = cell.paragraphs[0]
            if align:
                p.alignment = align
            run = p.add_run(str(text))
            run.font.name = font or BODY_FONT
            run.font.size = Pt(size or BODY_SIZE)
            run.bold = bold
            if color:
                run.font.color.rgb = color

        def add_header_row(table, texts, row_idx=0):
            """Style a header row with dark background and white text."""
            for i, txt in enumerate(texts):
                cell = table.rows[row_idx].cells[i]
                set_cell_shading(cell, HEADER_BG)
                set_cell_text(cell, txt, bold=True, color=HEADER_FG, size=BODY_SIZE)

        def add_zebra_row(table, row_idx):
            """Add light gray background to even rows."""
            if row_idx % 2 == 0:
                for cell in table.rows[row_idx].cells:
                    set_cell_shading(cell, LIGHT_GRAY_BG)

        def auto_width(table):
            """Set table width to 100%."""
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl = table._tbl
            tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
            tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:type="pct" w:w="5000"/>')
            tblPr.append(tblW)

        summary = self._summary()
        bugs = self._bugs()
        tests = self._tests()
        perf = self._perf()
        a11y = self._a11y()
        project = self.data.get("project_name", "N/A")
        target = self.data.get("environment", self.data.get("target_url", "N/A"))
        test_date = self.data.get("test_date", datetime.now().strftime("%Y-%m-%d"))
        tester = self.data.get("tester", "QA Ultra Tester")

        total = summary.get("total_tests", summary.get("total", 0))
        if not total and tests:
            total = len(tests)
            passed = sum(1 for t in tests if str(t.get("status", "")).strip().lower() in ["pass", "passed", "success", "ok"])
            failed = sum(1 for t in tests if str(t.get("status", "")).strip().lower() in ["fail", "failed", "error", "defect"])
            flaky = sum(1 for t in tests if t.get("flaky"))
        else:
            passed = summary.get("passed", 0)
            failed = summary.get("failed", 0)
            flaky = summary.get("flaky", 0)
        duration = summary.get("duration_ms", 0)
        quality = summary.get("quality_score", self.data.get("quality_score", 0))
        success_rate = f"{passed*100//total}%" if total else "0%"

        # Determine production readiness
        critical_bugs = sum(1 for b in bugs if b.get("severity", "").lower() == "critical")
        high_bugs = sum(1 for b in bugs if b.get("severity", "").lower() == "high")
        if critical_bugs > 0:
            readiness = "NOT READY - Critical Issues Found"
            readiness_color = RED
        elif high_bugs > 0:
            readiness = "CONDITIONAL - High Priority Fixes Required"
            readiness_color = ORANGE
        elif failed > 0:
            readiness = "PASSED WITH NOTES"
            readiness_color = YELLOW_DARK
        else:
            readiness = "PASSED / PRODUCTION READY"
            readiness_color = GREEN

        # ════════════════════════════════════════════════════════
        # PAGE 1: COVER / HEADER
        # ════════════════════════════════════════════════════════
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_p.space_before = Pt(60)
        title_p.space_after = Pt(8)
        run = title_p.add_run("LAPORAN HASIL PENGUJIAN QA E2E")
        run.font.size = Pt(22)
        run.font.color.rgb = DARK_BLUE
        run.bold = True
        run.font.name = TITLE_FONT

        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_p.space_after = Pt(24)
        run = sub_p.add_run(project.upper())
        run.font.size = Pt(16)
        run.font.color.rgb = DARK_BLUE
        run.bold = True
        run.font.name = TITLE_FONT

        # Cover info table (with background)
        COVER_LABEL_BG = "1B3A5C"
        COVER_VALUE_BG = "E8EDF2"
        cover_tbl = doc.add_table(rows=5, cols=2)
        cover_tbl.style = 'Table Grid'
        cover_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        auto_width(cover_tbl)
        cover_data = [
            ("Target URL", target),
            ("Fitur Uji", project),
            ("Tanggal Uji", test_date),
            ("Metodologi", tester),
            ("Status Kesiapan", readiness),
        ]
        for i, (label, value) in enumerate(cover_data):
            set_cell_shading(cover_tbl.rows[i].cells[0], COVER_LABEL_BG)
            set_cell_text(cover_tbl.rows[i].cells[0], label, bold=True, color=HEADER_FG, size=BODY_SIZE)
            set_cell_shading(cover_tbl.rows[i].cells[1], COVER_VALUE_BG)
            if label == "Status Kesiapan":
                set_cell_text(cover_tbl.rows[i].cells[1], value, bold=True, color=readiness_color, size=BODY_SIZE)
            else:
                set_cell_text(cover_tbl.rows[i].cells[1], value, size=BODY_SIZE)

        doc.add_page_break()

        # ════════════════════════════════════════════════════════
        # SECTION 1: RINGKASAN EKSEKUTIF
        # ════════════════════════════════════════════════════════
        h = doc.add_heading('1. Ringkasan Eksekutif (Executive Summary)', 1)
        for run in h.runs:
            run.font.color.rgb = DARK_BLUE
            run.font.name = TITLE_FONT
            run.font.size = Pt(TITLE_SIZE)

        # Narrative paragraph
        bug_summary_text = ""
        if bugs:
            sev_counts = {}
            for b in bugs:
                s = b.get("severity", "medium").capitalize()
                sev_counts[s] = sev_counts.get(s, 0) + 1
            parts = [f"{v} {k}" for k, v in sev_counts.items()]
            bug_summary_text = f" Ditemukan {len(bugs)} defect ({', '.join(parts)})."
        else:
            bug_summary_text = " Tidak ditemukan defect."

        narrative = (
            f"Pengujian end-to-end (E2E) telah selesai dieksekusi pada menu {project}. "
            f"Dari total {total} skenario pengujian, {passed} skenario berhasil PASSED "
            f"dan {failed} skenario FAILED dengan tingkat keberhasilan {success_rate}."
            f"{bug_summary_text}"
        )
        p = doc.add_paragraph(narrative)
        p.paragraph_format.space_after = Pt(12)

        # Summary table
        sum_tbl = doc.add_table(rows=2, cols=4)
        sum_tbl.style = 'Table Grid'
        auto_width(sum_tbl)
        add_header_row(sum_tbl, ["Total Skenario", "Passed", "Defect / Failed", "Success Rate"])
        set_cell_text(sum_tbl.rows[1].cells[0], str(total), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(sum_tbl.rows[1].cells[1], f"{passed} Skenario", color=GREEN, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(sum_tbl.rows[1].cells[2], f"{failed} Defect", color=RED if failed > 0 else GREEN, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(sum_tbl.rows[1].cells[3], success_rate, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 2: STATISTIK TEMUAN BUG
        # ════════════════════════════════════════════════════════
        if bugs:
            h = doc.add_heading('2. Statistik Temuan Bug', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            # Severity count table
            sev_order = ["critical", "high", "medium", "low"]
            sev_labels = {"critical": "Critical", "high": "High", "medium": "Medium", "low": "Low"}
            sev_tbl = doc.add_table(rows=len(sev_order) + 2, cols=2)
            sev_tbl.style = 'Table Grid'
            auto_width(sev_tbl)
            add_header_row(sev_tbl, ["Severity", "Jumlah"])
            for idx, sev in enumerate(sev_order, 1):
                count = sum(1 for b in bugs if b.get("severity", "").lower() == sev)
                set_cell_text(sev_tbl.rows[idx].cells[0], sev_labels[sev], color=severity_colors.get(sev))
                set_cell_text(sev_tbl.rows[idx].cells[1], str(count), align=WD_ALIGN_PARAGRAPH.CENTER)
                add_zebra_row(sev_tbl, idx)
            # Total row
            total_row = len(sev_order) + 1
            set_cell_text(sev_tbl.rows[total_row].cells[0], "Total", bold=True)
            set_cell_text(sev_tbl.rows[total_row].cells[1], str(len(bugs)), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(sev_tbl.rows[total_row].cells[0], LIGHT_GRAY_BG)
            set_cell_shading(sev_tbl.rows[total_row].cells[1], LIGHT_GRAY_BG)

            doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 3: RINCIAN TEMUAN BUG
        # ════════════════════════════════════════════════════════
        if bugs:
            h = doc.add_heading('3. Rincian Temuan Bug', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            # Bug overview table
            bug_overview = doc.add_table(rows=len(bugs) + 1, cols=5)
            bug_overview.style = 'Table Grid'
            auto_width(bug_overview)
            add_header_row(bug_overview, ["ID", "Judul", "Severity", "Priority", "Status"])

            for idx, bug in enumerate(bugs, 1):
                sev = bug.get("severity", "medium").lower()
                set_cell_text(bug_overview.rows[idx].cells[0], bug.get("id", ""), bold=True)
                set_cell_text(bug_overview.rows[idx].cells[1], bug.get("title", ""))
                set_cell_text(bug_overview.rows[idx].cells[2], sev.capitalize(), bold=True, color=severity_colors.get(sev))
                set_cell_text(bug_overview.rows[idx].cells[3], bug.get("priority", ""))
                set_cell_text(bug_overview.rows[idx].cells[4], bug.get("status", "Open"))
                add_zebra_row(bug_overview, idx)

            doc.add_page_break()

            # Detailed bug reports
            for bug in bugs:
                sev = bug.get("severity", "medium").lower()
                sev_color = severity_colors.get(sev, DARK_BLUE)

                # Bug heading
                bug_h = doc.add_heading(f'{bug.get("id","")}: {bug.get("title","")}', 2)
                for run in bug_h.runs:
                    run.font.color.rgb = sev_color
                    run.font.name = TITLE_FONT
                    run.font.size = Pt(13)

                # Info table (borderless)
                info_data = [
                    ("ID", bug.get("id", "")),
                    ("Severity", bug.get("severity", "").capitalize()),
                    ("Priority", bug.get("priority", "")),
                    ("Menu / Lokasi", bug.get("location", "")),
                    ("Tipe", bug.get("type", "")),
                ]
                info_tbl = doc.add_table(rows=len(info_data), cols=2)
                for i, (label, val) in enumerate(info_data):
                    set_cell_text(info_tbl.rows[i].cells[0], label, bold=True, size=10)
                    if label == "Severity":
                        set_cell_text(info_tbl.rows[i].cells[1], val, bold=True, color=sev_color, size=10)
                    else:
                        set_cell_text(info_tbl.rows[i].cells[1], val, size=10)
                # Remove borders
                for row in info_tbl.rows:
                    for cell in row.cells:
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcBorders = parse_xml(
                            f'<w:tcBorders {nsdecls("w")}>'
                            '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                            '</w:tcBorders>'
                        )
                        tcPr.append(tcBorders)

                # Steps
                steps = bug.get("steps", [])
                if steps:
                    p = doc.add_paragraph()
                    run = p.add_run("Langkah Reproduksi:")
                    run.bold = True
                    run.font.color.rgb = DARK_BLUE
                    for idx_s, step in enumerate(steps, 1):
                        clean = strip_step_number(step)
                        doc.add_paragraph(f'{idx_s}. {clean}')

                # Actual
                p = doc.add_paragraph()
                run = p.add_run("Hasil Aktual: ")
                run.bold = True
                run.font.color.rgb = RED
                p.add_run(bug.get("actual", ""))

                # Expected
                p = doc.add_paragraph()
                run = p.add_run("Hasil yang Diharapkan: ")
                run.bold = True
                run.font.color.rgb = GREEN
                p.add_run(bug.get("expected", ""))

                # Recommendation
                rec = bug.get("recommendation", "")
                if rec:
                    p = doc.add_paragraph()
                    run = p.add_run("Rekomendasi Perbaikan: ")
                    run.bold = True
                    run.font.color.rgb = DARK_BLUE
                    p.add_run(rec)

                # Evidence — screenshot embedded
                evidence = bug.get("evidence", "")
                screenshot = evidence if isinstance(evidence, str) else ""
                if not screenshot and isinstance(evidence, dict):
                    screenshot = evidence.get("screenshot", "")
                if screenshot:
                    spath = Path(screenshot)
                    if not spath.is_absolute():
                        spath = self.output_dir / spath
                    if spath.exists():
                        doc.add_paragraph('')
                        p = doc.add_paragraph()
                        run = p.add_run("Bukti Visual:")
                        run.bold = True
                        run.font.color.rgb = DARK_BLUE
                        doc.add_picture(str(spath), width=Inches(5.5))
                        last_p = doc.paragraphs[-1]
                        last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cap = doc.add_paragraph()
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = cap.add_run(f'Gambar: Bukti {bug.get("id","")} - {bug.get("title","")[:60]}')
                        run.font.size = Pt(9)
                        run.italic = True
                        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

                doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 4: HASIL PENGUJIAN FITUR & VALIDASI
        # ════════════════════════════════════════════════════════
        next_section = 4 if bugs else 2
        doc.add_page_break()
        h = doc.add_heading(f'{next_section}. Hasil Pengujian Fitur & Validasi', 1)
        for run in h.runs:
            run.font.color.rgb = DARK_BLUE

        if tests:
            test_tbl = doc.add_table(rows=len(tests) + 1, cols=5)
            test_tbl.style = 'Table Grid'
            auto_width(test_tbl)
            add_header_row(test_tbl, ["No", "Fitur / Skenario", "Hasil Pengujian", "Status", "Keterangan"])

            for idx, test in enumerate(tests, 1):
                status = test.get("status", "").strip().lower()
                status_text = "PASSED" if status in ["pass", "passed"] else "FAILED"
                status_color = GREEN if status in ["pass", "passed"] else RED
                error = test.get("error", "")
                duration_text = format_duration(test.get("duration_ms", test.get("duration", 0)))

                set_cell_text(test_tbl.rows[idx].cells[0], str(idx), align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(test_tbl.rows[idx].cells[1], test.get("name", ""))
                set_cell_text(test_tbl.rows[idx].cells[2], f"Durasi: {duration_text}" + (f"\n{error}" if error else ""))
                set_cell_text(test_tbl.rows[idx].cells[3], status_text, bold=True, color=status_color, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(test_tbl.rows[idx].cells[4], test.get("layer", "").upper())
                add_zebra_row(test_tbl, idx)

        doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 5: PERFORMANCE METRICS
        # ════════════════════════════════════════════════════════
        if perf:
            next_section += 1
            h = doc.add_heading(f'{next_section}. Metrik Performa (Core Web Vitals)', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            perf_items = []
            for key, val in perf.items():
                metric_key = key.replace("_ms", "").replace("_", " ").upper()
                # Map common keys
                key_map = {
                    "PAGE LOAD MS": "Page Load", "TTFB MS": "TTFB", "FCP MS": "FCP",
                    "LCP MS": "LCP", "CLS": "CLS", "TOTAL REQUESTS": "Total Requests",
                    "TOTAL TRANSFER KB": "Total Transfer",
                    "PAGE_LOAD_MS": "Page Load", "TTFB_MS": "TTFB", "FCP_MS": "FCP",
                    "LCP_MS": "LCP",
                }
                display_key = key_map.get(metric_key, key)
                info = METRIC_INFO.get(display_key, {})
                full_name = info.get("full", display_key)

                if isinstance(val, dict):
                    val_text = f'{val.get("value","")}{val.get("unit","")}'
                    rating = val.get("rating", "")
                elif "ms" in key.lower() or key.lower() in ["ttfb_ms", "fcp_ms", "lcp_ms", "page_load_ms"]:
                    val_text = format_duration(val) if isinstance(val, (int, float)) else str(val)
                    rating = ""
                elif key.lower() == "cls":
                    val_text = str(val)
                    rating = "Good" if val < 0.1 else "Needs Improvement" if val < 0.25 else "Poor"
                else:
                    val_text = str(val)
                    rating = ""
                perf_items.append((display_key, full_name, val_text, rating))

            perf_tbl = doc.add_table(rows=len(perf_items) + 1, cols=4)
            perf_tbl.style = 'Table Grid'
            auto_width(perf_tbl)
            add_header_row(perf_tbl, ["Metrik", "Nama Lengkap", "Nilai", "Rating"])
            for idx, (metric, full, val_t, rating) in enumerate(perf_items, 1):
                set_cell_text(perf_tbl.rows[idx].cells[0], metric, bold=True)
                set_cell_text(perf_tbl.rows[idx].cells[1], full)
                set_cell_text(perf_tbl.rows[idx].cells[2], val_t, align=WD_ALIGN_PARAGRAPH.CENTER)
                rating_color = GREEN if rating == "Good" else (ORANGE if "Improvement" in rating else (RED if rating == "Poor" else None))
                set_cell_text(perf_tbl.rows[idx].cells[3], rating, color=rating_color, align=WD_ALIGN_PARAGRAPH.CENTER)
                add_zebra_row(perf_tbl, idx)

            doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 6: ACCESSIBILITY
        # ════════════════════════════════════════════════════════
        if a11y:
            next_section += 1
            h = doc.add_heading(f'{next_section}. Aksesibilitas (WCAG 2.1 AA)', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            a11y_tbl = doc.add_table(rows=len(a11y) + 1, cols=4)
            a11y_tbl.style = 'Table Grid'
            auto_width(a11y_tbl)
            add_header_row(a11y_tbl, ["Rule", "Impact", "Count", "Deskripsi"])
            for idx, issue in enumerate(a11y, 1):
                impact = issue.get("impact", "").lower()
                impact_color = RED if impact == "critical" else (ORANGE if impact == "serious" else None)
                set_cell_text(a11y_tbl.rows[idx].cells[0], issue.get("rule", ""))
                set_cell_text(a11y_tbl.rows[idx].cells[1], impact.capitalize(), color=impact_color, bold=True)
                set_cell_text(a11y_tbl.rows[idx].cells[2], str(issue.get("count", "")), align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_text(a11y_tbl.rows[idx].cells[3], issue.get("description", ""))
                add_zebra_row(a11y_tbl, idx)

            doc.add_paragraph('')

        # ════════════════════════════════════════════════════════
        # SECTION 7: BUKTI VISUAL (All Screenshots)
        # ════════════════════════════════════════════════════════
        ss_dir = self.output_dir / "screenshots"
        screenshots = []
        if ss_dir.exists():
            for ext in ["*.png", "*.jpg", "*.jpeg"]:
                screenshots.extend(sorted(ss_dir.glob(ext)))

        if screenshots:
            next_section += 1
            doc.add_page_break()
            h = doc.add_heading(f'{next_section}. Bukti Visual Pengujian (Visual Evidence)', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            for img_idx, img_path in enumerate(screenshots, 1):
                try:
                    doc.add_picture(str(img_path), width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap = doc.add_paragraph()
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    img_name = img_path.stem.replace("-", " ").replace("_", " ").title()
                    run = cap.add_run(f'Gambar {img_idx}: {img_name}')
                    run.font.size = Pt(9)
                    run.italic = True
                    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                    doc.add_paragraph('')
                except Exception:
                    pass

        # ════════════════════════════════════════════════════════
        # SECTION 8: REKOMENDASI TINDAK LANJUT
        # ════════════════════════════════════════════════════════
        if bugs:
            next_section += 1
            h = doc.add_heading(f'{next_section}. Rekomendasi Tindak Lanjut', 1)
            for run in h.runs:
                run.font.color.rgb = DARK_BLUE
                run.font.name = TITLE_FONT
                run.font.size = Pt(TITLE_SIZE)

            for idx, bug in enumerate(bugs, 1):
                sev = bug.get("severity", "medium").lower()
                priority = bug.get("priority", "")
                rec = bug.get("recommendation", "")
                p = doc.add_paragraph()
                run = p.add_run(f'{idx}. Prioritas {priority} ({bug.get("id","")})  ')
                run.bold = True
                run.font.color.rgb = severity_colors.get(sev, DARK_BLUE)
                p.add_run(rec)

        # Save
        filepath = self.output_dir / filename
        try:
            doc.save(str(filepath))
            print(f"DOCX report saved: {filepath}")
            return filepath
        except PermissionError:
            alt_path = self.output_dir / f"QA_TEST_REPORT_{int(datetime.now().timestamp())}.docx"
            doc.save(str(alt_path))
            print(f"Warning: File {filepath} is locked (currently opened in Word). Saved as: {alt_path}")
            return alt_path

    # ── PREMIUM HTML ─────────────────────────────────────────
    def generate_html(self, filename="QA_TEST_REPORT.html"):
        score = self._quality_score()
        summary = self._summary()
        bugs = self._bugs()
        tests = self._tests()
        perf = self._perf()
        a11y = self._a11y()

        total = summary.get("total_tests", summary.get("total", 0))
        if not total and tests:
            total = len(tests)
            passed = sum(1 for t in tests if str(t.get("status", "")).strip().lower() in ["pass", "passed", "success", "ok"])
            failed = sum(1 for t in tests if str(t.get("status", "")).strip().lower() in ["fail", "failed", "error", "defect"])
            flaky_count = sum(1 for t in tests if t.get("flaky"))
        else:
            passed = summary.get("passed", 0)
            failed = summary.get("failed", 0)
            flaky_count = summary.get("flaky", 0)
        duration = summary.get("duration_ms", 0)
        duration_display = format_duration(duration)

        score_class = "good" if score >= 80 else "warn" if score >= 60 else "bad"
        readiness = "READY" if score >= 80 and not any(b.get("severity") == "CRITICAL" for b in bugs) else "NOT READY"
        readiness_class = "ready" if readiness == "READY" else "not-ready"

        # Video section
        video_html = ""
        video_path = self._find_video()
        if video_path:
            b64_video = self._embed_file(video_path, "video")
            p = Path(video_path)
            ext = p.suffix.lower().lstrip(".")
            mime = {"mp4": "video/mp4", "webm": "video/webm"}.get(ext, "video/webm")
            if b64_video:
                video_html = f'''
            <section class="report-section">
                <h2>Test Recording</h2>
                <p class="section-desc">Video rekaman proses testing otomatis</p>
                <div class="video-container">
                    <video controls preload="metadata" style="width:100%;max-height:500px;border-radius:8px;">
                        <source src="{b64_video}" type="{mime}">
                        Browser tidak support video tag.
                    </video>
                </div>
            </section>'''
            else:
                # fallback: link to file
                video_html = f'''
            <section class="report-section">
                <h2>Test Recording</h2>
                <p class="section-desc">Video rekaman proses testing otomatis</p>
                <div class="video-container">
                    <video controls preload="metadata" style="width:100%;max-height:500px;border-radius:8px;">
                        <source src="{video_path}" type="{mime}">
                        Browser tidak support video tag.
                    </video>
                </div>
            </section>'''

        # Bug rows
        bug_html = ""
        for bug in bugs:
            sev = bug.get("severity", "MEDIUM")
            steps = bug.get("steps", [])
            # Strip leading numbers to prevent double numbering with <ol>
            steps_html = "".join(f"<li>{strip_step_number(s)}</li>" for s in steps) if steps else "<li>N/A</li>"
            # Evidence screenshot extraction - handle string, dict, or direct field
            evidence = bug.get("evidence", "")
            screenshot = None
            video = None
            if isinstance(evidence, str) and evidence:
                screenshot = evidence
            elif isinstance(evidence, dict):
                screenshot = evidence.get("screenshot") or evidence.get("image")
                video = evidence.get("video")
            if not screenshot:
                screenshot = bug.get("screenshot") or bug.get("image")

            media_html = ""
            # Screenshot
            if screenshot:
                b64_img = self._embed_file(screenshot)
                img_src = b64_img if b64_img else screenshot
                media_html += f'<div class="evidence"><img src="{img_src}" alt="Bukti Bug" class="lightbox-trigger" data-caption="Bukti {bug.get("id","")}: {bug.get("title","")}"></div>'

            # Video per-bug (if exists)
            if video:
                b64_vid = self._embed_file(video, "video")
                if b64_vid:
                    media_html += f'''<div class="evidence">
                        <video controls preload="metadata" style="width:100%;max-height:300px;border-radius:8px;margin-top:8px;">
                            <source src="{b64_vid}" type="video/webm">
                        </video>
                    </div>'''

            bug_html += f'''
            <div class="bug-card {sev.lower()}">
                <div class="bug-header">
                    <span class="severity-badge {sev.lower()}">{sev}</span>
                    <span class="bug-id">{bug.get("id","")}</span>
                    <span class="bug-title">{bug.get("title","")}</span>
                </div>
                <div class="bug-body">
                    <div class="bug-field"><strong>Location:</strong> {bug.get("location","")}</div>
                    <div class="bug-field"><strong>Steps to Reproduce:</strong><ol>{steps_html}</ol></div>
                    <div class="bug-field"><strong>Actual:</strong> {bug.get("actual","")}</div>
                    <div class="bug-field"><strong>Expected:</strong> {bug.get("expected","")}</div>
                    <div class="bug-field"><strong>Recommendation:</strong> {bug.get("recommendation","")}</div>
                    {media_html}
                </div>
            </div>'''

        # Test results rows — with screenshot for failed tests
        test_rows = ""
        for t in tests:
            status_raw = t.get("status", "").strip().lower()
            is_pass = status_raw in ["pass", "passed"]
            status_cls = "pass" if is_pass else "fail"
            flaky_mark = '<span class="flaky-badge">FLAKY</span>' if t.get("flaky") else ""
            dur_ms = t.get("duration_ms", t.get("duration", 0))

            # Screenshot column for failed tests
            screenshot_cell = ""
            if not is_pass:
                evidence = t.get("evidence", {})
                test_ss = t.get("screenshot") or (evidence.get("screenshot") if isinstance(evidence, dict) else None)
                if test_ss:
                    b64 = self._embed_file(test_ss)
                    ss_src = b64 if b64 else test_ss
                    test_name_clean = t.get("name","").replace('"', '&quot;')
                    screenshot_cell = f'<button class="ss-link ss-btn" data-src="{ss_src}" data-caption="Bukti Test: {test_name_clean}">SS</button>'
                else:
                    screenshot_cell = '<span class="no-ss">-</span>'
            else:
                screenshot_cell = '<span class="no-ss">-</span>'

            # Error details for failed tests
            details = t.get("details", "") or t.get("error", "")
            details_html = f'<div class="test-error">{details}</div>' if details and not is_pass else ""
            status_display = "PASSED" if is_pass else "FAILED"

            test_rows += f'''
            <tr class="{status_cls}-row">
                <td>{t.get("name","")}{details_html}</td>
                <td><span class="layer-badge">{t.get("layer","")}</span></td>
                <td><span class="status-badge {status_cls}">{status_display}</span>{flaky_mark}</td>
                <td>{format_duration(dur_ms)}</td>
                <td>{t.get("browser","Chromium")}</td>
                <td>{screenshot_cell}</td>
            </tr>'''

        # Performance rows with full names + descriptions
        perf_html = ""
        if perf:
            perf_rows = ""
            for metric, val in perf.items():
                info = METRIC_INFO.get(metric, {})
                full_name = info.get("full", metric)
                desc = info.get("desc", "")
                if isinstance(val, dict):
                    v = val.get("value", "")
                    u = val.get("unit", "")
                    r = val.get("rating", "unknown")
                else:
                    v = val
                    u = ""
                    r = "unknown"
                bar_pct = min(100, max(5, 100 - int(float(v) / 50) if r != "unknown" else 50))
                perf_rows += f'''
                <tr>
                    <td>
                        <strong>{metric}</strong>
                        <div class="metric-fullname">{full_name}</div>
                        <div class="metric-desc">{desc}</div>
                    </td>
                    <td>{v}{u}</td>
                    <td><div class="perf-bar"><div class="perf-fill {r}" style="width:{bar_pct}%"></div></div></td>
                    <td><span class="rating-badge {r}">{r.replace("_"," ").title()}</span></td>
                </tr>'''
            perf_html = f'''
            <section class="report-section">
                <h2>Performance (Core Web Vitals)</h2>
                <p class="section-desc">Metrik performa sesuai standar Google Core Web Vitals</p>
                <table class="data-table">
                    <thead><tr><th>Metric</th><th>Value</th><th>Bar</th><th>Rating</th></tr></thead>
                    <tbody>{perf_rows}</tbody>
                </table>
            </section>'''

        # Accessibility rows
        a11y_html = ""
        if a11y:
            a11y_rows = ""
            for issue in a11y:
                sev = issue.get("severity", "moderate")
                a11y_rows += f'''
                <tr>
                    <td><span class="severity-badge {sev}">{sev.upper()}</span></td>
                    <td>{issue.get("description","")}</td>
                    <td><code>{issue.get("element","")}</code></td>
                    <td>{issue.get("wcag","")}</td>
                </tr>'''
            a11y_html = f'''
            <section class="report-section">
                <h2>Accessibility (WCAG 2.1 AA)</h2>
                <table class="data-table">
                    <thead><tr><th>Severity</th><th>Description</th><th>Element</th><th>WCAG</th></tr></thead>
                    <tbody>{a11y_rows}</tbody>
                </table>
            </section>'''

        # Visual Evidence Gallery (All screenshots)
        gallery_html = ""
        ss_dir = self.output_dir / "screenshots"
        screenshots = []
        if ss_dir.exists():
            for ext in ["*.png", "*.jpg", "*.jpeg", "*.webp"]:
                screenshots.extend(sorted(ss_dir.glob(ext)))

        if screenshots:
            gallery_items = ""
            for img_idx, img_path in enumerate(screenshots, 1):
                b64_img = self._embed_file(str(img_path))
                img_name = img_path.stem.replace("-", " ").replace("_", " ").title()
                src = b64_img if b64_img else str(img_path)
                gallery_items += f'''
                <div class="gallery-card">
                    <div class="gallery-img-wrap">
                        <img src="{src}" alt="{img_name}" loading="lazy" class="lightbox-trigger" data-caption="Gambar {img_idx}: {img_name}">
                    </div>
                    <div class="gallery-caption">
                        <strong>Gambar {img_idx}:</strong> {img_name}
                    </div>
                </div>'''
            
            gallery_html = f'''
            <section class="report-section">
                <h2>Bukti Visual Pengujian (Visual Evidence)</h2>
                <p class="section-desc">Tangkapan layar bukti pengujian, error, bug, dan ketidaksesuaian UI/UX</p>
                <div class="gallery-grid">
                    {gallery_items}
                </div>
            </section>'''

        # Severity breakdown
        crit = sum(1 for b in bugs if b.get("severity") == "CRITICAL")
        high = sum(1 for b in bugs if b.get("severity") == "HIGH")
        med = sum(1 for b in bugs if b.get("severity") == "MEDIUM")
        low = sum(1 for b in bugs if b.get("severity") == "LOW")

        # Build HTML — note: we use string concat instead of f-string for the JS
        # to avoid {{ }} escaping issues with the toggle function
        html_parts = []
        html_parts.append('''<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>QA Test Report</title>
    <style>
        :root {
            --bg: #f0f2f5; --surface: #ffffff; --text: #1a1a2e; --text-secondary: #6b7280;
            --border: #e5e7eb; --accent: #3b82f6; --success: #10b981; --danger: #ef4444;
            --warning: #f59e0b; --orange: #f97316; --info: #6366f1;
        }
        [data-theme="dark"] {
            --bg: #0f172a; --surface: #1e293b; --text: #e2e8f0; --text-secondary: #94a3b8;
            --border: #334155; --accent: #60a5fa;
        }
        * { margin:0; padding:0; box-sizing:border-box; }
        body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background:var(--bg); color:var(--text); line-height:1.6; }
        .container { max-width:1280px; margin:0 auto; padding:24px; }

        /* Header */
        .report-header { background: linear-gradient(135deg, #1e293b 0%, #334155 100%); color:#fff; padding:40px; border-radius:16px; margin-bottom:24px; position:relative; overflow:hidden; }
        .report-header::after { content:''; position:absolute; top:-50%; right:-20%; width:400px; height:400px; background:radial-gradient(circle, rgba(59,130,246,0.15) 0%, transparent 70%); }
        .report-header h1 { font-size:2rem; font-weight:700; margin-bottom:4px; }
        .report-header .meta { opacity:0.8; font-size:0.9rem; }
        .report-header .meta span { margin-right:20px; }
        .theme-toggle { position:absolute; top:20px; right:20px; background:rgba(255,255,255,0.15); border:none; color:#fff; padding:8px 14px; border-radius:8px; cursor:pointer; font-size:0.85rem; backdrop-filter:blur(4px); transition: background 0.2s; }
        .theme-toggle:hover { background:rgba(255,255,255,0.25); }
        .readiness { display:inline-block; margin-top:12px; padding:6px 18px; border-radius:20px; font-weight:700; font-size:0.95rem; }
        .readiness.ready { background:#10b981; color:#fff; }
        .readiness.not-ready { background:#ef4444; color:#fff; }

        /* Summary Cards */
        .summary-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(180px,1fr)); gap:16px; margin-bottom:24px; }
        .summary-card { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:20px; text-align:center; transition:transform 0.2s, box-shadow 0.2s; }
        .summary-card:hover { transform:translateY(-2px); box-shadow:0 8px 25px rgba(0,0,0,0.08); }
        .summary-card .value { font-size:2.2rem; font-weight:800; line-height:1; }
        .summary-card .label { font-size:0.85rem; color:var(--text-secondary); margin-top:4px; }
        .summary-card .value.good { color:var(--success); }
        .summary-card .value.warn { color:var(--warning); }
        .summary-card .value.bad { color:var(--danger); }
        .summary-card .value.pass { color:var(--success); }
        .summary-card .value.fail { color:var(--danger); }
        .summary-card .value.flaky { color:var(--warning); }

        /* Sections */
        .report-section { background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:24px; margin-bottom:24px; }
        .report-section h2 { font-size:1.3rem; margin-bottom:6px; padding-bottom:8px; border-bottom:2px solid var(--border); }
        .section-desc { font-size:0.85rem; color:var(--text-secondary); margin-bottom:16px; }

        /* Video */
        .video-container { border-radius:8px; overflow:hidden; background:#000; }

        /* Severity Breakdown */
        .severity-breakdown { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:20px; }
        .sev-chip { display:flex; align-items:center; gap:6px; padding:6px 14px; border-radius:8px; font-size:0.85rem; font-weight:600; }
        .sev-chip.critical { background:#fef2f2; color:#dc2626; border:1px solid #fecaca; }
        .sev-chip.high { background:#fff7ed; color:#ea580c; border:1px solid #fed7aa; }
        .sev-chip.medium { background:#fffbeb; color:#d97706; border:1px solid #fde68a; }
        .sev-chip.low { background:#f0fdf4; color:#16a34a; border:1px solid #bbf7d0; }
        [data-theme="dark"] .sev-chip.critical { background:#450a0a; border-color:#991b1b; }
        [data-theme="dark"] .sev-chip.high { background:#431407; border-color:#9a3412; }
        [data-theme="dark"] .sev-chip.medium { background:#451a03; border-color:#92400e; }
        [data-theme="dark"] .sev-chip.low { background:#052e16; border-color:#166534; }

        /* Bug Cards */
        .bug-card { border:1px solid var(--border); border-radius:10px; margin-bottom:16px; overflow:hidden; transition:box-shadow 0.2s; }
        .bug-card:hover { box-shadow:0 4px 15px rgba(0,0,0,0.06); }
        .bug-card.critical { border-left:4px solid #ef4444; }
        .bug-card.high { border-left:4px solid #f97316; }
        .bug-card.medium { border-left:4px solid #f59e0b; }
        .bug-card.low { border-left:4px solid #10b981; }
        .bug-header { padding:14px 18px; background:var(--bg); display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
        .bug-body { padding:18px; }
        .bug-field { margin-bottom:10px; }
        .bug-field ol { margin-left:20px; margin-top:4px; }
        .bug-field ol li { margin-bottom:2px; }
        .evidence { margin-top:12px; }
        .evidence img { max-width:100%; border-radius:8px; border:1px solid var(--border); box-shadow:0 2px 8px rgba(0,0,0,0.06); }

        /* Visual Gallery */
        .gallery-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:20px; margin-top:12px; }
        .gallery-card { background:var(--bg); border:1px solid var(--border); border-radius:10px; overflow:hidden; transition:transform 0.2s, box-shadow 0.2s; }
        .gallery-card:hover { transform:translateY(-2px); box-shadow:0 6px 20px rgba(0,0,0,0.1); }
        .gallery-img-wrap { overflow:hidden; background:#000; display:flex; align-items:center; justify-content:center; max-height:240px; }
        .gallery-img-wrap img { width:100%; height:auto; object-fit:contain; transition:transform 0.3s ease; }
        .gallery-img-wrap:hover img { transform:scale(1.03); }
        .gallery-caption { padding:12px 14px; font-size:0.85rem; color:var(--text); border-top:1px solid var(--border); }

        /* Badges */
        .severity-badge { display:inline-block; padding:3px 10px; border-radius:6px; font-size:0.75rem; font-weight:700; color:#fff; text-transform:uppercase; }
        .severity-badge.critical { background:#ef4444; }
        .severity-badge.high { background:#f97316; }
        .severity-badge.medium { background:#f59e0b; color:#1a1a2e; }
        .severity-badge.low { background:#10b981; }
        .severity-badge.serious { background:#f97316; }
        .severity-badge.moderate { background:#f59e0b; color:#1a1a2e; }
        .severity-badge.minor { background:#10b981; }
        .severity-badge.violation { background:#ef4444; }
        .status-badge { padding:3px 10px; border-radius:6px; font-size:0.8rem; font-weight:600; }
        .status-badge.pass { background:#d1fae5; color:#065f46; }
        .status-badge.fail { background:#fee2e2; color:#991b1b; }
        [data-theme="dark"] .status-badge.pass { background:#064e3b; color:#6ee7b7; }
        [data-theme="dark"] .status-badge.fail { background:#450a0a; color:#fca5a5; }
        .flaky-badge { background:#fef3c7; color:#92400e; padding:2px 8px; border-radius:6px; font-size:0.7rem; font-weight:600; margin-left:6px; }
        .layer-badge { background:var(--bg); padding:2px 8px; border-radius:6px; font-size:0.8rem; font-weight:500; }
        .rating-badge { padding:3px 10px; border-radius:6px; font-size:0.8rem; font-weight:600; }
        .rating-badge.good { background:#d1fae5; color:#065f46; }
        .rating-badge.needs_improvement { background:#fef3c7; color:#92400e; }
        .rating-badge.poor { background:#fee2e2; color:#991b1b; }
        [data-theme="dark"] .rating-badge.good { background:#064e3b; color:#6ee7b7; }
        [data-theme="dark"] .rating-badge.needs_improvement { background:#451a03; color:#fbbf24; }
        [data-theme="dark"] .rating-badge.poor { background:#450a0a; color:#fca5a5; }
        .bug-id { font-weight:700; font-family:monospace; }
        .bug-title { font-weight:600; }

        /* Test error details */
        .test-error { font-size:0.8rem; color:var(--danger); margin-top:4px; font-family:monospace; background:rgba(239,68,68,0.06); padding:4px 8px; border-radius:4px; }
        .ss-link { display:inline-block; padding:2px 8px; background:var(--accent); color:#fff; border-radius:4px; text-decoration:none; font-size:0.75rem; font-weight:600; border:none; cursor:pointer; }
        .ss-link:hover { opacity:0.85; }
        .no-ss { color:var(--text-secondary); }
        .fail-row { background:rgba(239,68,68,0.03); }
        [data-theme="dark"] .fail-row { background:rgba(239,68,68,0.08); }

        /* Lightbox Modal */
        .lightbox-trigger { cursor: zoom-in; }
        .lightbox-modal { display:none; position:fixed; z-index:9999; inset:0; background:rgba(0,0,0,0.88); backdrop-filter:blur(6px); justify-content:center; align-items:center; flex-direction:column; padding:20px; }
        .lightbox-modal.active { display:flex; animation:fadeIn 0.2s ease; }
        .lightbox-content { max-width:92vw; max-height:85vh; border-radius:8px; box-shadow:0 10px 40px rgba(0,0,0,0.5); object-fit:contain; }
        .lightbox-caption { color:#e2e8f0; font-size:0.95rem; margin-top:12px; text-align:center; max-width:800px; }
        .lightbox-close { position:absolute; top:20px; right:24px; color:#fff; font-size:2rem; font-weight:bold; cursor:pointer; background:rgba(255,255,255,0.15); width:44px; height:44px; border-radius:50%; display:flex; align-items:center; justify-content:center; border:none; transition:background 0.2s; }
        .lightbox-close:hover { background:rgba(255,255,255,0.3); }
        @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }

        /* Metric descriptions */
        .metric-fullname { font-size:0.8rem; color:var(--text-secondary); font-weight:500; }
        .metric-desc { font-size:0.75rem; color:var(--text-secondary); font-style:italic; margin-top:2px; }

        /* Tables */
        .data-table { width:100%; border-collapse:collapse; }
        .data-table th { text-align:left; padding:10px 14px; background:var(--bg); font-size:0.85rem; font-weight:600; color:var(--text-secondary); border-bottom:2px solid var(--border); }
        .data-table td { padding:10px 14px; border-bottom:1px solid var(--border); font-size:0.9rem; }
        .data-table tr:hover td { background:rgba(59,130,246,0.04); }

        /* Perf bars */
        .perf-bar { width:100%; height:8px; background:var(--bg); border-radius:4px; overflow:hidden; }
        .perf-fill { height:100%; border-radius:4px; transition:width 0.5s ease; }
        .perf-fill.good { background:var(--success); }
        .perf-fill.needs_improvement { background:var(--warning); }
        .perf-fill.poor { background:var(--danger); }

        /* Footer */
        .report-footer { text-align:center; padding:20px; color:var(--text-secondary); font-size:0.8rem; }

        /* Print */
        @media print {
            body { background:#fff; }
            .theme-toggle { display:none; }
            .report-header { background:#1e293b !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
            .summary-card:hover { transform:none; box-shadow:none; }
            .video-container { display:none; }
        }
        @media (max-width: 768px) {
            .container { padding:12px; }
            .report-header { padding:24px; }
            .report-header h1 { font-size:1.4rem; }
            .summary-grid { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <div class="container">
''')

        # Header section with dynamic data
        html_parts.append(f'''        <header class="report-header">
            <button class="theme-toggle" id="themeBtn">Dark / Light</button>
            <h1>QA Test Report</h1>
            <div class="meta">
                <span>Project: {self.data.get("project_name","N/A")}</span>
                <span>Target: {self.data.get("target_url","N/A")}</span>
                <span>Generated: {self.data.get("generated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</span>
            </div>
            <div class="readiness {readiness_class}">Production: {readiness}</div>
        </header>

        <div class="summary-grid">
            <div class="summary-card">
                <div class="value {score_class}">{score}</div>
                <div class="label">Quality Score</div>
            </div>
            <div class="summary-card">
                <div class="value">{total}</div>
                <div class="label">Total Tests</div>
            </div>
            <div class="summary-card">
                <div class="value pass">{passed}</div>
                <div class="label">Passed</div>
            </div>
            <div class="summary-card">
                <div class="value fail">{failed}</div>
                <div class="label">Failed</div>
            </div>
            <div class="summary-card">
                <div class="value flaky">{flaky_count}</div>
                <div class="label">Flaky</div>
            </div>
            <div class="summary-card">
                <div class="value">{duration_display}</div>
                <div class="label">Duration</div>
            </div>
        </div>

        {video_html}

        <section class="report-section">
            <h2>Bug Summary ({len(bugs)} found)</h2>
            <div class="severity-breakdown">
                <div class="sev-chip critical">CRITICAL: {crit}</div>
                <div class="sev-chip high">HIGH: {high}</div>
                <div class="sev-chip medium">MEDIUM: {med}</div>
                <div class="sev-chip low">LOW: {low}</div>
            </div>
            {bug_html if bug_html else '<p style="color:var(--text-secondary)">No bugs found.</p>'}
        </section>

        <section class="report-section">
            <h2>Test Results</h2>
            <table class="data-table">
                <thead><tr><th>Test Name</th><th>Layer</th><th>Status</th><th>Duration</th><th>Browser</th><th>Evidence</th></tr></thead>
                <tbody>
                    {test_rows if test_rows else '<tr><td colspan="6" style="text-align:center;color:var(--text-secondary)">No test results.</td></tr>'}
                </tbody>
            </table>
        </section>

        {perf_html}
        {a11y_html}
        {gallery_html}

        <footer class="report-footer">
            QA Ultra Tester v2.1 &mdash; Generated from test-results.json (single source of truth)
        </footer>
    </div>

    <!-- Lightbox Modal -->
    <div id="lightboxModal" class="lightbox-modal">
        <button class="lightbox-close" id="lightboxClose">&times;</button>
        <img class="lightbox-content" id="lightboxImg" src="" alt="Enlarged View">
        <div class="lightbox-caption" id="lightboxCaption"></div>
    </div>
''')

        # JavaScript — separate to avoid f-string escaping issues
        html_parts.append('''    <script>
        (function() {
            var btn = document.getElementById('themeBtn');
            var html = document.documentElement;

            // Load saved theme
            var saved = localStorage.getItem('qa-report-theme');
            if (saved) {
                html.setAttribute('data-theme', saved);
            }

            btn.addEventListener('click', function() {
                var current = html.getAttribute('data-theme');
                var next = current === 'dark' ? 'light' : 'dark';
                html.setAttribute('data-theme', next);
                localStorage.setItem('qa-report-theme', next);
            });

            // Lightbox Modal Logic
            var modal = document.getElementById('lightboxModal');
            var modalImg = document.getElementById('lightboxImg');
            var modalCaption = document.getElementById('lightboxCaption');
            var closeBtn = document.getElementById('lightboxClose');

            function openLightbox(src, caption) {
                if (!src) return;
                modalImg.src = src;
                modalCaption.textContent = caption || '';
                modal.classList.add('active');
                document.body.style.overflow = 'hidden';
            }

            function closeLightbox() {
                modal.classList.remove('active');
                modalImg.src = '';
                document.body.style.overflow = '';
            }

            // Bind image clicks with .lightbox-trigger
            document.querySelectorAll('.lightbox-trigger').forEach(function(img) {
                img.addEventListener('click', function() {
                    openLightbox(this.src, this.getAttribute('data-caption') || this.alt);
                });
            });

            // Bind table SS buttons
            document.querySelectorAll('.ss-btn').forEach(function(btn) {
                btn.addEventListener('click', function(e) {
                    e.preventDefault();
                    var src = this.getAttribute('data-src');
                    var cap = this.getAttribute('data-caption');
                    openLightbox(src, cap);
                });
            });

            // Close actions
            closeBtn.addEventListener('click', closeLightbox);
            modal.addEventListener('click', function(e) {
                if (e.target === modal) {
                    closeLightbox();
                }
            });
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && modal.classList.contains('active')) {
                    closeLightbox();
                }
            });
        })();
    </script>
</body>
</html>''')

        html = "".join(html_parts)

        filepath = self.output_dir / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"HTML report saved: {filepath}")
        return filepath

    # ── JSON (write back / create) ───────────────────────────
    def generate_json(self, filename="test-results.json"):
        """Write current data back as JSON (useful for programmatic use)."""
        filepath = self.output_dir / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)
        print(f"JSON report saved: {filepath}")
        return filepath


# ── CLI entry point ──────────────────────────────────────────
if __name__ == "__main__":
    output_dir = sys.argv[1] if len(sys.argv) > 1 else "hasil-test"
    gen = QAReportGenerator(output_dir)

    if gen.load_from_json():
        gen.generate_html()
        gen.generate_docx()
        print(f"\nAll reports generated in {gen.output_dir}")
    else:
        # Demo mode with sample data
        print("No test-results.json found. Generating demo report...")
        demo_data = {
            "project_name": "Demo Project",
            "target_url": "https://example.com",
            "generated_at": datetime.now().isoformat(),
            "quality_score": 72,
            "summary": {"total": 15, "passed": 11, "failed": 3, "flaky": 1, "duration_ms": 185000},
            "bugs": [
                {
                    "id": "BUG-001", "severity": "CRITICAL",
                    "title": "SQL Injection on Login",
                    "location": "https://example.com/login",
                    "steps": ["Open login page", "Input SQL payload in email field", "Click login button"],
                    "actual": "Login successful without valid credentials",
                    "expected": "Login should fail with invalid credentials",
                    "evidence": {"screenshot": "", "video": "", "network_log": ""},
                    "recommendation": "Implement parameterized queries"
                },
                {
                    "id": "BUG-002", "severity": "HIGH",
                    "title": "XSS Reflected on Search",
                    "location": "https://example.com/search",
                    "steps": ["Navigate to search page", "Input <script>alert(1)</script> in search box", "Submit form"],
                    "actual": "Script executed in browser",
                    "expected": "Input should be sanitized",
                    "evidence": {"screenshot": "", "video": ""},
                    "recommendation": "Escape all user input in output"
                }
            ],
            "test_results": [
                {"name": "Login Happy Path", "layer": "UI", "status": "PASS", "duration_ms": 1250, "browser": "Chromium", "retries": 0, "flaky": False, "details": ""},
                {"name": "Dashboard Load", "layer": "UI", "status": "PASS", "duration_ms": 3400, "browser": "Chromium", "retries": 0, "flaky": False, "details": ""},
                {"name": "Form Submit Invalid", "layer": "UI", "status": "FAIL", "duration_ms": 2100, "browser": "Chromium", "retries": 2, "flaky": False, "details": "Timeout waiting for error message", "screenshot": ""},
                {"name": "API GET /users", "layer": "API", "status": "PASS", "duration_ms": 450, "browser": "Chromium", "retries": 0, "flaky": False, "details": ""},
                {"name": "Cart Checkout Flow", "layer": "UI", "status": "FAIL", "duration_ms": 8200, "browser": "Chromium", "retries": 1, "flaky": False, "details": "Expected price $10.00 but got $0.00"},
                {"name": "Mobile Responsive Header", "layer": "UI", "status": "PASS", "duration_ms": 1800, "browser": "Chromium", "retries": 0, "flaky": True, "details": ""}
            ],
            "performance": {
                "LCP": {"value": 2300, "unit": "ms", "rating": "good"},
                "FID": {"value": 45, "unit": "ms", "rating": "good"},
                "CLS": {"value": 0.05, "unit": "", "rating": "good"},
                "TTFB": {"value": 820, "unit": "ms", "rating": "needs_improvement"},
                "FCP": {"value": 1200, "unit": "ms", "rating": "good"}
            },
            "accessibility": [
                {"type": "color_contrast", "element": ".text-muted", "description": "Contrast ratio 2.5:1 below 4.5:1 minimum", "wcag": "1.4.3", "severity": "serious"},
                {"type": "missing_alt", "element": "img.hero-banner", "description": "Image missing alt attribute", "wcag": "1.1.1", "severity": "critical"}
            ]
        }
        gen.set_data(demo_data)
        gen.generate_json()
        gen.generate_html()
        gen.generate_docx()
        print(f"\nDemo reports generated in {gen.output_dir}")
