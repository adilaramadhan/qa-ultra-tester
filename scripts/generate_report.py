#!/usr/bin/env python3
"""
QA Ultra Tester - Report Generator
Generates Word (.docx) and HTML reports with embedded screenshots
"""

import os
import json
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class QAReportGenerator:
    def __init__(self, output_dir="hasil-test"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.bugs = []
        self.test_results = []
        self.performance_metrics = {}
        self.accessibility_issues = []

    def add_bug(self, bug_id, severity, title, location, steps, actual, expected, recommendation, screenshot_path=None):
        self.bugs.append({
            "id": bug_id,
            "severity": severity,
            "title": title,
            "location": location,
            "steps": steps,
            "actual": actual,
            "expected": expected,
            "recommendation": recommendation,
            "screenshot": screenshot_path
        })

    def add_test_result(self, test_name, status, duration, browser="Chromium", retries=0):
        self.test_results.append({
            "name": test_name,
            "status": status,
            "duration": duration,
            "browser": browser,
            "retries": retries,
            "flaky": retries > 0 and status == "PASS"
        })

    def set_performance_metrics(self, metrics):
        self.performance_metrics = metrics

    def add_accessibility_issue(self, issue_type, element, description, wcag_criterion):
        self.accessibility_issues.append({
            "type": issue_type,
            "element": element,
            "description": description,
            "wcag": wcag_criterion
        })

    def calculate_quality_score(self):
        if not self.test_results:
            return 0

        total = len(self.test_results)
        passed = sum(1 for t in self.test_results if t["status"] == "PASS")
        critical_bugs = sum(1 for b in self.bugs if b["severity"] == "CRITICAL")
        high_bugs = sum(1 for b in self.bugs if b["severity"] == "HIGH")

        base_score = (passed / total * 100) if total > 0 else 0
        penalty = (critical_bugs * 15) + (high_bugs * 8) + (len(self.accessibility_issues) * 2)

        return max(0, min(100, int(base_score - penalty)))

    def generate_docx(self, filename="QA_TEST_REPORT.docx"):
        if not HAS_DOCX:
            print("Warning: python-docx not installed. Skipping DOCX generation.")
            return None

        doc = Document()

        # Title
        title = doc.add_heading('QA TEST REPORT', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Metadata
        meta = doc.add_paragraph()
        meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta.add_run(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
        meta.add_run(f'Quality Score: {self.calculate_quality_score()}/100')

        doc.add_page_break()

        # Executive Summary
        doc.add_heading('Executive Summary', 1)
        total = len(self.test_results)
        passed = sum(1 for t in self.test_results if t["status"] == "PASS")
        failed = sum(1 for t in self.test_results if t["status"] == "FAIL")
        flaky = sum(1 for t in self.test_results if t.get("flaky"))

        summary_table = doc.add_table(rows=5, cols=2)
        summary_table.style = 'Table Grid'
        cells = [
            ("Total Tests", str(total)),
            ("Passed", f"{passed} ({passed*100//total if total else 0}%)"),
            ("Failed", f"{failed} ({failed*100//total if total else 0}%)"),
            ("Flaky", f"{flaky} ({flaky*100//total if total else 0}%)"),
            ("Critical Bugs", str(sum(1 for b in self.bugs if b["severity"] == "CRITICAL")))
        ]
        for i, (label, value) in enumerate(cells):
            summary_table.rows[i].cells[0].text = label
            summary_table.rows[i].cells[1].text = value

        doc.add_page_break()

        # Bug Details
        doc.add_heading('Bug Details', 1)
        for bug in self.bugs:
            severity_colors = {
                "CRITICAL": RGBColor(220, 53, 69),
                "HIGH": RGBColor(255, 193, 7),
                "MEDIUM": RGBColor(255, 193, 7),
                "LOW": RGBColor(40, 167, 69)
            }

            heading = doc.add_heading(f'[{bug["severity"]}] {bug["id"]}: {bug["title"]}', 2)
            heading.runs[0].font.color.rgb = severity_colors.get(bug["severity"], RGBColor(0, 0, 0))

            doc.add_paragraph(f'Location: {bug["location"]}')
            doc.add_paragraph(f'Actual: {bug["actual"]}')
            doc.add_paragraph(f'Expected: {bug["expected"]}')
            doc.add_paragraph(f'Recommendation: {bug["recommendation"]}')

            if bug.get("screenshot") and os.path.exists(bug["screenshot"]):
                doc.add_picture(bug["screenshot"], width=Inches(5.5))
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph('')

        # Performance Metrics
        if self.performance_metrics:
            doc.add_page_break()
            doc.add_heading('Performance Metrics', 1)
            perf_table = doc.add_table(rows=len(self.performance_metrics), cols=2)
            perf_table.style = 'Table Grid'
            for i, (metric, value) in enumerate(self.performance_metrics.items()):
                perf_table.rows[i].cells[0].text = metric
                perf_table.rows[i].cells[1].text = str(value)

        # Accessibility Issues
        if self.accessibility_issues:
            doc.add_page_break()
            doc.add_heading('Accessibility Issues (WCAG 2.1 AA)', 1)
            for issue in self.accessibility_issues:
                doc.add_paragraph(f'• [{issue["type"]}] {issue["description"]}')
                doc.add_paragraph(f'  Element: {issue["element"]}')
                doc.add_paragraph(f'  WCAG: {issue["wcag"]}')

        filepath = self.output_dir / filename
        doc.save(str(filepath))
        return filepath

    def generate_json(self, filename="test-results.json"):
        results = {
            "generated_at": datetime.now().isoformat(),
            "quality_score": self.calculate_quality_score(),
            "summary": {
                "total": len(self.test_results),
                "passed": sum(1 for t in self.test_results if t["status"] == "PASS"),
                "failed": sum(1 for t in self.test_results if t["status"] == "FAIL"),
                "flaky": sum(1 for t in self.test_results if t.get("flaky"))
            },
            "bugs": self.bugs,
            "test_results": self.test_results,
            "performance": self.performance_metrics,
            "accessibility": self.accessibility_issues
        }

        filepath = self.output_dir / filename
        with open(filepath, 'w') as f:
            json.dump(results, f, indent=2)
        return filepath

    def generate_html(self, filename="QA_TEST_REPORT.html"):
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>QA Test Report</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 40px; background: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        h1 {{ color: #333; border-bottom: 3px solid #007bff; padding-bottom: 10px; }}
        h2 {{ color: #555; margin-top: 30px; }}
        .summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 20px 0; }}
        .card {{ background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; }}
        .card h3 {{ margin: 0; font-size: 2em; }}
        .card p {{ margin: 5px 0 0; color: #666; }}
        .score {{ font-size: 3em; font-weight: bold; }}
        .score.good {{ color: #28a745; }}
        .score.medium {{ color: #ffc107; }}
        .score.bad {{ color: #dc3545; }}
        .bug {{ border-left: 4px solid; padding: 15px; margin: 15px 0; background: #fff; border-radius: 4px; }}
        .bug.CRITICAL {{ border-color: #dc3545; }}
        .bug.HIGH {{ border-color: #fd7e14; }}
        .bug.MEDIUM {{ border-color: #ffc107; }}
        .bug.LOW {{ border-color: #28a745; }}
        .severity {{ font-weight: bold; padding: 2px 8px; border-radius: 4px; color: white; }}
        .severity.CRITICAL {{ background: #dc3545; }}
        .severity.HIGH {{ background: #fd7e14; }}
        .severity.MEDIUM {{ background: #ffc107; color: #333; }}
        .severity.LOW {{ background: #28a745; }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background: #f8f9fa; }}
        .pass {{ color: #28a745; }}
        .fail {{ color: #dc3545; }}
        .flaky {{ color: #ffc107; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>QA Test Report</h1>
        <p>Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>

        <div class="summary">
            <div class="card">
                <h3 class="score {'good' if self.calculate_quality_score() >= 80 else 'medium' if self.calculate_quality_score() >= 60 else 'bad'}">{self.calculate_quality_score()}</h3>
                <p>Quality Score</p>
            </div>
            <div class="card">
                <h3>{len(self.test_results)}</h3>
                <p>Total Tests</p>
            </div>
            <div class="card">
                <h3 class="pass">{sum(1 for t in self.test_results if t['status'] == 'PASS')}</h3>
                <p>Passed</p>
            </div>
            <div class="card">
                <h3 class="fail">{sum(1 for t in self.test_results if t['status'] == 'FAIL')}</h3>
                <p>Failed</p>
            </div>
        </div>

        <h2>Bug Details ({len(self.bugs)} found)</h2>
        {''.join(f'''
        <div class="bug {bug['severity']}">
            <span class="severity {bug['severity']}">{bug['severity']}</span>
            <strong>{bug['id']}: {bug['title']}</strong>
            <p><strong>Location:</strong> {bug['location']}</p>
            <p><strong>Actual:</strong> {bug['actual']}</p>
            <p><strong>Expected:</strong> {bug['expected']}</p>
            <p><strong>Recommendation:</strong> {bug['recommendation']}</p>
        </div>
        ''' for bug in self.bugs)}

        <h2>Test Results</h2>
        <table>
            <tr><th>Test Name</th><th>Status</th><th>Browser</th><th>Duration</th><th>Flaky</th></tr>
            {''.join(f'''
            <tr>
                <td>{t['name']}</td>
                <td class="{'pass' if t['status'] == 'PASS' else 'fail'}">{t['status']}</td>
                <td>{t['browser']}</td>
                <td>{t['duration']}ms</td>
                <td class="flaky">{'⚠️' if t.get('flaky') else '✅'}</td>
            </tr>
            ''' for t in self.test_results)}
        </table>
    </div>
</body>
</html>"""

        filepath = self.output_dir / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html)
        return filepath


if __name__ == "__main__":
    # Example usage
    generator = QAReportGenerator()

    generator.add_test_result("Login Test", "PASS", 1250, "Chromium")
    generator.add_test_result("Dashboard Load", "PASS", 3400, "Chromium")
    generator.add_test_result("Form Submit", "FAIL", 2100, "Chromium", retries=2)

    generator.add_bug(
        "BUG-001", "CRITICAL", "SQL Injection on Login",
        "https://example.com/login",
        ["1. Open login page", "2. Input SQL injection payload", "3. Click login"],
        "Login successful without valid credentials",
        "Login should fail with invalid credentials",
        "Implement parameterized queries"
    )

    generator.set_performance_metrics({
        "LCP": "2.3s",
        "FID": "45ms",
        "CLS": "0.05",
        "TTFB": "320ms"
    })

    generator.generate_docx()
    generator.generate_json()
    generator.generate_html()

    print(f"Reports generated in {generator.output_dir}")
