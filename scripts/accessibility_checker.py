#!/usr/bin/env python3
"""
QA Ultra Tester - Accessibility Checker
Checks WCAG 2.1 AA compliance
"""

import json
from pathlib import Path
from datetime import datetime


class AccessibilityChecker:
    WCAG_CRITERIA = {
        "color_contrast": {"level": "AA", "criterion": "1.4.3"},
        "text_resize": {"level": "AA", "criterion": "1.4.4"},
        "images_alt": {"level": "A", "criterion": "1.1.1"},
        "keyboard_navigation": {"level": "A", "criterion": "2.1.1"},
        "focus_visible": {"level": "A", "criterion": "2.4.7"},
        "skip_navigation": {"level": "A", "criterion": "2.4.1"},
        "form_labels": {"level": "A", "criterion": "1.3.1"},
        "error_identification": {"level": "A", "criterion": "3.3.1"},
        "language_attribute": {"level": "A", "criterion": "3.1.1"},
        "page_title": {"level": "A", "criterion": "2.4.2"},
        "heading_hierarchy": {"level": "A", "criterion": "1.3.1"},
        "link_purpose": {"level": "A", "criterion": "2.4.4"},
        "aria_landmarks": {"level": "A", "criterion": "1.3.1"},
        "table_headers": {"level": "A", "criterion": "1.3.1"}
    }

    def __init__(self, output_dir="hasil-test/accessibility"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.issues = []
        self.passes = []

    def add_issue(self, issue_type, element, description, wcag_criterion, severity="violation"):
        self.issues.append({
            "type": issue_type,
            "element": element,
            "description": description,
            "wcag": wcag_criterion,
            "severity": severity,
            "timestamp": datetime.now().isoformat()
        })

    def add_pass(self, check_type, element, description, wcag_criterion):
        self.passes.append({
            "type": check_type,
            "element": element,
            "description": description,
            "wcag": wcag_criterion,
            "timestamp": datetime.now().isoformat()
        })

    def calculate_compliance_score(self):
        total_checks = len(self.issues) + len(self.passes)
        if total_checks == 0:
            return 100
        return int((len(self.passes) / total_checks) * 100)

    def get_violations_by_severity(self):
        violations = {"critical": [], "serious": [], "moderate": [], "minor": []}
        for issue in self.issues:
            severity = issue.get("severity", "moderate")
            if severity in violations:
                violations[severity].append(issue)
        return violations

    def generate_report(self):
        report = {
            "generated_at": datetime.now().isoformat(),
            "compliance_score": self.calculate_compliance_score(),
            "summary": {
                "total_checks": len(self.issues) + len(self.passes),
                "violations": len(self.issues),
                "passes": len(self.passes),
                "by_severity": {
                    "critical": len([i for i in self.issues if i.get("severity") == "critical"]),
                    "serious": len([i for i in self.issues if i.get("severity") == "serious"]),
                    "moderate": len([i for i in self.issues if i.get("severity") == "moderate"]),
                    "minor": len([i for i in self.issues if i.get("severity") == "minor"])
                }
            },
            "violations": self.issues,
            "passes": self.passes
        }

        filepath = self.output_dir / "accessibility-report.json"
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)

        return report

    def print_summary(self):
        print("\n" + "=" * 50)
        print("ACCESSIBILITY AUDIT SUMMARY")
        print("=" * 50)
        print(f"Compliance Score: {self.calculate_compliance_score()}%")
        print(f"Total Checks: {len(self.issues) + len(self.passes)}")
        print(f"Violations: {len(self.issues)}")
        print(f"Passes: {len(self.passes)}")

        if self.issues:
            print("\nTop Violations:")
            for issue in self.issues[:5]:
                print(f"  ❌ [{issue['severity'].upper()}] {issue['description']}")
                print(f"     Element: {issue['element']}")
                print(f"     WCAG: {issue['wcag']}")
                print()

        print("=" * 50)


if __name__ == "__main__":
    checker = AccessibilityChecker()

    # Example violations
    checker.add_issue(
        "color_contrast",
        ".text-muted",
        "Color contrast ratio 2.5:1 is below minimum 4.5:1",
        "1.4.3",
        "serious"
    )

    checker.add_issue(
        "images_alt",
        '<img src="logo.png">',
        "Image missing alt attribute",
        "1.1.1",
        "critical"
    )

    # Example passes
    checker.add_pass(
        "keyboard_navigation",
        '<button>Submit</button>',
        "Button is keyboard accessible",
        "2.1.1"
    )

    report = checker.generate_report()
    checker.print_summary()

    print(f"\nReport saved to: {checker.output_dir}/accessibility-report.json")
