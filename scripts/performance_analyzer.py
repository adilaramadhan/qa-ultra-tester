#!/usr/bin/env python3
"""
QA Ultra Tester - Performance Analyzer
Analyzes Core Web Vitals and performance metrics
"""

import json
from pathlib import Path
from datetime import datetime


class PerformanceAnalyzer:
    THRESHOLDS = {
        "LCP": {"good": 2500, "needs_improvement": 4000},
        "FID": {"good": 100, "needs_improvement": 300},
        "CLS": {"good": 0.1, "needs_improvement": 0.25},
        "TTFB": {"good": 800, "needs_improvement": 1800},
        "FCP": {"good": 1800, "needs_improvement": 3000},
        "INP": {"good": 200, "needs_improvement": 500}
    }

    def __init__(self, output_dir="hasil-test/performance"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.metrics = []

    def evaluate_metric(self, name, value):
        if name not in self.THRESHOLDS:
            return "unknown"

        thresholds = self.THRESHOLDS[name]
        if value <= thresholds["good"]:
            return "good"
        elif value <= thresholds["needs_improvement"]:
            return "needs_improvement"
        else:
            return "poor"

    def add_metric(self, name, value, unit="ms"):
        rating = self.evaluate_metric(name, value)
        self.metrics.append({
            "name": name,
            "value": value,
            "unit": unit,
            "rating": rating,
            "timestamp": datetime.now().isoformat()
        })

    def calculate_performance_score(self):
        if not self.metrics:
            return 0

        scores = []
        for metric in self.metrics:
            if metric["rating"] == "good":
                scores.append(100)
            elif metric["rating"] == "needs_improvement":
                scores.append(50)
            else:
                scores.append(0)

        return int(sum(scores) / len(scores)) if scores else 0

    def generate_report(self):
        report = {
            "generated_at": datetime.now().isoformat(),
            "performance_score": self.calculate_performance_score(),
            "metrics": self.metrics,
            "summary": {
                "good": sum(1 for m in self.metrics if m["rating"] == "good"),
                "needs_improvement": sum(1 for m in self.metrics if m["rating"] == "needs_improvement"),
                "poor": sum(1 for m in self.metrics if m["rating"] == "poor")
            }
        }

        filepath = self.output_dir / "performance-report.json"
        with open(filepath, 'w') as f:
            json.dump(report, f, indent=2)

        return report

    def print_summary(self):
        print("\n" + "=" * 50)
        print("PERFORMANCE SUMMARY")
        print("=" * 50)

        for metric in self.metrics:
            icon = "✅" if metric["rating"] == "good" else "⚠️" if metric["rating"] == "needs_improvement" else "❌"
            print(f"{icon} {metric['name']}: {metric['value']}{metric['unit']} ({metric['rating']})")

        print(f"\nPerformance Score: {self.calculate_performance_score()}/100")
        print("=" * 50)


if __name__ == "__main__":
    analyzer = PerformanceAnalyzer()

    # Example metrics
    analyzer.add_metric("LCP", 2300)
    analyzer.add_metric("FID", 45)
    analyzer.add_metric("CLS", 0.05)
    analyzer.add_metric("TTFB", 320)
    analyzer.add_metric("FCP", 1200)

    report = analyzer.generate_report()
    analyzer.print_summary()

    print(f"\nReport saved to: {analyzer.output_dir}/performance-report.json")
