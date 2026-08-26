#!/usr/bin/env python3
"""
QA Ultra Tester - Test Data Generator
Generates synthetic test data for comprehensive testing
"""

import json
import random
import string
from datetime import datetime, timedelta
from pathlib import Path


class TestDataGenerator:
    def __init__(self, output_dir="hasil-test/test-data"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def random_string(self, length=10):
        return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

    def random_email(self):
        domains = ["gmail.com", "yahoo.com", "outlook.com", "test.com", "example.org"]
        return f"{self.random_string(8)}@{random.choice(domains)}"

    def random_phone(self):
        return f"+62{random.randint(8100000000, 8999999999)}"

    def random_date(self, start_year=2020, end_year=2025):
        start = datetime(start_year, 1, 1)
        end = datetime(end_year, 12, 31)
        delta = end - start
        random_days = random.randint(0, delta.days)
        return (start + timedelta(days=random_days)).strftime("%Y-%m-%d")

    def generate_users(self, count=10):
        users = []
        for i in range(count):
            users.append({
                "id": i + 1,
                "name": f"User {self.random_string(6)}",
                "email": self.random_email(),
                "phone": self.random_phone(),
                "address": f"Jl. {self.random_string(8)} No. {random.randint(1, 100)}",
                "city": random.choice(["Jakarta", "Surabaya", "Bandung", "Yogyakarta", "Medan"]),
                "created_at": self.random_date()
            })
        return users

    def generate_boundary_data(self):
        return {
            "strings": {
                "empty": "",
                "single_char": "a",
                "max_length": "x" * 10000,
                "special_chars": "!@#$%^&*()_+-=[]{}|;':\",./<>?",
                "unicode_chinese": "测试中文字符串",
                "unicode_arabic": "اختبار النص العربي",
                "unicode_emoji": "🔴🟡🟢🔵🟣⚫⚪",
                "sql_injection": "' OR '1'='1'; DROP TABLE users;--",
                "xss_basic": "<script>alert('XSS')</script>",
                "xss_event": "<img src=x onerror=alert(1)>",
                "xss_svg": "<svg onload=alert(1)>",
                "path_traversal": "../../../etc/passwd",
                "null_byte": "test%00.jpg",
                "html_tags": "<div><p>HTML injection</p></div>",
                "json_payload": '{"__proto__": {"admin": true}}',
                "xml_bomb": "<!DOCTYPE foo [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]>"
            },
            "numbers": {
                "zero": 0,
                "negative": -123,
                "max_int": 2147483647,
                "overflow": 99999999999999999,
                "decimal": 3.141592653589793,
                "scientific": "1e308"
            },
            "emails": {
                "valid": "user@example.com",
                "no_at": "userexample.com",
                "no_domain": "user@",
                "double_at": "user@@example.com",
                "special_chars": "user.name+tag@example.com",
                "long_domain": "user@" + "a" * 255 + ".com"
            },
            "dates": {
                "past": "2020-01-01",
                "today": datetime.now().strftime("%Y-%m-%d"),
                "future": "2030-12-31",
                "invalid_format": "2025/13/45",
                "leap_year": "2024-02-29",
                "non_leap_feb29": "2023-02-29"
            }
        }

    def generate_api_test_data(self):
        return {
            "endpoints": [
                {
                    "method": "GET",
                    "path": "/api/users",
                    "tests": [
                        {"name": "Valid request", "headers": {"Authorization": "Bearer token"}, "expected_status": 200},
                        {"name": "No auth", "headers": {}, "expected_status": 401},
                        {"name": "Invalid token", "headers": {"Authorization": "Bearer invalid"}, "expected_status": 401}
                    ]
                },
                {
                    "method": "POST",
                    "path": "/api/users",
                    "tests": [
                        {"name": "Valid create", "body": {"name": "Test", "email": "test@example.com"}, "expected_status": 201},
                        {"name": "Missing required field", "body": {"name": "Test"}, "expected_status": 400},
                        {"name": "Duplicate email", "body": {"name": "Test", "email": "existing@example.com"}, "expected_status": 409}
                    ]
                }
            ]
        }

    def generate_fixtures(self):
        return {
            "users": self.generate_users(5),
            "boundary_data": self.generate_boundary_data(),
            "api_test_data": self.generate_api_test_data(),
            "test_urls": {
                "login": "/login",
                "dashboard": "/dashboard",
                "settings": "/settings",
                "profile": "/profile"
            }
        }

    def save_fixtures(self, filename="fixtures.json"):
        fixtures = self.generate_fixtures()
        filepath = self.output_dir / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(fixtures, f, indent=2, ensure_ascii=False)
        return filepath

    def save_users(self, filename="users.json", count=10):
        users = self.generate_users(count)
        filepath = self.output_dir / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=2, ensure_ascii=False)
        return filepath


if __name__ == "__main__":
    generator = TestDataGenerator()

    print("Generating test data...")
    fixtures_path = generator.save_fixtures()
    users_path = generator.save_users(count=20)

    print(f"Fixtures saved to: {fixtures_path}")
    print(f"Users saved to: {users_path}")
    print("Test data generation complete!")
