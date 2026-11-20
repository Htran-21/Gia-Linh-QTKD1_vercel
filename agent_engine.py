#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EdTech Future — AI Agent Điều phối Gia hạn & Chăm sóc Học viên (Task 2)
Kiến trúc tích hợp 4 Kỹ năng Nghiệp vụ (Agent Skills) & Gemini/OpenAI Function Calling:
  - Skill 1: Multi-Source Data Ingestion & Cleaning (.xlsx, Google Sheets, Lark Base, JSON)
  - Skill 2: Cross-Check Student Eligibility, Contact Window & Tech Ticket Blocker
  - Skill 3: Deterministic Package Recommendation & Policy Pricing Engine
  - Skill 4: LLM Proposal & Coordination Drafting + Anti-Hallucination Guardrail & Audit Logger
"""

import json
import os
import urllib.request
from datetime import datetime, timedelta
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUTS_DIR = BASE_DIR / "outputs"
LOGS_DIR = BASE_DIR / "logs"
WEB_DIR = BASE_DIR / "web"

for d in (DATA_DIR, OUTPUTS_DIR, LOGS_DIR):
    d.mkdir(parents=True, exist_ok=True)


# ==============================================================================
# AGENT FUNCTION CALLING TOOL DECLARATIONS (GEMINI / OPENAI SCHEMA)
# ==============================================================================
AGENT_TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "ingest_and_clean_multisource_data",
            "description": "Đọc dữ liệu từ 6 bảng (.xlsx, Google Sheets, Lark Base, JSON), chuẩn hóa theo mốc as_of_date và phân biệt rõ dữ liệu thiếu (NULL) với chuyên cần bằng 0.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_type": {"type": "string", "enum": ["xlsx", "gsheets", "larkbase", "json"]},
                    "source_uri": {"type": "string"},
                    "as_of_date": {"type": "string", "description": "Mốc ngày đối chiếu chuẩn YYYY-MM-DD"}
                },
                "required": ["source_type", "as_of_date"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "cross_check_student_eligibility_and_risks",
            "description": "Đối chiếu tài khoản, tính ngày bắt đầu liên hệ (trước hết hạn 30 ngày), kiểm tra tỷ lệ thực học và chặn cứng nếu có Ticket kỹ thuật đang Open/In Progress.",
            "parameters": {
                "type": "object",
                "properties": {
                    "account_id": {"type": "string"},
                    "as_of_date": {"type": "string"}
                },
                "required": ["account_id", "as_of_date"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_package_and_policy_pricing",
            "description": "Đề xuất giữ gói cũ hoặc nâng cấp gói mới, kiểm tra điều kiện chính sách ưu đãi (EARLY_BIRD_10, HIGH_ACHIEVER_15, LOYALTY_VIP_20) và tính học phí chính xác.",
            "parameters": {
                "type": "object",
                "properties": {
                    "account_id": {"type": "string"},
                    "requested_discount_code": {"type": "string"}
                },
                "required": ["account_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "validate_output_guardrails",
            "description": "Kiểm tra hậu kiểm chống ảo giác: cấm cam kết đã sửa xong lỗi kỹ thuật khi ticket còn mở, cấm cam kết đã duyệt ưu đãi khi chờ Quản lý phê duyệt.",
            "parameters": {
                "type": "object",
                "properties": {
                    "account_id": {"type": "string"},
                    "draft_message": {"type": "string"},
                    "is_blocked_by_tech": {"type": "boolean"},
                    "requires_manager_approval": {"type": "boolean"}
                },
                "required": ["account_id", "draft_message"]
            }
        }
    }
]


def log_audit_event(run_id: str, step: str, skill: str, tool: str, account_id: str, rules: list, message: str, details: dict = None):
    ts = datetime.now().astimezone().isoformat()
    entry = {
        "run_id": run_id,
        "timestamp": ts,
        "step": step,
        "skill_name": skill,
        "tool_called": tool,
        "account_id": account_id,
        "rules_applied": rules,
        "status": "SUCCESS",
        "message": message,
        "details": details or {}
    }
    with open(LOGS_DIR / "agent_execution_audit.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    print("EdTech Future AI Agent Engine initialized. Use start_localhost.bat or server.ps1 to launch on http://localhost:8080")
