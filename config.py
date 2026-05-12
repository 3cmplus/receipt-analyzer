# -*- coding: utf-8 -*-
import os
from pathlib import Path

_ROOT = Path(__file__).resolve().parent

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / ".env", override=False)
except ImportError:
    _env_path = _ROOT / ".env"
    if _env_path.exists():
        for line in _env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            os.environ.setdefault(key.strip(), val.strip())


GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").lstrip("﻿").strip()
GEMINI_MODEL   = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash").strip()

if not GEMINI_API_KEY or GEMINI_API_KEY == "여기에_API_키를_입력하세요":
    raise EnvironmentError(".env 파일에 GEMINI_API_KEY를 입력해주세요.")
