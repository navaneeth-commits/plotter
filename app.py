from __future__ import annotations

import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from typing import Any

from dotenv import load_dotenv
from google import genai
from flask import Flask, jsonify, request

load_dotenv()  # reads the .env file (if present) and loads it into os.environ

app = Flask(__name__)

# Gemini setup — key is read from the environment, never hardcoded.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise RuntimeError(
        "GEMINI_API_KEY environment variable is not set. "
        "Set it before starting the server, e.g.\n"
        "  PowerShell: $env:GEMINI_API_KEY=\"your_key_here\"\n"
        "  bash/zsh:   export GEMINI_API_KEY=\"your_key_here\""
    )

client = genai.Client(api_key=GEMINI_API_KEY)
MODEL_NAME = "gemini-2.5-flash"

# CORS: the README has you open index.html directly as a file:// page,
# which sends "Origin: null" — so we can't lock this to a normal origin
# without also changing how the frontend is served. Left as "*" to match
# current usage, but since /eda triggers a billed LLM call, the safer
# long-term fix is to serve index.html from a local dev server (e.g.
# `python -m http.server`) and set EDA_ALLOWED_ORIGIN to that origin.
ALLOWED_ORIGIN = os.environ.get("EDA_ALLOWED_ORIGIN", "*")

ALLOWED_CHART_TYPES = {"bar", "line", "scatter", "histogram", "box", "pie"}
FALLBACK_RESPONSE = {"insights": ["Basic analysis fallback"], "charts": []}
LLM_TIMEOUT_SECONDS = 20


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    return response


def build_prompt(summary: dict[str, Any]) -> str:
    compact_summary = json.dumps(summary, ensure_ascii=True, separators=(",", ":"))
    return f"""
You are a local data analyst.

Return STRICT JSON only.
No markdown. No explanation.

Rules:
- Use only given columns
- y must be numeric
- Return 2 to 4 charts
- Allowed types: bar, line, scatter, histogram, box, pie

Format:
{{
  "insights": ["..."],
  "charts": [
    {{
      "type": "bar",
      "x": "column",
      "y": "column",
      "group": null
    }}
  ]
}}

Dataset:
{compact_summary}
"""


def fallback_response() -> dict[str, Any]:
    return {
        "insights": list(FALLBACK_RESPONSE["insights"]),
        "charts": list(FALLBACK_RESPONSE["charts"]),
    }


def extract_json(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    match = re.search(r"\{[\s\S]*\}", cleaned)
    return match.group(0) if match else None


def parse_llm_output(text: str) -> dict[str, Any]:
    json_text = extract_json(text)
    if not json_text:
        raise ValueError("No JSON found in LLM output")
    return json.loads(json_text)


def validate_chart(chart: dict[str, Any], column_types: dict[str, str]) -> dict[str, Any] | None:
    chart_type = str(chart.get("type", "")).strip().lower()
    x = chart.get("x")
    y = chart.get("y")
    group = chart.get("group")

    if chart_type not in ALLOWED_CHART_TYPES:
        return None

    if chart_type == "histogram":
        x = x or y
    if chart_type == "box" and group is None and column_types.get(x) in {"category", "date"}:
        group = x

    if x not in column_types or y not in column_types:
        return None
    if group is not None and group not in column_types:
        return None
    if column_types.get(y) != "numeric":
        return None

    return {"type": chart_type, "x": x, "y": y, "group": group}


def validate_response(data: dict[str, Any], summary: dict[str, Any]) -> dict[str, Any]:
    columns = summary.get("columns", [])
    column_types = {
        item.get("name"): item.get("type")
        for item in columns
        if isinstance(item, dict) and item.get("name") and item.get("type")
    }

    insights_raw = data.get("insights", [])
    charts_raw = data.get("charts", [])

    insights = [
        str(item).strip()
        for item in insights_raw
        if isinstance(item, str) and str(item).strip()
    ][:5]

    charts = []
    if isinstance(charts_raw, list):
        for item in charts_raw:
            if isinstance(item, dict):
                validated = validate_chart(item, column_types)
                if validated:
                    charts.append(validated)

    deduped = []
    seen = set()
    for chart in charts:
        key = (chart["type"], chart["x"], chart["y"], chart["group"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(chart)

    return {
        "insights": insights or list(FALLBACK_RESPONSE["insights"]),
        "charts": deduped[:4],
    }


# ✅ FIXED Gemini call
def _generate_content(prompt: str) -> str:
    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=prompt
    )

    try:
        return response.text or ""
    except Exception:
        return ""


def call_llm(prompt: str) -> str:
    executor = ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_generate_content, prompt)
    try:
        return future.result(timeout=LLM_TIMEOUT_SECONDS)
    except TimeoutError:
        future.cancel()
        raise
    finally:
        executor.shutdown(wait=False, cancel_futures=True)


@app.route("/eda", methods=["POST", "OPTIONS"])
def eda():
    if request.method == "OPTIONS":
        return jsonify(fallback_response())

    start = time.perf_counter()

    try:
        data = request.get_json(force=True) or {}
        print("\n===== REQUEST =====", flush=True)
        print("SUMMARY:", data, flush=True)
    except Exception as exc:
        print("\n===== INPUT ERROR =====", flush=True)
        print(str(exc), flush=True)
        print("TIME:", round(time.perf_counter() - start, 2), flush=True)
        return jsonify(fallback_response())

    summary = data.get("summary")
    if not isinstance(summary, dict):
        print("\n===== INPUT ERROR =====", flush=True)
        print("Missing or invalid summary payload", flush=True)
        print("TIME:", round(time.perf_counter() - start, 2), flush=True)
        return jsonify(fallback_response())

    try:
        prompt = build_prompt(summary)
        raw = call_llm(prompt)

        print("\n===== RAW LLM =====", flush=True)
        print(raw, flush=True)

        parsed = parse_llm_output(raw)
        response = validate_response(parsed, summary)

        print("\n===== SUCCESS =====", flush=True)
        print("TIME:", round(time.perf_counter() - start, 2), flush=True)
        return jsonify(response)

    except TimeoutError:
        print("\n===== TIMEOUT =====", flush=True)
        print("TIME:", round(time.perf_counter() - start, 2), flush=True)
        return jsonify(fallback_response())
    except Exception as exc:
        print("\n===== ERROR =====", flush=True)
        print(str(exc), flush=True)
        print("TIME:", round(time.perf_counter() - start, 2), flush=True)
        return jsonify(fallback_response())


app.config["MAX_CONTENT_LENGTH"] = 256 * 1024  # 256 KB — this payload is just column metadata + 5 sample rows

if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug_mode)