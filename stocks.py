import html
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openai import OpenAI
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://consensus.hankyung.com/"
LIST_URL = f"{BASE_URL}analysis/list"
REPORT_TYPES = {
    "CO": "기업",
    "IN": "산업",
    "MA": "시장",
}

OUTPUT_ROOT = Path(".")
PDF_ROOT = OUTPUT_ROOT / "reports_pdf"
JSON_ROOT = OUTPUT_ROOT / "reports_json"
MOBILE_ROOT = OUTPUT_ROOT / "reports_mobile"
EXCEL_PREFIX = "hankyung_consensus_summary"

PAGE_SIZE = 20
MAX_PDF_PAGES = int(os.getenv("HANKYUNG_MAX_PAGES", "5"))
MODEL_NAME = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
REQUEST_SLEEP_SECONDS = float(os.getenv("HANKYUNG_SLEEP_SECONDS", "1"))
TARGET_DATE_ENV = os.getenv("HANKYUNG_TARGET_DATE")

SUMMARY_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "report_type": {"type": "string", "enum": list(REPORT_TYPES.values())},
        "company_name": {"type": ["string", "null"]},
        "stock_code": {"type": ["string", "null"]},
        "industry_name": {"type": ["string", "null"]},
        "market_topic": {"type": ["string", "null"]},
        "investment_opinion": {"type": ["string", "null"]},
        "target_price": {"type": ["string", "null"]},
        "current_price": {"type": ["string", "null"]},
        "earnings_momentum": {"type": ["string", "null"]},
        "key_points": {"type": "array", "items": {"type": "string"}},
        "financial_metrics": {"type": "array", "items": {"type": "string"}},
        "valuation": {"type": ["string", "null"]},
        "risks": {"type": "array", "items": {"type": "string"}},
        "summary": {"type": "string"},
        "target_price_revision_direction": {"type": ["string", "null"]},
        "target_price_revision_rate": {"type": ["number", "null"]},
        "investment_opinion_revision_direction": {"type": ["string", "null"]},
        "earnings_revision_positive": {"type": "boolean"},
        "structural_growth_positive": {"type": "boolean"},
        "industry_cycle_positive": {"type": "boolean"},
        "supply_shortage_positive": {"type": "boolean"},
        "asp_rising_positive": {"type": "boolean"},
        "capex_cycle_positive": {"type": "boolean"},
        "backlog_positive": {"type": "boolean"},
        "new_customer_positive": {"type": "boolean"},
        "market_share_positive": {"type": "boolean"},
        "valuation_risk_negative": {"type": "boolean"},
        "front_running_negative": {"type": "boolean"},
        "peakout_negative": {"type": "boolean"},
        "margin_slowdown_negative": {"type": "boolean"},
        "demand_slowdown_negative": {"type": "boolean"},
    },
    "required": [
        "report_type",
        "company_name",
        "stock_code",
        "industry_name",
        "market_topic",
        "investment_opinion",
        "target_price",
        "current_price",
        "earnings_momentum",
        "key_points",
        "financial_metrics",
        "valuation",
        "risks",
        "summary",
        "target_price_revision_direction",
        "target_price_revision_rate",
        "investment_opinion_revision_direction",
        "earnings_revision_positive",
        "structural_growth_positive",
        "industry_cycle_positive",
        "supply_shortage_positive",
        "asp_rising_positive",
        "capex_cycle_positive",
        "backlog_positive",
        "new_customer_positive",
        "market_share_positive",
        "valuation_risk_negative",
        "front_running_negative",
        "peakout_negative",
        "margin_slowdown_negative",
        "demand_slowdown_negative",
    ],
    "additionalProperties": False,
}

MARKET_REACTION_TEMPLATE = {
    "status": "pending",
    "d1_return_pct": None,
    "d5_return_pct": None,
    "d20_return_pct": None,
}

SCORING_REQUIRED_KEYS = {
    "alpha_score",
    "alpha_grade",
    "score_reason",
    "positive_factors",
    "negative_factors",
    "revision_signal",
    "structural_growth_signal",
    "valuation_risk_signal",
    "market_reaction_placeholder",
    "target_price_revision_direction",
    "target_price_revision_rate",
    "investment_opinion_revision_direction",
    "earnings_revision_positive",
    "structural_growth_positive",
    "industry_cycle_positive",
    "supply_shortage_positive",
    "asp_rising_positive",
    "capex_cycle_positive",
    "backlog_positive",
    "new_customer_positive",
    "market_share_positive",
    "valuation_risk_negative",
    "front_running_negative",
    "peakout_negative",
    "margin_slowdown_negative",
    "demand_slowdown_negative",
}


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/136.0.0.0 Safari/537.36"
            ),
            "Referer": BASE_URL,
        }
    )
    retry = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def category_excel_path(category: str) -> Path:
    return OUTPUT_ROOT / f"{EXCEL_PREFIX}_{category}.xlsx"


def ensure_output_dirs() -> None:
    MOBILE_ROOT.mkdir(parents=True, exist_ok=True)
    for category in REPORT_TYPES.values():
        (PDF_ROOT / category).mkdir(parents=True, exist_ok=True)
        (JSON_ROOT / category).mkdir(parents=True, exist_ok=True)


def clean_text(value: str | None) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", value).strip()


def to_none_if_blank(value: str | None) -> str | None:
    cleaned = clean_text(value)
    return cleaned or None


def split_company_info(title: str) -> tuple[str | None, str | None]:
    match = re.search(r"^(?P<name>.+?)\((?P<code>\d{6})\)", title)
    if not match:
        return None, None
    return clean_text(match.group("name")), match.group("code")


def parse_report_row(row, report_code: str) -> dict[str, Any] | None:
    cells = row.find_all("td", recursive=False)
    if not cells:
        return None

    if len(cells) == 1 and "결과가 없습니다." in clean_text(cells[0].get_text(" ", strip=True)):
        return None

    title_link = row.select_one("a[href*='report_idx=']")
    if title_link is None:
        return None

    report_match = re.search(r"report_idx=(\d+)", title_link.get("href", ""))
    if report_match is None:
        return None

    report_idx = report_match.group(1)
    category = REPORT_TYPES[report_code]
    title = clean_text(title_link.get_text(" ", strip=True))
    published_at = clean_text(cells[0].get_text(" ", strip=True))

    item: dict[str, Any] = {
        "report_idx": report_idx,
        "report_type_code": report_code,
        "category": category,
        "published_at": published_at,
        "title": title,
        "author": None,
        "source": None,
        "industry_name": None,
        "company_name": None,
        "stock_code": None,
        "investment_opinion": None,
        "target_price": None,
        "pdf_url": f"{BASE_URL}analysis/downpdf?report_idx={report_idx}",
    }

    if report_code == "CO" and len(cells) >= 9:
        item["target_price"] = to_none_if_blank(cells[2].get_text(" ", strip=True))
        item["investment_opinion"] = to_none_if_blank(cells[3].get_text(" ", strip=True))
        item["author"] = to_none_if_blank(cells[4].get_text(" ", strip=True))
        item["source"] = to_none_if_blank(cells[5].get_text(" ", strip=True))
        company_name, stock_code = split_company_info(title)
        item["company_name"] = company_name
        item["stock_code"] = stock_code
        return item

    if report_code == "IN" and len(cells) >= 7:
        item["industry_name"] = to_none_if_blank(cells[2].get_text(" ", strip=True))
        item["author"] = to_none_if_blank(cells[3].get_text(" ", strip=True))
        item["source"] = to_none_if_blank(cells[4].get_text(" ", strip=True))
        return item

    if report_code == "MA" and len(cells) >= 6:
        item["author"] = to_none_if_blank(cells[2].get_text(" ", strip=True))
        item["source"] = to_none_if_blank(cells[3].get_text(" ", strip=True))
        return item

    raise ValueError(
        f"Unexpected table layout for {category} report {report_idx}: {len(cells)} cells"
    )


def parse_date(value: str) -> datetime.date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def get_latest_available_date(session: requests.Session) -> str:
    response = session.get(BASE_URL, timeout=30)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    for row in soup.select("table tbody tr"):
        cells = row.find_all("td", recursive=False)
        if not cells:
            continue
        candidate = clean_text(cells[0].get_text(" ", strip=True))
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", candidate):
            return candidate
    raise ValueError("Could not detect latest available report date.")


def get_target_date(session: requests.Session) -> str:
    if TARGET_DATE_ENV:
        parse_date(TARGET_DATE_ENV)
        return TARGET_DATE_ENV
    return get_latest_available_date(session)


def get_report_list_for_date(
    session: requests.Session,
    report_code: str,
    target_date: str,
) -> list[dict[str, Any]]:
    reports: list[dict[str, Any]] = []
    page = 1

    while True:
        response = session.get(
            LIST_URL,
            params={
                "report_type": report_code,
                "sdate": target_date,
                "edate": target_date,
                "now_page": str(page),
                "pagenum": str(PAGE_SIZE),
            },
            timeout=30,
        )
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        rows = soup.select("table tbody tr")
        page_reports = [parse_report_row(row, report_code) for row in rows]
        page_reports = [report for report in page_reports if report is not None]

        if not page_reports:
            break

        reports.extend(page_reports)
        if len(page_reports) < PAGE_SIZE:
            break

        page += 1

    return reports


def download_pdf(session: requests.Session, pdf_url: str, save_path: Path) -> None:
    response = session.get(pdf_url, timeout=60, allow_redirects=True)
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()
    content = response.content
    if "pdf" not in content_type and not content.startswith(b"%PDF"):
        preview = response.text[:300]
        raise ValueError(
            f"Expected PDF but received Content-Type={content_type!r}, preview={preview!r}"
        )

    save_path.parent.mkdir(parents=True, exist_ok=True)
    save_path.write_bytes(content)


def extract_pdf_text(pdf_path: Path, max_pages: int) -> str:
    texts: list[str] = []
    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc):
            if page_index >= max_pages:
                break
            texts.append(page.get_text())
    return "\n".join(texts).strip()


def strip_code_fence(text: str) -> str:
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped)
        stripped = re.sub(r"\s*```$", "", stripped)
    return stripped.strip()


def repair_json_string(text: str) -> str:
    repaired = text
    repaired = repaired.replace("\ufeff", "")
    repaired = repaired.replace("“", '"').replace("”", '"')
    repaired = repaired.replace("’", "'").replace("‘", "'")
    repaired = re.sub(r",(\s*[}\]])", r"\1", repaired)
    return repaired.strip()


def parse_json_response(raw_text: str) -> dict[str, Any]:
    candidates: list[str] = []
    stripped = strip_code_fence(raw_text)
    candidates.append(stripped)

    start = stripped.find("{")
    end = stripped.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidates.append(stripped[start : end + 1])

    for candidate in candidates:
        for attempt in (candidate, repair_json_string(candidate)):
            if not attempt:
                continue
            try:
                parsed = json.loads(attempt)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                return parsed

    raise ValueError("Failed to parse JSON payload from model response.")


def dedupe_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        cleaned = clean_text(value)
        if not cleaned:
            continue
        lowered = cleaned.casefold()
        if lowered in seen:
            continue
        seen.add(lowered)
        result.append(cleaned)
    return result


def normalize_market_reaction_placeholder(value: Any) -> dict[str, Any]:
    payload = dict(MARKET_REACTION_TEMPLATE)
    if isinstance(value, dict):
        for key in payload:
            payload[key] = value.get(key)
        if not payload.get("status"):
            payload["status"] = "pending"
    return payload


def clamp_score(value: float) -> int:
    return int(max(0, min(100, round(value))))


def alpha_grade_for(score: int | None) -> str | None:
    if score is None:
        return None
    if score >= 85:
        return "S"
    if score >= 70:
        return "A"
    if score >= 55:
        return "B"
    if score >= 40:
        return "C"
    return "D"


def summarize_revision_signal(summary: dict[str, Any]) -> str | None:
    parts: list[str] = []
    target_direction = clean_text(summary.get("target_price_revision_direction"))
    target_rate = summary.get("target_price_revision_rate")
    if target_direction == "up":
        if isinstance(target_rate, (int, float)) and target_rate > 0:
            parts.append(f"목표주가 상향(+{target_rate:.1f}%)")
        else:
            parts.append("목표주가 상향")
    elif target_direction == "down":
        parts.append("목표주가 하향")

    opinion_direction = clean_text(summary.get("investment_opinion_revision_direction"))
    if opinion_direction == "up":
        parts.append("투자의견 상향")
    elif opinion_direction == "down":
        parts.append("투자의견 하향")

    if summary.get("earnings_revision_positive"):
        parts.append("EPS/영업이익/매출 추정치 상향")
    return ", ".join(parts) if parts else None


def summarize_structural_growth_signal(summary: dict[str, Any]) -> str | None:
    labels: list[str] = []
    if summary.get("structural_growth_positive"):
        labels.append("구조적 성장")
    if summary.get("industry_cycle_positive"):
        labels.append("산업 업황 개선")
    if summary.get("supply_shortage_positive"):
        labels.append("공급 부족")
    if summary.get("asp_rising_positive"):
        labels.append("ASP 상승")
    if summary.get("capex_cycle_positive"):
        labels.append("CAPEX 사이클")
    if summary.get("backlog_positive"):
        labels.append("수주잔고 확대")
    if summary.get("new_customer_positive"):
        labels.append("신규 고객")
    if summary.get("market_share_positive"):
        labels.append("시장점유율 확대")
    labels = dedupe_preserve_order(labels)
    return ", ".join(labels) if labels else None


def summarize_valuation_risk_signal(summary: dict[str, Any]) -> str | None:
    labels: list[str] = []
    if summary.get("valuation_risk_negative"):
        labels.append("valuation 부담")
    if summary.get("front_running_negative"):
        labels.append("주가 선반영")
    if summary.get("peakout_negative"):
        labels.append("피크아웃 우려")
    if summary.get("margin_slowdown_negative"):
        labels.append("마진 둔화")
    if summary.get("demand_slowdown_negative"):
        labels.append("수요 둔화")
    labels = dedupe_preserve_order(labels)
    return ", ".join(labels) if labels else None


def compute_company_alpha(summary: dict[str, Any]) -> dict[str, Any]:
    score = 50.0
    positives: list[str] = list(summary.get("positive_factors") or [])
    negatives: list[str] = list(summary.get("negative_factors") or [])

    target_direction = clean_text(summary.get("target_price_revision_direction"))
    target_rate = summary.get("target_price_revision_rate")
    if target_direction == "up":
        score += 10
        positives.append("목표주가 상향")
        if isinstance(target_rate, (int, float)) and target_rate > 0:
            rate_bonus = min(float(target_rate), 30.0) * 0.35
            score += rate_bonus
            positives.append(f"목표주가 상향률 +{float(target_rate):.1f}%")
    elif target_direction == "down":
        score -= 9
        negatives.append("목표주가 하향")

    opinion_direction = clean_text(summary.get("investment_opinion_revision_direction"))
    if opinion_direction == "up":
        score += 8
        positives.append("투자의견 상향")
    elif opinion_direction == "down":
        score -= 8
        negatives.append("투자의견 하향")

    if summary.get("earnings_revision_positive"):
        score += 12
        positives.append("EPS/영업이익/매출 추정치 상향")

    if summary.get("structural_growth_positive"):
        score += 10
        positives.append("실적 개선이 구조적 성장에 기반")

    cycle_bonus = 0
    cycle_factors = [
        ("industry_cycle_positive", "산업 업황 개선"),
        ("supply_shortage_positive", "공급 부족"),
        ("asp_rising_positive", "ASP 상승"),
        ("capex_cycle_positive", "CAPEX cycle"),
        ("backlog_positive", "수주잔고 확대"),
        ("new_customer_positive", "신규 고객"),
        ("market_share_positive", "시장점유율 확대"),
    ]
    for key, label in cycle_factors:
        if summary.get(key):
            cycle_bonus += 4
            positives.append(label)
    score += min(cycle_bonus, 20)

    risk_penalties = [
        ("valuation_risk_negative", 10, "valuation 부담"),
        ("front_running_negative", 6, "주가 선반영"),
        ("peakout_negative", 10, "피크아웃 우려"),
        ("margin_slowdown_negative", 7, "마진 둔화"),
        ("demand_slowdown_negative", 8, "수요 둔화"),
    ]
    for key, penalty, label in risk_penalties:
        if summary.get(key):
            score -= penalty
            negatives.append(label)

    final_score = clamp_score(score)
    positives = dedupe_preserve_order(positives)
    negatives = dedupe_preserve_order(negatives)

    reason_parts: list[str] = []
    if positives:
        reason_parts.append("긍정: " + ", ".join(positives[:4]))
    if negatives:
        reason_parts.append("부정: " + ", ".join(negatives[:3]))

    return {
        "alpha_score": final_score,
        "alpha_grade": alpha_grade_for(final_score),
        "score_reason": " / ".join(reason_parts) if reason_parts else "명확한 alpha 신호가 제한적임",
        "positive_factors": positives,
        "negative_factors": negatives,
        "revision_signal": summarize_revision_signal(summary),
        "structural_growth_signal": summarize_structural_growth_signal(summary),
        "valuation_risk_signal": summarize_valuation_risk_signal(summary),
        "market_reaction_placeholder": normalize_market_reaction_placeholder(
            summary.get("market_reaction_placeholder")
        ),
    }


def enrich_summary(report: dict[str, Any], raw_summary: dict[str, Any]) -> dict[str, Any]:
    summary = dict(raw_summary)
    defaults: dict[str, Any] = {
        "report_type": report["category"],
        "company_name": None,
        "stock_code": None,
        "industry_name": None,
        "market_topic": None,
        "investment_opinion": None,
        "target_price": None,
        "current_price": None,
        "earnings_momentum": None,
        "key_points": [],
        "financial_metrics": [],
        "valuation": None,
        "risks": [],
        "summary": "",
        "target_price_revision_direction": None,
        "target_price_revision_rate": None,
        "investment_opinion_revision_direction": None,
        "earnings_revision_positive": False,
        "structural_growth_positive": False,
        "industry_cycle_positive": False,
        "supply_shortage_positive": False,
        "asp_rising_positive": False,
        "capex_cycle_positive": False,
        "backlog_positive": False,
        "new_customer_positive": False,
        "market_share_positive": False,
        "valuation_risk_negative": False,
        "front_running_negative": False,
        "peakout_negative": False,
        "margin_slowdown_negative": False,
        "demand_slowdown_negative": False,
        "positive_factors": [],
        "negative_factors": [],
        "revision_signal": None,
        "structural_growth_signal": None,
        "valuation_risk_signal": None,
        "score_reason": None,
        "alpha_score": None,
        "alpha_grade": None,
        "market_reaction_placeholder": dict(MARKET_REACTION_TEMPLATE),
    }
    for key, default in defaults.items():
        if key not in summary:
            summary[key] = default if not isinstance(default, dict) else dict(default)

    summary["report_type"] = report["category"]
    summary["market_reaction_placeholder"] = normalize_market_reaction_placeholder(
        summary.get("market_reaction_placeholder")
    )

    if report["category"] == "기업":
        summary.update(compute_company_alpha(summary))
    else:
        summary["alpha_score"] = None
        summary["alpha_grade"] = None
        summary["score_reason"] = None
        summary["positive_factors"] = dedupe_preserve_order(summary.get("positive_factors", []))
        summary["negative_factors"] = dedupe_preserve_order(summary.get("negative_factors", []))
        summary["revision_signal"] = None
        summary["structural_growth_signal"] = None
        summary["valuation_risk_signal"] = None
    return summary


def summary_has_scoring_fields(summary: dict[str, Any]) -> bool:
    return SCORING_REQUIRED_KEYS.issubset(summary.keys())


def build_prompt(report: dict[str, Any], pdf_text: str) -> str:
    if report["category"] == "기업":
        focus = (
            "기업 리포트다. 회사명, 종목코드, 투자의견, 목표주가, 실적 모멘텀, "
            "핵심 포인트, 밸류에이션, 리스크를 정리하라. "
            "또한 목표주가/투자의견 상향 여부, 이익추정 상향 여부, 구조적 성장 여부, "
            "업황/공급부족/ASP/CAPEX/수주잔고/신규고객/점유율 확대 신호와 "
            "valuation 부담/선반영/피크아웃/마진둔화/수요둔화 리스크를 명시적으로 판단하라."
        )
    elif report["category"] == "산업":
        focus = (
            "산업 리포트다. 개별 종목 목표주가를 억지로 채우지 말고, 산업명/업황/수급/"
            "정책/체인 이슈 중심으로 요약하라. 기업 점수용 revision/alpha 신호는 null 또는 false로 둬라."
        )
    else:
        focus = (
            "시장 리포트다. 개별 종목 의견을 억지로 만들지 말고, 시장 주제/매크로/수급/"
            "전략 관점으로 요약하라. 기업 점수용 revision/alpha 신호는 null 또는 false로 둬라."
        )

    return f"""
다음 한경 컨센서스 리포트를 JSON 스키마에 맞춰 요약하라.
{focus}
값이 없으면 null 또는 빈 배열을 사용하고, 추정하지 마라.

[메타데이터]
- 분류: {report["category"]}
- 제목: {report["title"]}
- 작성자: {report.get("author")}
- 제공출처: {report.get("source")}
- 회사명: {report.get("company_name")}
- 종목코드: {report.get("stock_code")}
- 산업명: {report.get("industry_name")}
- 목록상 목표주가: {report.get("target_price")}
- 목록상 투자의견: {report.get("investment_opinion")}

[본문]
{pdf_text[:15000]}
""".strip()


def summarize_report(
    client: OpenAI,
    report: dict[str, Any],
    pdf_text: str,
) -> dict[str, Any]:
    prompt = build_prompt(report, pdf_text)

    response = client.responses.create(
        model=MODEL_NAME,
        input=prompt,
        text={
            "format": {
                "type": "json_schema",
                "name": "hankyung_report_summary",
                "schema": SUMMARY_SCHEMA,
                "strict": True,
            }
        },
    )

    try:
        return parse_json_response(response.output_text)
    except ValueError:
        repaired = client.responses.create(
            model=MODEL_NAME,
            input=(
                "아래 텍스트를 같은 JSON 스키마에 맞는 유효한 JSON 한 개로만 복구하라.\n\n"
                f"{response.output_text}"
            ),
            text={
                "format": {
                    "type": "json_schema",
                    "name": "hankyung_report_summary_repair",
                    "schema": SUMMARY_SCHEMA,
                    "strict": True,
                }
            },
        )
        return parse_json_response(repaired.output_text)


def to_excel_row(report: dict[str, Any], summary: dict[str, Any], pdf_path: Path) -> dict[str, Any]:
    return {
        "수집일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "report_idx": str(report["report_idx"]),
        "분류": report["category"],
        "발행일": report["published_at"],
        "제목": report["title"],
        "작성자": report.get("author"),
        "제공출처": report.get("source"),
        "회사명": summary.get("company_name") or report.get("company_name"),
        "종목코드": summary.get("stock_code") or report.get("stock_code"),
        "산업명": summary.get("industry_name") or report.get("industry_name"),
        "시장주제": summary.get("market_topic"),
        "투자의견": summary.get("investment_opinion") or report.get("investment_opinion"),
        "목표주가": summary.get("target_price") or report.get("target_price"),
        "현재주가": summary.get("current_price"),
        "실적모멘텀": summary.get("earnings_momentum"),
        "핵심포인트": "\n".join(summary.get("key_points", [])),
        "재무지표": "\n".join(summary.get("financial_metrics", [])),
        "밸류에이션": summary.get("valuation"),
        "리스크": "\n".join(summary.get("risks", [])),
        "요약": summary.get("summary"),
        "alpha_score": summary.get("alpha_score"),
        "alpha_grade": summary.get("alpha_grade"),
        "score_reason": summary.get("score_reason"),
        "positive_factors": "\n".join(summary.get("positive_factors", [])),
        "negative_factors": "\n".join(summary.get("negative_factors", [])),
        "revision_signal": summary.get("revision_signal"),
        "structural_growth_signal": summary.get("structural_growth_signal"),
        "valuation_risk_signal": summary.get("valuation_risk_signal"),
        "market_reaction_placeholder": json.dumps(
            summary.get("market_reaction_placeholder", MARKET_REACTION_TEMPLATE),
            ensure_ascii=False,
        ),
        "pdf_path": str(pdf_path),
        "pdf_url": report["pdf_url"],
    }


def save_json(payload: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_existing_ids(excel_path: Path) -> set[str]:
    if not excel_path.exists():
        return set()
    existing_df = pd.read_excel(excel_path, dtype={"report_idx": str})
    if "report_idx" not in existing_df.columns:
        return set()
    return set(existing_df["report_idx"].dropna().astype(str))


def append_rows_to_excel(rows: list[dict[str, Any]], excel_path: Path) -> int:
    new_df = pd.DataFrame(rows)
    if excel_path.exists():
        old_df = pd.read_excel(excel_path, dtype={"report_idx": str})
        final_df = pd.concat([old_df, new_df], ignore_index=True, sort=False)
    else:
        final_df = new_df

    if not final_df.empty and "report_idx" in final_df.columns:
        final_df["report_idx"] = final_df["report_idx"].astype(str)
        final_df = final_df.drop_duplicates(subset=["report_idx"], keep="last")

    final_df.to_excel(excel_path, index=False)
    return len(final_df)


def summarize_pending_reports(
    session: requests.Session,
    target_date: str,
) -> dict[str, Any]:
    category_summaries: list[dict[str, Any]] = []
    total_reports = 0
    new_report_count = 0

    for report_code, category in REPORT_TYPES.items():
        reports = get_report_list_for_date(session, report_code, target_date)
        existing_ids = load_existing_ids(category_excel_path(category))
        new_reports = [
            report for report in reports if str(report["report_idx"]) not in existing_ids
        ]
        total_reports += len(reports)
        new_report_count += len(new_reports)
        category_summaries.append(
            {
                "category": category,
                "total_reports": len(reports),
                "new_reports": len(new_reports),
                "new_report_ids": [str(report["report_idx"]) for report in new_reports],
            }
        )

    return {
        "target_date": target_date,
        "total_reports": total_reports,
        "new_report_count": new_report_count,
        "has_new_reports": new_report_count > 0,
        "categories": category_summaries,
    }


def fetch_current_quote(session: requests.Session, stock_code: str) -> dict[str, Any]:
    response = session.get(
        "https://finance.naver.com/item/main.naver",
        params={"code": stock_code},
        timeout=20,
    )
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    price_node = soup.select_one("div.rate_info div.today p.no_today span.blind")
    if price_node is None:
        raise ValueError(f"Failed to parse current price for stock {stock_code}")

    diff_nodes = soup.select("div.rate_info div.today p.no_exday span.blind")
    quote_as_of = None
    for node in soup.select("div.description span"):
        text = clean_text(node.get_text(" ", strip=True))
        if "기준" in text:
            quote_as_of = text
            break

    direction = None
    icon = soup.select_one("div.rate_info div.today p.no_exday span.ico")
    if icon is not None:
        classes = icon.get("class", [])
        for candidate in ("up", "down", "same"):
            if candidate in classes:
                direction = candidate
                break

    change_amount = clean_text(diff_nodes[0].get_text(" ", strip=True)) if len(diff_nodes) >= 1 else None
    change_rate = clean_text(diff_nodes[1].get_text(" ", strip=True)) if len(diff_nodes) >= 2 else None
    if change_rate and not change_rate.endswith("%"):
        change_rate = f"{change_rate}%"

    page_text = soup.get_text(" ", strip=True)
    high_52w = None
    for pattern in (
        r"52주\s*최고\s*([0-9,]+)",
        r"52주\s*최고가\s*([0-9,]+)",
    ):
        matched = re.search(pattern, page_text)
        if matched:
            high_52w = clean_text(matched.group(1))
            break

    return {
        "live_price": clean_text(price_node.get_text(" ", strip=True)),
        "live_price_change": change_amount,
        "live_price_rate": change_rate,
        "live_price_direction": direction,
        "live_price_as_of": quote_as_of,
        "live_price_high_52w": high_52w,
    }


def fetch_price_history(
    session: requests.Session,
    stock_code: str,
    count: int = 90,
) -> list[dict[str, Any]]:
    endpoints = [
        (
            "https://fchart.stock.naver.com/sise.nhn",
            {"symbol": stock_code, "timeframe": "day", "count": str(count), "requestType": "0"},
        ),
        (
            "https://finance.naver.com/item/fchart.naver",
            {"code": stock_code, "timeframe": "day", "count": str(count), "requestType": "0"},
        ),
    ]

    for url, params in endpoints:
        try:
            response = session.get(url, params=params, timeout=20)
            response.raise_for_status()

            points: list[dict[str, Any]] = []
            for matched in re.finditer(
                r'data="(?P<date>\d{8})\|(?P<open>[0-9.]+)\|(?P<high>[0-9.]+)\|(?P<low>[0-9.]+)\|(?P<close>[0-9.]+)\|(?P<volume>[0-9.]+)"',
                response.text,
            ):
                raw_date = matched.group("date")
                close = float(matched.group("close"))
                points.append(
                    {
                        "date": f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:8]}",
                        "close": int(round(close)),
                    }
                )
            if points:
                return points[-count:]
        except Exception:
            continue
    return []


def attach_current_quotes(
    session: requests.Session,
    entries: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    cache: dict[str, dict[str, Any]] = {}
    enriched: list[dict[str, Any]] = []

    for entry in entries:
        enriched_entry = dict(entry)
        stock_code = clean_text(entry.get("stock_code"))
        if stock_code:
            if stock_code not in cache:
                try:
                    quote_payload = fetch_current_quote(session, stock_code)
                    quote_payload["price_history"] = fetch_price_history(session, stock_code)
                    cache[stock_code] = quote_payload
                except Exception:
                    cache[stock_code] = {
                        "live_price": None,
                        "live_price_change": None,
                        "live_price_rate": None,
                        "live_price_direction": None,
                        "live_price_as_of": None,
                        "live_price_high_52w": None,
                        "price_history": [],
                    }
            enriched_entry.update(cache[stock_code])
        enriched.append(enriched_entry)

    return enriched


def build_mobile_entry(
    report: dict[str, Any],
    summary: dict[str, Any],
    pdf_path: Path,
    json_path: Path,
) -> dict[str, Any]:
    relative_pdf = Path("..") / pdf_path
    relative_json = Path("..") / json_path
    return {
        "report_idx": str(report["report_idx"]),
        "category": report["category"],
        "published_at": report["published_at"],
        "title": report["title"],
        "author": report.get("author"),
        "source": report.get("source"),
        "company_name": summary.get("company_name") or report.get("company_name"),
        "stock_code": summary.get("stock_code") or report.get("stock_code"),
        "industry_name": summary.get("industry_name") or report.get("industry_name"),
        "market_topic": summary.get("market_topic"),
        "investment_opinion": summary.get("investment_opinion") or report.get("investment_opinion"),
        "target_price": summary.get("target_price") or report.get("target_price"),
        "summary": summary.get("summary"),
        "key_points": summary.get("key_points", []),
        "risks": summary.get("risks", []),
        "alpha_score": summary.get("alpha_score"),
        "alpha_grade": summary.get("alpha_grade"),
        "score_reason": summary.get("score_reason"),
        "positive_factors": summary.get("positive_factors", []),
        "negative_factors": summary.get("negative_factors", []),
        "revision_signal": summary.get("revision_signal"),
        "structural_growth_signal": summary.get("structural_growth_signal"),
        "valuation_risk_signal": summary.get("valuation_risk_signal"),
        "market_reaction_placeholder": summary.get(
            "market_reaction_placeholder", dict(MARKET_REACTION_TEMPLATE)
        ),
        "live_price": None,
        "live_price_change": None,
        "live_price_rate": None,
        "live_price_direction": None,
        "live_price_as_of": None,
        "live_price_high_52w": None,
        "price_history": [],
        "pdf_path": relative_pdf.as_posix(),
        "json_path": relative_json.as_posix(),
        "pdf_url": report["pdf_url"],
    }


def process_report(
    session: requests.Session,
    client: OpenAI,
    report: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    category = report["category"]
    report_idx = report["report_idx"]
    pdf_path = PDF_ROOT / category / f"{report_idx}.pdf"
    json_path = JSON_ROOT / category / f"{report_idx}.json"

    if not pdf_path.exists():
        download_pdf(session, report["pdf_url"], pdf_path)

    pdf_text = extract_pdf_text(pdf_path, MAX_PDF_PAGES)
    if not pdf_text:
        pdf_text = (
            "PDF 본문 텍스트를 추출하지 못했습니다. "
            "제목, 작성자, 제공출처, 분류, 목록 메타데이터만 기반으로 "
            "보수적으로 요약하고 추정하지 마십시오."
        )

    summary = enrich_summary(report, summarize_report(client, report, pdf_text))

    json_payload = {
        "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "metadata": report,
        "summary": summary,
        "pdf_path": str(pdf_path),
    }
    save_json(json_payload, json_path)

    excel_row = to_excel_row(report, summary, pdf_path)
    mobile_entry = build_mobile_entry(report, summary, pdf_path, json_path)
    return excel_row, mobile_entry


def render_mobile_html(
    target_date: str,
    entries: list[dict[str, Any]],
    available_dates: list[str],
) -> str:
    payload = json.dumps(entries, ensure_ascii=False)
    dates_payload = json.dumps(available_dates, ensure_ascii=False)
    template = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>한경 컨센서스 __TARGET_DATE__</title>
  <style>
    :root {{
      --bg: #f5efe4;
      --paper: #fffdf8;
      --ink: #1d2a33;
      --muted: #60717c;
      --line: #d8cdb7;
      --accent: #005f73;
      --accent-2: #ca6702;
      --chip: #e9f2f4;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Apple SD Gothic Neo", sans-serif;
      background:
        radial-gradient(circle at top right, rgba(202,103,2,0.12), transparent 30%),
        linear-gradient(180deg, #f8f1e7 0%, #f2eadb 100%);
      color: var(--ink);
    }}
    .wrap {{
      max-width: 960px;
      margin: 0 auto;
      padding: 16px 14px 40px;
    }}
    .hero {{
      background: rgba(255,255,255,0.76);
      backdrop-filter: blur(8px);
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 24px;
      padding: 18px;
      box-shadow: 0 10px 30px rgba(71, 61, 44, 0.08);
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 28px;
      line-height: 1.15;
    }}
    .sub {{
      color: var(--muted);
      font-size: 14px;
      margin-bottom: 14px;
      line-height: 1.5;
    }}
    .credit {{
      color: var(--muted);
      font-size: 12px;
      margin-top: -8px;
      margin-bottom: 14px;
      letter-spacing: 0.02em;
      font-weight: 700;
    }}
    .toolbar {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }}
    .toolbar.secondary {{
      margin-top: 10px;
    }}
    .search-panel {{
      margin-top: 14px;
      padding: 14px;
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 18px;
      background: rgba(255,255,255,0.62);
    }}
    .search-title {{
      font-size: 14px;
      font-weight: 700;
      margin-bottom: 8px;
    }}
    .search-row {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .search-row input {{
      flex: 1 1 220px;
      min-width: 0;
      border: 1px solid var(--line);
      background: var(--paper);
      color: var(--ink);
      border-radius: 12px;
      padding: 11px 12px;
      font-size: 14px;
    }}
    .search-row button {{
      border: 1px solid var(--accent);
      background: var(--accent);
      color: white;
      border-radius: 12px;
      padding: 11px 14px;
      font-size: 14px;
      cursor: pointer;
      font-weight: 600;
    }}
    .search-hint {{
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }}
    .search-results {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }}
    .search-results button {{
      border: 1px solid rgba(0,95,115,0.12);
      background: #f3f7f8;
      color: var(--accent);
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 13px;
      cursor: pointer;
      font-weight: 600;
    }}
    .toolbar.cluster {{
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }}
    .toolbar button {{
      border: 1px solid var(--line);
      background: var(--paper);
      color: var(--ink);
      border-radius: 999px;
      padding: 10px 14px;
      font-size: 14px;
      cursor: pointer;
    }}
    .toolbar button.active {{
      background: var(--accent);
      color: white;
      border-color: var(--accent);
    }}
    .date-box {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
      color: var(--muted);
      font-size: 13px;
    }}
    .date-box select {{
      border: 1px solid var(--line);
      background: var(--paper);
      color: var(--ink);
      border-radius: 12px;
      padding: 10px 12px;
      font-size: 14px;
      min-width: 148px;
    }}
    .stats {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-top: 10px;
    }}
    .stat {{
      background: var(--chip);
      border-radius: 16px;
      padding: 10px 12px;
      font-size: 13px;
    }}
    .list {{
      margin-top: 14px;
      display: grid;
      gap: 14px;
    }}
    .card {{
      background: rgba(255,255,255,0.84);
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 20px;
      padding: 16px;
      box-shadow: 0 8px 20px rgba(71, 61, 44, 0.06);
    }}
    .group-card {{
      background: rgba(255,255,255,0.92);
      border: 1px solid rgba(0,95,115,0.14);
      border-radius: 24px;
      padding: 18px;
      box-shadow: 0 10px 24px rgba(0,95,115,0.08);
    }}
    .row {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      align-items: center;
      margin-bottom: 8px;
    }}
    .badge {{
      display: inline-flex;
      align-items: center;
      border-radius: 999px;
      padding: 5px 10px;
      font-size: 12px;
      background: var(--chip);
      color: var(--accent);
      font-weight: 600;
    }}
    .title {{
      margin: 0 0 8px;
      font-size: 18px;
      line-height: 1.35;
    }}
    .meta {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
      margin-bottom: 10px;
    }}
    .summary {{
      margin: 0 0 12px;
      line-height: 1.55;
      font-size: 14px;
    }}
    .points {{
      margin: 0;
      padding-left: 18px;
      color: var(--ink);
      font-size: 14px;
      line-height: 1.55;
    }}
    .links {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-top: 12px;
    }}
    .links a {{
      color: var(--accent-2);
      text-decoration: none;
      font-weight: 600;
      font-size: 14px;
    }}
    .links button {{
      border: 0;
      background: transparent;
      color: var(--accent-2);
      text-decoration: none;
      font-weight: 600;
      font-size: 14px;
      padding: 0;
      cursor: pointer;
      font-family: inherit;
    }}
    .group-head {{
      display: grid;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 14px;
    }}
    .group-title-row {{
      display: flex;
      flex-wrap: wrap;
      align-items: flex-end;
      gap: 8px 14px;
    }}
    .group-title {{
      margin: 0;
      font-size: 22px;
      line-height: 1.2;
    }}
    .group-sub {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
      margin-top: 6px;
    }}
    .price-box {{
      display: flex;
      flex-wrap: wrap;
      align-items: baseline;
      gap: 4px 10px;
      min-width: 0;
    }}
    .price-label {{
      color: var(--muted);
      font-size: 12px;
      margin: 0;
    }}
    .price-main {{
      font-size: 18px;
      font-weight: 700;
      line-height: 1.1;
    }}
    .price-change {{
      font-size: 13px;
      color: var(--muted);
    }}
    .price-change.up {{
      color: #d11f1f !important;
      font-weight: 700;
    }}
    .price-change.down {{
      color: #1f57d1 !important;
      font-weight: 700;
    }}
    .price-extra {{
      width: 100%;
      font-size: 12px;
      color: var(--muted);
    }}
    .price-chart {{
      width: min(220px, 100%);
      height: 52px;
      margin-top: 6px;
      padding: 6px 8px;
      border-radius: 14px;
      background: linear-gradient(180deg, #fff8ed 0%, #fff2e1 100%);
      border: 1px solid rgba(202,103,2,0.18);
    }}
    .price-chart svg {{
      display: block;
      width: 100%;
      height: 100%;
    }}
    .price-caption {{
      width: 100%;
      margin-top: 2px;
      font-size: 11px;
      color: var(--muted);
    }}
    .group-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 10px 0 14px;
    }}
    .score-box {{
      min-width: 160px;
      background: linear-gradient(180deg, #edf7f1 0%, #e3f3e9 100%);
      border: 1px solid rgba(0,95,115,0.18);
      border-radius: 18px;
      padding: 12px 14px;
    }}
    .score-grade {{
      font-size: 13px;
      color: var(--muted);
      margin-bottom: 4px;
    }}
    .score-main {{
      font-size: 28px;
      font-weight: 700;
      line-height: 1.1;
    }}
    .score-sub {{
      margin-top: 6px;
      font-size: 13px;
      color: var(--muted);
    }}
    .pill {{
      background: #f3f7f8;
      color: var(--accent);
      border-radius: 999px;
      padding: 7px 10px;
      font-size: 12px;
      font-weight: 600;
    }}
    .report-stack {{
      display: grid;
      gap: 10px;
    }}
    details.report-detail {{
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 16px;
      padding: 12px 14px;
      background: rgba(250,247,241,0.96);
    }}
    details.report-detail summary {{
      cursor: pointer;
      list-style: none;
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: flex-start;
      font-weight: 600;
    }}
    details.report-detail summary::-webkit-details-marker {{
      display: none;
    }}
    .report-brief {{
      color: var(--muted);
      font-size: 13px;
      margin-top: 2px;
      font-weight: 400;
      line-height: 1.45;
    }}
    .report-body {{
      margin-top: 12px;
    }}
    .section-title {{
      margin: 2px 0;
      font-size: 16px;
      font-weight: 700;
    }}
    .history-section {{
      margin-top: 16px;
      display: grid;
      gap: 12px;
    }}
    .history-summary {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
      margin: 4px 0 0;
    }}
    .history-empty {{
      background: rgba(255,255,255,0.84);
      border: 1px dashed rgba(216,205,183,0.9);
      border-radius: 18px;
      padding: 16px;
      color: var(--muted);
      font-size: 13px;
    }}
    .factor-grid {{
      display: grid;
      gap: 10px;
      margin-top: 12px;
    }}
    .factor-box {{
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 14px;
      padding: 10px 12px;
      background: rgba(255,255,255,0.7);
    }}
    .factor-title {{
      margin: 0 0 6px;
      font-size: 13px;
      font-weight: 700;
      color: var(--muted);
    }}
    .factor-text {{
      margin: 0;
      font-size: 13px;
      line-height: 1.5;
    }}
    .history-group-card {{
      padding-bottom: 14px;
    }}
    .history-report-card {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      border: 1px solid rgba(216,205,183,0.9);
      border-radius: 18px;
      padding: 14px 16px;
      background: rgba(255,255,255,0.82);
    }}
    .history-report-main {{
      min-width: 0;
      flex: 1;
    }}
    .history-report-title {{
      margin: 0;
      font-size: 18px;
      line-height: 1.35;
    }}
    .history-title-row {{
      display: flex;
      gap: 8px;
      align-items: center;
      flex-wrap: wrap;
    }}
    .history-report-badges {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: flex-end;
      min-width: 112px;
    }}
    @media (max-width: 640px) {{
      .group-title {{
        font-size: 20px;
      }}
      .history-report-card {{
        display: grid;
      }}
      .history-report-badges {{
        justify-content: flex-start;
        min-width: 0;
      }}
      .price-chart {{
        width: 100%;
      }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>한경 컨센서스 데일리</h1>
      <div class="sub" id="subline">대상 일자: __TARGET_DATE__</div>
      <div class="credit">Built by Kujah with Codex</div>
      <div class="toolbar cluster">
        <div class="toolbar" id="modes"></div>
        <div class="date-box">
          <span>브리프 날짜</span>
          <select id="date-select"></select>
        </div>
      </div>
      <div class="toolbar secondary" id="filters"></div>
      <div class="search-panel">
        <div class="search-title">종목 이력 검색</div>
        <div class="search-row">
          <input id="stock-search" type="search" placeholder="종목명 또는 6자리 종목코드" autocomplete="off">
          <button id="search-button" type="button">검색</button>
        </div>
        <div class="search-hint">저장된 날짜별 아카이브를 기준으로 분석/요약 이력을 조회합니다.</div>
        <div class="search-results" id="search-results"></div>
      </div>
      <div class="stats" id="stats"></div>
    </section>
    <section class="history-section" id="history-section"></section>
    <section class="list" id="list"></section>
  </div>
  <script>
    const initialReports = __PAYLOAD__;
    const availableDates = __DATES_PAYLOAD__;
    const categories = ["전체", "기업", "산업", "시장"];
    const modes = [
      { key: "grouped", label: "종목 묶음" },
      { key: "reports", label: "전체 리포트" },
    ];
    let current = "전체";
    let currentMode = "grouped";
    let currentDate = "__TARGET_DATE__";
    let reports = initialReports;
    const reportCache = new Map([[currentDate, initialReports]]);
    let historyLoaded = false;

    function safe(value) {
      return value ? String(value) : "";
    }

    function parsePriceValue(value) {
      const digits = String(value || "").replace(/[^0-9.-]/g, "");
      if (!digits) return null;
      const parsed = Number(digits);
      return Number.isFinite(parsed) ? parsed : null;
    }

    function formatWon(value) {
      if (!Number.isFinite(value)) return "";
      return `${Math.round(value).toLocaleString("ko-KR")}원`;
    }

    function formatPercent(value) {
      if (!Number.isFinite(value)) return "";
      return `${value.toFixed(1)}%`;
    }

    function getUpsideData(item) {
      const livePrice = parsePriceValue(item.live_price);
      const targetPrice = parsePriceValue(item.target_price);
      if (!Number.isFinite(livePrice) || !Number.isFinite(targetPrice) || livePrice <= 0) {
        return null;
      }
      const diff = targetPrice - livePrice;
      return {
        diff,
        pct: (diff / livePrice) * 100,
      };
    }

    function upsideLabel(item) {
      const upside = getUpsideData(item);
      if (!upside) return "";
      return `현재가 기준 상승여력 ${formatWon(upside.diff)} (${formatPercent(upside.pct)})`;
    }

    function filteredReports() {
      return current === "전체"
        ? reports
        : reports.filter((item) => item.category === current);
    }

    function countByCategory(items) {
      const counts = {};
      for (const item of items) counts[item.category] = (counts[item.category] || 0) + 1;
      return counts;
    }

    function groupedData(items) {
      const groups = new Map();
      const industry = [];
      const market = [];

      for (const item of items) {
        if (item.category === "기업" && item.stock_code) {
          const key = item.stock_code;
          if (!groups.has(key)) {
            groups.set(key, {
              stock_code: item.stock_code,
              company_name: item.company_name || item.title,
              live_price: item.live_price,
              live_price_change: item.live_price_change,
              live_price_rate: item.live_price_rate,
              live_price_as_of: item.live_price_as_of,
              live_price_high_52w: item.live_price_high_52w,
              price_history: item.price_history || [],
              best_alpha_score: item.alpha_score,
              best_alpha_grade: item.alpha_grade,
              reports: [],
            });
          }
          const group = groups.get(key);
          if ((!group.price_history || !group.price_history.length) && item.price_history?.length) {
            group.price_history = item.price_history;
          }
          if ((item.alpha_score || -1) > (group.best_alpha_score || -1)) {
            group.best_alpha_score = item.alpha_score;
            group.best_alpha_grade = item.alpha_grade;
          }
          group.reports.push(item);
        } else if (item.category === "산업") {
          industry.push(item);
        } else {
          market.push(item);
        }
      }

      const groupedCompanies = Array.from(groups.values())
        .map((group) => ({
          ...group,
          reports: group.reports.sort((a, b) => {
            const scoreDiff = (b.alpha_score || -1) - (a.alpha_score || -1);
            if (scoreDiff !== 0) return scoreDiff;
            const sourceA = a.source || "";
            const sourceB = b.source || "";
            return sourceA.localeCompare(sourceB, "ko");
          }),
        }))
        .sort((a, b) => {
          const scoreDiff = (b.best_alpha_score || -1) - (a.best_alpha_score || -1);
          if (scoreDiff !== 0) return scoreDiff;
          const countDiff = b.reports.length - a.reports.length;
          if (countDiff !== 0) return countDiff;
          return (a.company_name || "").localeCompare(b.company_name || "", "ko");
        });

      return { groupedCompanies, industry, market };
    }

    function renderModes() {
      const root = document.getElementById("modes");
      root.innerHTML = "";
      for (const mode of modes) {
        const button = document.createElement("button");
        button.textContent = mode.label;
        if (mode.key === currentMode) button.classList.add("active");
        button.onclick = () => {
          currentMode = mode.key;
          render();
        };
        root.appendChild(button);
      }
    }

    function renderDateOptions() {
      const root = document.getElementById("date-select");
      root.innerHTML = availableDates
        .map((date) => `<option value="${safe(date)}"${date === currentDate ? " selected" : ""}>${safe(date)}</option>`)
        .join("");
      root.onchange = async (event) => {
        await switchDate(event.target.value);
      };
    }

    function renderFilters() {
      const root = document.getElementById("filters");
      root.innerHTML = "";
      for (const category of categories) {
        const button = document.createElement("button");
        button.textContent = category;
        if (category === current) button.classList.add("active");
        button.onclick = () => {
          current = category;
          render();
        };
        root.appendChild(button);
      }
    }

    function renderStats(items) {
      const root = document.getElementById("stats");
      root.innerHTML = "";
      document.getElementById("subline").textContent = "";
      document.getElementById("subline").style.display = "none";
    }

    function formatQuoteAsOf(value) {
      const text = safe(value);
      const matched = text.match(/(\\d{4}\\.\\d{2}\\.\\d{2})/);
      if (matched) return `${matched[1].slice(2)} 장마감 기준`;
      return text
        .replace(/^현재가\\s*/, "")
        .replace(/\\s*기준(?:\\(KRX 장마감\\)|\\(장마감\\)|\\(KRX장마감\\)|\\(KRX 마감\\))?/g, "")
        .trim();
    }

    function priceDirectionClass(group) {
      const explicit = safe(group.live_price_direction).toLowerCase();
      if (explicit === "up" || explicit === "rise") return "up";
      if (explicit === "down" || explicit === "fall") return "down";

      const changeAmount = parsePriceValue(group.live_price_change);
      const changeRate = parsePriceValue(group.live_price_rate);
      if (Number.isFinite(changeAmount) && changeAmount < 0) return "down";
      if (Number.isFinite(changeRate) && changeRate < 0) return "down";
      if (Number.isFinite(changeAmount) && changeAmount > 0) return "up";
      if (Number.isFinite(changeRate) && changeRate > 0) return "up";

      const text = `${safe(group.live_price_change)} ${safe(group.live_price_rate)}`;
      if (text.includes("-")) return "down";
      return "";
    }

    function renderSparkline(history) {
      const values = (history || [])
        .map((item) => parsePriceValue(item.close))
        .filter((value) => Number.isFinite(value));
      if (values.length < 2) return "";

      const width = 220;
      const height = 40;
      const padding = 2;
      const minValue = Math.min(...values);
      const maxValue = Math.max(...values);
      const range = Math.max(maxValue - minValue, 1);
      const points = values.map((value, index) => {
        const x = padding + ((width - padding * 2) * index) / (values.length - 1);
        const y = height - padding - ((value - minValue) / range) * (height - padding * 2);
        return `${x.toFixed(1)},${y.toFixed(1)}`;
      });
      const lastPoint = points[points.length - 1].split(",");
      const trendUp = values[values.length - 1] >= values[0];
      const stroke = trendUp ? "#0d7a5f" : "#ca6702";
      return `
        <div class="price-chart" aria-label="최근 3개월 주가 추이">
          <svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" role="img" aria-hidden="true">
            <polyline fill="none" stroke="${stroke}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${points.join(" ")}"></polyline>
            <circle cx="${lastPoint[0]}" cy="${lastPoint[1]}" r="2.8" fill="${stroke}"></circle>
          </svg>
        </div>
        <div class="price-caption">[3개월 주가 추이]</div>
      `;
    }

    function renderPriceBox(group) {
      if (!group.live_price) {
        return `
          <div class="price-box">
            <div class="price-label">현재가</div>
            <div class="price-main">조회 실패</div>
          </div>
        `;
      }

      const change = [safe(group.live_price_change), safe(group.live_price_rate)].filter(Boolean).join(" / ");
      const directionClass = priceDirectionClass(group);
      const high52w = safe(group.live_price_high_52w);
      return `
        <div class="price-box">
          <div class="price-label">${formatQuoteAsOf(group.live_price_as_of)}</div>
          <div class="price-main">${safe(group.live_price)}원</div>
          <div class="price-change ${directionClass}">${change}</div>
          ${high52w ? `<div class="price-extra">52주 최고가 ${high52w}원</div>` : ""}
          ${renderSparkline(group.price_history)}
        </div>
      `;
    }

    function renderAlphaBox(item) {
      if (item.alpha_score === null || item.alpha_score === undefined) return "";
      return `
        <div class="score-box">
          <div class="score-grade">Alpha ${safe(item.alpha_grade || "-")}</div>
          <div class="score-main">${safe(item.alpha_score)}</div>
          <div class="score-sub">${safe(item.score_reason || "")}</div>
        </div>
      `;
    }

    function stockWebUrl(stockCode) {
      if (!stockCode) return "";
      return `https://finance.naver.com/item/main.naver?code=${encodeURIComponent(stockCode)}`;
    }

    async function copyStockCode(stockCode) {
      if (!stockCode) return;
      try {
        await navigator.clipboard.writeText(String(stockCode));
      } catch (error) {
        const textarea = document.createElement("textarea");
        textarea.value = String(stockCode);
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.focus();
        textarea.select();
        document.execCommand("copy");
        document.body.removeChild(textarea);
      }
      alert(`종목코드 복사: ${stockCode}`);
    }

    async function fetchReportsForDate(date) {
      if (!date) return [];
      if (reportCache.has(date)) return reportCache.get(date) || [];
      try {
        const response = await fetch(`./archive/${date}.json`, { cache: "no-store" });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const payload = await response.json();
        const nextReports = payload.reports || [];
        reportCache.set(date, nextReports);
        return nextReports;
      } catch (error) {
        return null;
      }
    }

    async function ensureHistoryLoaded() {
      if (historyLoaded) return true;
      for (const date of availableDates) {
        const loaded = await fetchReportsForDate(date);
        if (loaded === null) return false;
      }
      historyLoaded = true;
      return true;
    }

    function allCachedReports() {
      return Array.from(reportCache.values()).flat();
    }

    function stockCatalog() {
      const grouped = new Map();
      for (const item of allCachedReports()) {
        if (item.category !== "기업" || !item.stock_code) continue;
        const key = item.stock_code;
        if (!grouped.has(key)) {
          grouped.set(key, {
            stock_code: item.stock_code,
            company_name: item.company_name || item.title || item.stock_code,
            report_count: 0,
            dates: new Set(),
            latest_date: item.published_at || "",
          });
        }
        const stock = grouped.get(key);
        stock.report_count += 1;
        if (item.company_name) stock.company_name = item.company_name;
        if (item.published_at) stock.dates.add(item.published_at);
        if ((item.published_at || "") > stock.latest_date) stock.latest_date = item.published_at || "";
      }
      return Array.from(grouped.values());
    }

    function renderSearchResults(htmlText) {
      document.getElementById("search-results").innerHTML = htmlText;
    }

    function renderHistoryEmpty(message) {
      const root = document.getElementById("history-section");
      root.innerHTML = `<div class="history-empty">${safe(message)}</div>`;
    }

    function matchingStocks(query) {
      const keyword = String(query || "").trim().toLowerCase();
      if (!keyword) return [];
      return stockCatalog()
        .map((item) => {
          const name = String(item.company_name || "").toLowerCase();
          const code = String(item.stock_code || "").toLowerCase();
          let score = -1;
          if (code === keyword) score = 100;
          else if (name === keyword) score = 95;
          else if (code.startsWith(keyword)) score = 80;
          else if (name.startsWith(keyword)) score = 70;
          else if (code.includes(keyword)) score = 60;
          else if (name.includes(keyword)) score = 50;
          return { ...item, score };
        })
        .filter((item) => item.score >= 0)
        .sort((a, b) => {
          const scoreDiff = b.score - a.score;
          if (scoreDiff !== 0) return scoreDiff;
          const dateDiff = (b.latest_date || "").localeCompare(a.latest_date || "");
          if (dateDiff !== 0) return dateDiff;
          return (a.company_name || "").localeCompare(b.company_name || "", "ko");
        });
    }

    function renderHistoryReportCard(item) {
      const meta = [
        safe(item.source),
        safe(item.author),
        [safe(item.investment_opinion), safe(item.target_price)].filter(Boolean).join(" "),
        upsideLabel(item),
      ].filter(Boolean).join(" 쨌 ");
      return `
        <article class="history-report-card">
          <div class="history-report-main">
            <h3 class="history-report-title">${safe(item.title)}</h3>
            <div class="report-brief">${meta}</div>
          </div>
          <div class="history-report-badges">
            ${item.alpha_score !== null && item.alpha_score !== undefined ? `<span class="badge">${safe(item.alpha_grade || "-")} ${safe(item.alpha_score)}</span>` : ""}
            <span class="badge">#${safe(item.report_idx)}</span>
          </div>
        </article>
      `;
    }

    function renderHistoryForStock(stockCode) {
      renderHistoryForStockCompact(stockCode);
    }

    function renderHistoryReportCardCompact(item) {
      const meta = [
        safe(item.source),
        safe(item.author),
        [safe(item.investment_opinion), safe(item.target_price)].filter(Boolean).join(" "),
        upsideLabel(item),
      ].filter(Boolean).join(" / ");
      return `
        <article class="history-report-card">
          <div class="history-report-main">
            <div class="history-title-row">
              <h3 class="history-report-title">${safe(item.title)}</h3>
              ${item.published_at ? `<span class="badge">${safe(item.published_at)}</span>` : ""}
            </div>
            <div class="report-brief">${meta}</div>
            <p class="summary">${safe(item.summary)}</p>
            <div class="links">
              <a href="${safe(item.pdf_url)}" target="_blank" rel="noreferrer">원문</a>
            </div>
          </div>
          <div class="history-report-badges">
            ${item.alpha_score !== null && item.alpha_score !== undefined ? `<span class="badge">${safe(item.alpha_grade || "-")} ${safe(item.alpha_score)}</span>` : ""}
            <span class="badge">#${safe(item.report_idx)}</span>
          </div>
        </article>
      `;
    }

    function renderHistoryForStockCompact(stockCode) {
      const entries = allCachedReports()
        .filter((item) => item.stock_code === stockCode)
        .sort((a, b) => {
          const dateDiff = (b.published_at || "").localeCompare(a.published_at || "");
          if (dateDiff !== 0) return dateDiff;
          return String(b.report_idx || "").localeCompare(String(a.report_idx || ""));
        });
      const root = document.getElementById("history-section");
      if (!entries.length) {
        renderHistoryEmpty("?대떦 醫낅ぉ????λ맂 ?대젰???놁뒿?덈떎.");
        return;
      }

      const companyName = entries.find((item) => item.company_name)?.company_name || stockCode;
      const uniqueDates = new Set(entries.map((item) => item.published_at).filter(Boolean));
      const bestEntry = entries.find((item) => item.alpha_score !== null && item.alpha_score !== undefined);
      const cards = entries.map((item) => renderHistoryReportCardCompact(item)).join("");

      root.innerHTML = `
        <article class="group-card history-group-card">
          <div class="group-head">
            <div>
              <div class="group-title-row">
                <h2 class="group-title">${safe(companyName)}</h2>
              </div>
            </div>
          </div>
          <div class="group-meta">
            <span class="pill"><button type="button" data-copy-stock="${safe(stockCode)}">종목코드 복사</button></span>
            <span class="pill"><a href="${stockWebUrl(stockCode)}" target="_blank" rel="noreferrer">네이버금융</a></span>
          </div>
          <div class="report-stack">${cards}</div>
        </article>
      `;
    }

    async function runStockSearch() {
      const input = document.getElementById("stock-search");
      const query = String(input.value || "").trim();
      if (!query) {
        renderSearchResults("");
        renderHistoryEmpty("종목명 또는 6자리 종목코드로 검색하면 저장된 분석 이력을 볼 수 있습니다.");
        return;
      }

      renderSearchResults('<span class="badge">아카이브 확인 중...</span>');
      const loaded = await ensureHistoryLoaded();
      if (!loaded) {
        renderSearchResults("");
        renderHistoryEmpty("아카이브를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
        return;
      }

      const matches = matchingStocks(query).slice(0, 12);
      if (!matches.length) {
        renderSearchResults("");
        renderHistoryEmpty(`"${query}"와 일치하는 종목 이력이 없습니다.`);
        return;
      }

      renderSearchResults(
        matches
          .map(
            (item) => `
              <button type="button" data-stock-history="${safe(item.stock_code)}">
                ${safe(item.company_name)} (${safe(item.stock_code)}) · ${item.report_count}건
              </button>
            `
          )
          .join("")
      );
      renderHistoryForStockCompact(matches[0].stock_code);
    }

    function renderFactorSections(item) {
      if (item.category !== "기업") return "";
      const sections = [];
      if (item.revision_signal) {
        sections.push(`<div class="factor-box"><div class="factor-title">Revision Signal</div><p class="factor-text">${safe(item.revision_signal)}</p></div>`);
      }
      if (item.structural_growth_signal) {
        sections.push(`<div class="factor-box"><div class="factor-title">Structural Growth</div><p class="factor-text">${safe(item.structural_growth_signal)}</p></div>`);
      }
      if (item.valuation_risk_signal) {
        sections.push(`<div class="factor-box"><div class="factor-title">Valuation / Risk</div><p class="factor-text">${safe(item.valuation_risk_signal)}</p></div>`);
      }
      const positiveText = (item.positive_factors || []).map((value) => `+ ${safe(value)}`).join("<br>");
      const negativeText = (item.negative_factors || []).map((value) => `- ${safe(value)}`).join("<br>");
      if (positiveText) {
        sections.push(`<div class="factor-box"><div class="factor-title">Positive Factors</div><p class="factor-text">${positiveText}</p></div>`);
      }
      if (negativeText) {
        sections.push(`<div class="factor-box"><div class="factor-title">Negative Factors</div><p class="factor-text">${negativeText}</p></div>`);
      }
      if (item.market_reaction_placeholder) {
        sections.push(`<div class="factor-box"><div class="factor-title">Market Reaction</div><p class="factor-text">D+1 / D+5 / D+20 추후 연결 예정</p></div>`);
      }
      return sections.length ? `<div class="factor-grid">${sections.join("")}</div>` : "";
    }

    function renderGrouped(items) {
      const root = document.getElementById("list");
      const { groupedCompanies, industry, market } = groupedData(items);

      const groupCards = groupedCompanies.map((group) => {
        const reportsHtml = group.reports.map((report) => {
          const pointItems = (report.key_points || []).slice(0, 4).map((point) => `<li>${safe(point)}</li>`).join("");
          return `
            <details class="report-detail">
              <summary>
                <div>
                  <div>${safe(report.title)}</div>
                  <div class="report-brief">${[
                    safe(report.source),
                    safe(report.author),
                    [safe(report.investment_opinion), safe(report.target_price)].filter(Boolean).join(" "),
                    upsideLabel(report),
                  ].filter(Boolean).join(" · ")}</div>
                </div>
                <div>
                  ${report.alpha_score !== null && report.alpha_score !== undefined ? `<div class="badge">${safe(report.alpha_grade || "-")} ${safe(report.alpha_score)}</div>` : ""}
                  <div class="badge">#${safe(report.report_idx)}</div>
                </div>
              </summary>
              <div class="report-body">
                <p class="summary">${safe(report.summary)}</p>
                ${pointItems ? `<ul class="points">${pointItems}</ul>` : ""}
                ${renderFactorSections(report)}
                <div class="links">
                  <a href="${safe(report.pdf_url)}" target="_blank" rel="noreferrer">원문 링크</a>
                </div>
              </div>
            </details>
          `;
        }).join("");

        return `
          <article class="group-card">
            <div class="group-head">
              <div>
                <div class="group-title-row">
                  <h2 class="group-title">${safe(group.company_name)}</h2>
                  ${renderPriceBox(group)}
                </div>
              </div>
            </div>
            <div class="group-meta">
              ${group.stock_code ? `<span class="pill"><button type="button" data-copy-stock="${safe(group.stock_code)}">종목코드 복사</button></span>` : ""}
              ${group.stock_code ? `<span class="pill"><a href="${stockWebUrl(group.stock_code)}" target="_blank" rel="noreferrer">네이버금융</a></span>` : ""}
            </div>
            <div class="report-stack">${reportsHtml}</div>
          </article>
        `;
      }).join("");

      const renderStandaloneSection = (title, entries) => entries.length
        ? `
          <div class="section-title">${title}</div>
          ${entries.map((item) => `
            <article class="card">
              <div class="row">
                <span class="badge">${safe(item.category)}</span>
                <span class="badge">${safe(item.published_at)}</span>
                <span class="badge">#${safe(item.report_idx)}</span>
              </div>
              <h2 class="title">${safe(item.title)}</h2>
              <div class="meta">작성자 ${safe(item.author)} · 출처 ${safe(item.source)}${item.market_topic ? ` · 주제 ${safe(item.market_topic)}` : ""}${item.industry_name ? ` · 산업 ${safe(item.industry_name)}` : ""}</div>
              <p class="summary">${safe(item.summary)}</p>
              ${renderFactorSections(item)}
              <div class="links">
                <a href="${safe(item.pdf_url)}" target="_blank" rel="noreferrer">원문 링크</a>
              </div>
            </article>
          `).join("")}
        `
        : "";

      root.innerHTML = groupCards
        + renderStandaloneSection("산업 리포트", industry)
        + renderStandaloneSection("시장 리포트", market);
    }

    function renderList(items) {
      const root = document.getElementById("list");
      root.innerHTML = items.map((item) => {
        const meta = [
          item.author ? `작성자 ${item.author}` : "",
          item.source ? `출처 ${item.source}` : "",
          item.company_name ? `회사 ${item.company_name}` : "",
          item.stock_code ? `코드 ${item.stock_code}` : "",
          item.industry_name ? `산업 ${item.industry_name}` : "",
          item.market_topic ? `주제 ${item.market_topic}` : "",
          item.investment_opinion ? `의견 ${item.investment_opinion}` : "",
          item.target_price ? `목표가 ${item.target_price}` : "",
          upsideLabel(item),
          item.alpha_score !== null && item.alpha_score !== undefined ? `Alpha ${item.alpha_grade || "-"} ${item.alpha_score}` : "",
          item.live_price ? `현재가 ${item.live_price}원` : "",
        ].filter(Boolean).join(" · ");
        const points = (item.key_points || []).slice(0, 4).map((point) => `<li>${safe(point)}</li>`).join("");
        return `
          <article class="card">
            <div class="row">
              <span class="badge">${safe(item.category)}</span>
              <span class="badge">${safe(item.published_at)}</span>
              <span class="badge">#${safe(item.report_idx)}</span>
            </div>
            <h2 class="title">${safe(item.title)}</h2>
            <div class="meta">${meta}</div>
            ${renderAlphaBox(item)}
            <p class="summary">${safe(item.summary)}</p>
            ${points ? `<ul class="points">${points}</ul>` : ""}
            ${renderFactorSections(item)}
            <div class="links">
              <a href="${safe(item.pdf_url)}" target="_blank" rel="noreferrer">원문 링크</a>
            </div>
          </article>
        `;
      }).join("");
    }

    function render() {
      renderModes();
      renderDateOptions();
      renderFilters();
      const items = filteredReports();
      renderStats(items);
      if (currentMode === "grouped") {
        renderGrouped(items);
      } else {
        renderList(items);
      }
    }

    async function switchDate(nextDate) {
      if (!nextDate || nextDate === currentDate) return;
      const nextReports = await fetchReportsForDate(nextDate);
      if (nextReports === null) {
        alert(`브리프를 불러오지 못했습니다: ${nextDate}`);
        renderDateOptions();
        return;
      }
      currentDate = nextDate;
      reports = nextReports;
      render();
    }

    document.addEventListener("click", async (event) => {
      const trigger = event.target.closest("[data-copy-stock]");
      if (!trigger) return;
      event.preventDefault();
      await copyStockCode(trigger.getAttribute("data-copy-stock"));
    });

    document.addEventListener("click", (event) => {
      const trigger = event.target.closest("[data-stock-history]");
      if (!trigger) return;
      event.preventDefault();
      renderHistoryForStockCompact(trigger.getAttribute("data-stock-history"));
    });

    document.getElementById("search-button").addEventListener("click", () => {
      runStockSearch();
    });

    document.getElementById("stock-search").addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      runStockSearch();
    });

    render();
    renderHistoryEmpty("종목명 또는 6자리 종목코드로 검색하면 저장된 분석 이력을 볼 수 있습니다.");
  </script>
</body>
</html>
"""
    rendered = (
        template.replace("__TARGET_DATE__", html.escape(target_date))
        .replace("__PAYLOAD__", payload)
        .replace("__DATES_PAYLOAD__", dates_payload)
    )
    return rendered.replace("{{", "{").replace("}}", "}")


def save_mobile_outputs(target_date: str, entries: list[dict[str, Any]]) -> None:
    enriched_entries = attach_current_quotes(build_session(), entries)
    archive_root = MOBILE_ROOT / "archive"
    archive_root.mkdir(parents=True, exist_ok=True)
    payload = {
        "target_date": target_date,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count": len(enriched_entries),
        "reports": enriched_entries,
    }
    save_json(payload, archive_root / f"{target_date}.json")
    save_json(payload, MOBILE_ROOT / "latest.json")
    save_json(payload, MOBILE_ROOT / f"{target_date}.json")
    available_dates = sorted(
        {
            path.stem
            for path in archive_root.glob("*.json")
            if path.is_file() and path.stem != "index"
        },
        reverse=True,
    )
    save_json(
        {"latest": target_date, "available_dates": available_dates},
        archive_root / "index.json",
    )
    (MOBILE_ROOT / "index.html").write_text(
        render_mobile_html(target_date, enriched_entries, available_dates),
        encoding="utf-8",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Check whether the target date has any reports not yet stored locally.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    configure_stdout()
    ensure_output_dirs()

    session = build_session()
    target_date = get_target_date(session)

    if args.check_only:
        print(json.dumps(summarize_pending_reports(session, target_date), ensure_ascii=False, indent=2))
        return

    if not os.getenv("OPENAI_API_KEY"):
        raise EnvironmentError("OPENAI_API_KEY is not set.")

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    print(f"대상 일자: {target_date}")

    total_new_rows = 0
    mobile_entries: list[dict[str, Any]] = []

    for report_code, category in REPORT_TYPES.items():
        reports = get_report_list_for_date(session, report_code, target_date)
        print(f"[{category}] {target_date} 수집 건수: {len(reports)}")

        excel_path = category_excel_path(category)
        existing_ids = load_existing_ids(excel_path)
        new_rows: list[dict[str, Any]] = []

        for index, report in enumerate(reports, start=1):
            report_idx = str(report["report_idx"])
            json_path = JSON_ROOT / category / f"{report_idx}.json"
            pdf_path = PDF_ROOT / category / f"{report_idx}.pdf"

            if report_idx in existing_ids and json_path.exists() and pdf_path.exists():
                try:
                    payload = json.loads(json_path.read_text(encoding="utf-8"))
                    stored_summary = payload.get("summary", {})
                    if summary_has_scoring_fields(stored_summary):
                        normalized_summary = enrich_summary(report, stored_summary)
                        if normalized_summary != stored_summary:
                            payload["summary"] = normalized_summary
                            save_json(payload, json_path)
                        mobile_entries.append(
                            build_mobile_entry(report, normalized_summary, pdf_path, json_path)
                        )
                        print(f"[{category}][{index}/{len(reports)}] 중복 스킵: {report_idx}")
                        continue
                    print(f"[{category}][{index}/{len(reports)}] 재점수화 시작: {report_idx}")
                except Exception:
                    print(f"[{category}][{index}/{len(reports)}] 기존 JSON 재사용 실패, 재처리: {report_idx}")

            try:
                print(f"[{category}][{index}/{len(reports)}] 처리 시작: {report_idx}")
                row, mobile_entry = process_report(session, client, report)
                new_rows.append(row)
                mobile_entries.append(mobile_entry)
                total_new_rows += 1
                print(f"[{category}][{index}/{len(reports)}] 완료: {report_idx}")
                time.sleep(REQUEST_SLEEP_SECONDS)
            except Exception as exc:
                print(f"[{category}][{index}/{len(reports)}] 실패: {report_idx} / {exc}")

        if new_rows:
            final_count = append_rows_to_excel(new_rows, excel_path)
            print(f"[{category}] 엑셀 저장 완료: {excel_path} / 누적 {final_count}건")
        else:
            print(f"[{category}] 신규 저장 건수 없음")

    mobile_entries.sort(
        key=lambda item: (item["published_at"], item["category"], item["report_idx"]),
        reverse=True,
    )
    save_mobile_outputs(target_date, mobile_entries)

    print(f"전체 신규 저장 건수: {total_new_rows}")
    print(f"모바일 요약 생성 완료: {MOBILE_ROOT / 'index.html'}")


if __name__ == "__main__":
    main()
