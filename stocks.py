import html
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
    ],
    "additionalProperties": False,
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


def build_prompt(report: dict[str, Any], pdf_text: str) -> str:
    if report["category"] == "기업":
        focus = (
            "기업 리포트다. 회사명, 종목코드, 투자의견, 목표주가, 실적 모멘텀, "
            "핵심 포인트, 밸류에이션, 리스크를 정리하라."
        )
    elif report["category"] == "산업":
        focus = (
            "산업 리포트다. 개별 종목 목표주가를 억지로 채우지 말고, 산업명/업황/수급/"
            "정책/체인 이슈 중심으로 요약하라."
        )
    else:
        focus = (
            "시장 리포트다. 개별 종목 의견을 억지로 만들지 말고, 시장 주제/매크로/수급/"
            "전략 관점으로 요약하라."
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

    return {
        "live_price": clean_text(price_node.get_text(" ", strip=True)),
        "live_price_change": change_amount,
        "live_price_rate": change_rate,
        "live_price_direction": direction,
        "live_price_as_of": quote_as_of,
    }


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
                    cache[stock_code] = fetch_current_quote(session, stock_code)
                except Exception:
                    cache[stock_code] = {
                        "live_price": None,
                        "live_price_change": None,
                        "live_price_rate": None,
                        "live_price_direction": None,
                        "live_price_as_of": None,
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
        "live_price": None,
        "live_price_change": None,
        "live_price_rate": None,
        "live_price_direction": None,
        "live_price_as_of": None,
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

    summary = summarize_report(client, report, pdf_text)
    summary["report_type"] = category

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
    .toolbar {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 8px;
    }}
    .toolbar.secondary {{
      margin-top: 10px;
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
    .group-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      flex-wrap: wrap;
      margin-bottom: 14px;
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
      min-width: 220px;
      background: linear-gradient(180deg, #fff8ed 0%, #fff2e1 100%);
      border: 1px solid rgba(202,103,2,0.18);
      border-radius: 18px;
      padding: 12px 14px;
    }}
    .price-label {{
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 4px;
    }}
    .price-main {{
      font-size: 24px;
      font-weight: 700;
      line-height: 1.1;
    }}
    .price-change {{
      margin-top: 6px;
      font-size: 13px;
      color: var(--muted);
    }}
    .group-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 10px 0 14px;
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
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>한경 컨센서스 일일 브리프</h1>
      <div class="sub" id="subline">대상 일자: __TARGET_DATE__ · 기업은 종목별로 묶어서 보고, 현재가는 py stocks.py 실행 시점 최신값을 함께 표시함</div>
      <div class="toolbar cluster">
        <div class="toolbar" id="modes"></div>
        <div class="date-box">
          <span>브리프 날짜</span>
          <select id="date-select"></select>
        </div>
      </div>
      <div class="toolbar secondary" id="filters"></div>
      <div class="stats" id="stats"></div>
    </section>
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

    function safe(value) {
      return value ? String(value) : "";
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
      const standalone = [];

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
              reports: [],
            });
          }
          groups.get(key).reports.push(item);
        } else {
          standalone.push(item);
        }
      }

      const groupedCompanies = Array.from(groups.values())
        .map((group) => ({
          ...group,
          reports: group.reports.sort((a, b) => {
            const sourceA = a.source || "";
            const sourceB = b.source || "";
            return sourceA.localeCompare(sourceB, "ko");
          }),
        }))
        .sort((a, b) => {
          const countDiff = b.reports.length - a.reports.length;
          if (countDiff !== 0) return countDiff;
          return (a.company_name || "").localeCompare(b.company_name || "", "ko");
        });

      return { groupedCompanies, standalone };
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
      const counts = countByCategory(items);
      const grouped = groupedData(items);
      const root = document.getElementById("stats");
      const parts = [
        ["표시 중", items.length + "건"],
        ["묶인 종목", grouped.groupedCompanies.length + "개"],
        ["기업", (counts["기업"] || 0) + "건"],
        ["산업", (counts["산업"] || 0) + "건"],
        ["시장", (counts["시장"] || 0) + "건"],
      ];
      root.innerHTML = parts.map(([k, v]) => `<div class="stat">${k}: ${v}</div>`).join("");
      document.getElementById("subline").textContent =
        `대상 일자: ${currentDate} · 기업은 종목별로 묶어서 보고, 현재가는 py stocks.py 실행 시점 최신값을 함께 표시함`;
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
      return `
        <div class="price-box">
          <div class="price-label">현재가 ${safe(group.live_price_as_of)}</div>
          <div class="price-main">${safe(group.live_price)}원</div>
          <div class="price-change">${change}</div>
        </div>
      `;
    }

    function renderGrouped(items) {
      const root = document.getElementById("list");
      const { groupedCompanies, standalone } = groupedData(items);

      const groupCards = groupedCompanies.map((group) => {
        const uniqueSources = [...new Set(group.reports.map((report) => report.source).filter(Boolean))];
        const targets = [...new Set(group.reports.map((report) => report.target_price).filter(Boolean))];
        const opinions = [...new Set(group.reports.map((report) => report.investment_opinion).filter(Boolean))];

        const reportsHtml = group.reports.map((report) => {
          const pointItems = (report.key_points || []).slice(0, 4).map((point) => `<li>${safe(point)}</li>`).join("");
          return `
            <details class="report-detail">
              <summary>
                <div>
                  <div>${safe(report.title)}</div>
                  <div class="report-brief">${safe(report.source)} · ${safe(report.author)} · ${safe(report.investment_opinion)} ${safe(report.target_price)}</div>
                </div>
                <div class="badge">#${safe(report.report_idx)}</div>
              </summary>
              <div class="report-body">
                <p class="summary">${safe(report.summary)}</p>
                ${pointItems ? `<ul class="points">${pointItems}</ul>` : ""}
                <div class="links">
                  <a href="${safe(report.pdf_path)}" target="_blank" rel="noreferrer">PDF</a>
                  <a href="${safe(report.json_path)}" target="_blank" rel="noreferrer">JSON</a>
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
                <h2 class="group-title">${safe(group.company_name)}</h2>
                <div class="group-sub">종목코드 ${safe(group.stock_code)} · 리포트 ${group.reports.length}건</div>
              </div>
              ${renderPriceBox(group)}
            </div>
            <div class="group-meta">
              ${uniqueSources.slice(0, 6).map((source) => `<span class="pill">${safe(source)}</span>`).join("")}
              ${opinions.slice(0, 4).map((opinion) => `<span class="pill">의견 ${safe(opinion)}</span>`).join("")}
              ${targets.slice(0, 4).map((target) => `<span class="pill">목표가 ${safe(target)}</span>`).join("")}
            </div>
            <div class="report-stack">${reportsHtml}</div>
          </article>
        `;
      }).join("");

      const standaloneCards = standalone.length
        ? `
          <div class="section-title">산업/시장 리포트</div>
          ${standalone.map((item) => `
            <article class="card">
              <div class="row">
                <span class="badge">${safe(item.category)}</span>
                <span class="badge">${safe(item.published_at)}</span>
                <span class="badge">#${safe(item.report_idx)}</span>
              </div>
              <h2 class="title">${safe(item.title)}</h2>
              <div class="meta">작성자 ${safe(item.author)} · 출처 ${safe(item.source)}${item.market_topic ? ` · 주제 ${safe(item.market_topic)}` : ""}${item.industry_name ? ` · 산업 ${safe(item.industry_name)}` : ""}</div>
              <p class="summary">${safe(item.summary)}</p>
              <div class="links">
                <a href="${safe(item.pdf_path)}" target="_blank" rel="noreferrer">PDF</a>
                <a href="${safe(item.json_path)}" target="_blank" rel="noreferrer">JSON</a>
                <a href="${safe(item.pdf_url)}" target="_blank" rel="noreferrer">원문 링크</a>
              </div>
            </article>
          `).join("")}
        `
        : "";

      root.innerHTML = groupCards + standaloneCards;
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
            <p class="summary">${safe(item.summary)}</p>
            ${points ? `<ul class="points">${points}</ul>` : ""}
            <div class="links">
              <a href="${safe(item.pdf_path)}" target="_blank" rel="noreferrer">PDF</a>
              <a href="${safe(item.json_path)}" target="_blank" rel="noreferrer">JSON</a>
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
      if (!reportCache.has(nextDate)) {
        try {
          const response = await fetch(`./archive/${nextDate}.json`, { cache: "no-store" });
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const payload = await response.json();
          reportCache.set(nextDate, payload.reports || []);
        } catch (error) {
          alert(`브리프를 불러오지 못했습니다: ${nextDate}`);
          renderDateOptions();
          return;
        }
      }
      currentDate = nextDate;
      reports = reportCache.get(nextDate) || [];
      render();
    }

    render();
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
        {path.stem for path in archive_root.glob("*.json") if path.is_file()},
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


def main() -> None:
    configure_stdout()
    ensure_output_dirs()

    if not os.getenv("OPENAI_API_KEY"):
        raise EnvironmentError("OPENAI_API_KEY is not set.")

    session = build_session()
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    target_date = get_target_date(session)

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
                    mobile_entries.append(
                        build_mobile_entry(report, payload["summary"], pdf_path, json_path)
                    )
                except Exception:
                    pass
                print(f"[{category}][{index}/{len(reports)}] 중복 스킵: {report_idx}")
                continue

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
