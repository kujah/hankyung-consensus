from __future__ import annotations

import argparse
import html
import json
import os
import smtplib
import ssl
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise EnvironmentError(f"{name} is not set.")
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send Hankyung consensus digest via Gmail.")
    parser.add_argument("--input", default="reports_mobile/latest.json")
    parser.add_argument("--recipient", default="")
    return parser.parse_args()


def load_payload(path: str) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def group_company_reports(reports: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for report in reports:
        stock_code = str(report.get("stock_code") or "").strip()
        if not stock_code:
            continue
        bucket = grouped.setdefault(
            stock_code,
            {
                "stock_code": stock_code,
                "company_name": report.get("company_name") or report.get("title") or stock_code,
                "reports": [],
                "best_alpha_score": -1,
                "best_alpha_grade": "",
            },
        )
        bucket["reports"].append(report)
        alpha_score = report.get("alpha_score")
        if isinstance(alpha_score, (int, float)) and alpha_score > bucket["best_alpha_score"]:
            bucket["best_alpha_score"] = int(alpha_score)
            bucket["best_alpha_grade"] = str(report.get("alpha_grade") or "")

    result = list(grouped.values())
    result.sort(
        key=lambda item: (
            item["best_alpha_score"],
            len(item["reports"]),
            item["company_name"],
        ),
        reverse=True,
    )
    return result


def render_report_line(report: dict[str, Any]) -> str:
    meta_parts = [
        str(report.get("published_at") or ""),
        str(report.get("source") or ""),
        str(report.get("author") or ""),
        " ".join(
            part for part in [
                str(report.get("investment_opinion") or ""),
                str(report.get("target_price") or ""),
            ]
            if part
        ),
    ]
    badges = []
    if report.get("alpha_grade") and report.get("alpha_score") is not None:
        badges.append(f'{report["alpha_grade"]} {report["alpha_score"]}')
    badges.append(f'#{report.get("report_idx", "")}')
    meta_text = " / ".join(part for part in meta_parts if part)
    badge_text = " / ".join(part for part in badges if part)
    summary = html.escape(str(report.get("summary") or ""))
    summary_html = (
        f"<div style='margin-top:6px;color:#425466;font-size:13px;line-height:1.55'>{summary}</div>"
        if summary
        else ""
    )
    return (
        "<div style='border:1px solid #e6d9c2;border-radius:16px;padding:14px 16px;background:#fffdf8;margin-top:10px'>"
        f"<div style='font-size:12px;color:#6f7d86;margin-bottom:6px'>{html.escape(badge_text)}</div>"
        f"<div style='font-size:18px;font-weight:700;line-height:1.35;color:#16202a'>{html.escape(str(report.get('title') or ''))}</div>"
        f"<div style='margin-top:6px;color:#536471;font-size:13px;line-height:1.5'>{html.escape(meta_text)}</div>"
        f"{summary_html}"
        "</div>"
    )


def render_email_html(payload: dict[str, Any]) -> str:
    reports = list(payload.get("reports") or [])
    target_date = str(payload.get("target_date") or "")
    generated_at = str(payload.get("generated_at") or "")
    page_url = os.getenv("HANKYUNG_PAGE_URL", "").strip()

    categories = defaultdict(int)
    for report in reports:
        categories[str(report.get("category") or "Other")] += 1

    company_groups = group_company_reports(reports)
    market_reports = [
        report
        for report in reports
        if not report.get("stock_code") and report.get("market_topic")
    ]
    industry_reports = [
        report
        for report in reports
        if not report.get("stock_code") and report.get("industry_name")
    ]

    chips = "".join(
        f"<span style='display:inline-block;background:#eef5f6;color:#0b6578;border-radius:999px;padding:7px 11px;font-size:12px;font-weight:700;margin:0 8px 8px 0'>{html.escape(name)} {count}</span>"
        for name, count in sorted(categories.items())
    )

    sections: list[str] = [
        "<html><body style='margin:0;background:#f7efe3;color:#17212b;font-family:Segoe UI,Apple SD Gothic Neo,sans-serif'>",
        "<div style='max-width:960px;margin:0 auto;padding:24px 16px 40px'>",
        "<div style='background:#fffdf8;border:1px solid #e8dcc7;border-radius:24px;padding:24px 24px 18px;box-shadow:0 12px 30px rgba(80,60,20,0.08)'>",
        "<div style='font-size:28px;font-weight:800;line-height:1.2'>Hankyung Consensus Daily</div>",
        f"<div style='margin-top:8px;color:#5d6b75;font-size:14px'>Date {html.escape(target_date)} / Reports {len(reports)} / Generated {html.escape(generated_at)}</div>",
        f"<div style='margin-top:16px'>{chips}</div>",
    ]
    if page_url:
        sections.append(
            f"<div style='margin-top:6px'><a href='{html.escape(page_url)}' style='color:#b26000;font-weight:700;text-decoration:none'>Open mobile page</a></div>"
        )
    sections.append("</div>")

    if market_reports:
        sections.append("<div style='margin-top:18px'>")
        sections.append("<div style='font-size:18px;font-weight:800;margin-bottom:6px'>Market Reports</div>")
        sections.extend(render_report_line(report) for report in market_reports[:3])
        sections.append("</div>")

    if industry_reports:
        sections.append("<div style='margin-top:18px'>")
        sections.append("<div style='font-size:18px;font-weight:800;margin-bottom:6px'>Industry Reports</div>")
        sections.extend(render_report_line(report) for report in industry_reports[:3])
        sections.append("</div>")

    if company_groups:
        sections.append("<div style='margin-top:18px'>")
        sections.append("<div style='font-size:18px;font-weight:800;margin-bottom:6px'>Company Reports</div>")
        for company in company_groups[:12]:
            best_score_text = (
                f" / Best {html.escape(company['best_alpha_grade'])} {company['best_alpha_score']}"
                if company["best_alpha_score"] >= 0
                else ""
            )
            sections.append(
                "<div style='background:#fffdf8;border:1px solid #e8dcc7;border-radius:20px;padding:18px 18px 8px;margin-top:12px'>"
                f"<div style='font-size:22px;font-weight:800;line-height:1.25'>{html.escape(company['company_name'])}</div>"
                f"<div style='margin-top:6px;color:#5d6b75;font-size:13px'>Code {html.escape(company['stock_code'])} / Reports {len(company['reports'])}{best_score_text}</div>"
            )
            for report in company["reports"][:3]:
                sections.append(render_report_line(report))
            sections.append("</div>")
        sections.append("</div>")

    sections.append("</div></body></html>")
    return "".join(sections)


def send_email(subject: str, html_body: str, recipient_override: str = "") -> None:
    sender = require_env("GMAIL_USERNAME")
    password = require_env("GMAIL_APP_PASSWORD")
    recipient = recipient_override.strip() or os.getenv("HANKYUNG_EMAIL_TO", "").strip() or sender

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = recipient
    message.attach(MIMEText(html_body, "html", "utf-8"))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender, password)
        server.sendmail(sender, [recipient], message.as_string())


def main() -> int:
    args = parse_args()
    payload = load_payload(args.input)
    target_date = str(payload.get("target_date") or "")
    report_count = len(payload.get("reports") or [])
    subject = f"[Hankyung Consensus] {target_date} / {report_count} reports"
    html_body = render_email_html(payload)
    send_email(subject, html_body, args.recipient)
    print(f"Sent Hankyung email for {target_date} ({report_count} reports).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
