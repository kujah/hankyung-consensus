import argparse
import json
import os
import sys
from typing import Any

import requests


KAKAO_TOKEN_URL = "https://kauth.kakao.com/oauth/token"
KAKAO_MEMO_URL = "https://kapi.kakao.com/v2/api/talk/memo/default/send"


def env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value


def refresh_access_token() -> str:
    data = {
        "grant_type": "refresh_token",
        "client_id": env("KAKAO_REST_API_KEY"),
        "refresh_token": env("KAKAO_REFRESH_TOKEN"),
        "client_secret": env("KAKAO_CLIENT_SECRET"),
    }
    response = requests.post(KAKAO_TOKEN_URL, data=data, timeout=30)
    response.raise_for_status()
    payload = response.json()
    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError(f"Failed to refresh Kakao access token: {payload}")
    if payload.get("refresh_token"):
        print(
            "::warning::Kakao returned a renewed refresh token. Update the "
            "KAKAO_REFRESH_TOKEN secret manually before the old token expires."
        )
    return access_token


def build_message(args: argparse.Namespace) -> str:
    status_map = {
        "DONE": "완료",
        "FAILED": "실패",
    }
    note_map = {
        "updated": "모바일 브리프와 아카이브가 갱신되었습니다.",
        "collect_or_deploy_failed": "수집 또는 배포 단계 중 하나가 실패했습니다.",
        "deploy_failed": "수집은 완료됐지만 GitHub Pages 배포가 실패했습니다.",
    }
    base = [
        f"[한경 컨센서스] {status_map.get(args.status_label, args.status_label)}",
        f"대상 일자: {args.target_date or '확인 불가'}",
        f"리포트 수: {args.report_count or '확인 불가'}",
    ]
    if args.page_url:
        base.append(f"페이지: {args.page_url}")
    if args.run_url:
        base.append(f"실행 로그: {args.run_url}")
    if args.note:
        base.append(f"메모: {note_map.get(args.note, args.note)}")
    return "\n".join(base)


def build_template(args: argparse.Namespace) -> dict[str, Any]:
    landing_url = args.page_url or args.run_url or "https://github.com"
    return {
        "object_type": "text",
        "text": build_message(args),
        "link": {
            "web_url": landing_url,
            "mobile_web_url": landing_url,
        },
        "button_title": "브리프 열기" if args.page_url else "실행 로그 보기",
    }


def send_message(access_token: str, template_object: dict[str, Any]) -> None:
    response = requests.post(
        KAKAO_MEMO_URL,
        headers={"Authorization": f"Bearer {access_token}"},
        data={"template_object": json.dumps(template_object, ensure_ascii=False)},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    if payload.get("result_code") != 0:
        raise RuntimeError(f"Failed to send Kakao message: {payload}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--status-label", required=True)
    parser.add_argument("--target-date", default="")
    parser.add_argument("--report-count", default="")
    parser.add_argument("--page-url", default="")
    parser.add_argument("--run-url", default="")
    parser.add_argument("--note", default="")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        access_token = refresh_access_token()
        template_object = build_template(args)
        send_message(access_token, template_object)
        print("Kakao notification sent.")
        return 0
    except Exception as exc:
        print(f"::error::{exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
