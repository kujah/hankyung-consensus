import argparse
import json
import secrets
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlencode, urlparse

import requests


AUTH_URL = "https://kauth.kakao.com/oauth/authorize"
TOKEN_URL = "https://kauth.kakao.com/oauth/token"


class CallbackState:
    def __init__(self) -> None:
        self.done = threading.Event()
        self.code = ""
        self.error = ""
        self.error_description = ""


def build_handler(state: CallbackState, expected_state: str):
    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urlparse(self.path)
            query = parse_qs(parsed.query)
            returned_state = query.get("state", [""])[0]
            if returned_state and returned_state != expected_state:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"State mismatch.")
                state.error = "state_mismatch"
                state.error_description = "Returned state does not match."
                state.done.set()
                return

            state.code = query.get("code", [""])[0]
            state.error = query.get("error", [""])[0]
            state.error_description = query.get("error_description", [""])[0]

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            if state.code:
                body = "<h1>Kakao login completed.</h1><p>You can close this tab.</p>"
            else:
                body = (
                    "<h1>Kakao login failed.</h1>"
                    f"<p>{state.error or 'unknown_error'} {state.error_description}</p>"
                )
            self.wfile.write(body.encode("utf-8"))
            state.done.set()

        def log_message(self, format, *args):
            return

    return CallbackHandler


def exchange_code(
    rest_api_key: str,
    client_secret: str,
    redirect_uri: str,
    code: str,
) -> dict:
    response = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "authorization_code",
            "client_id": rest_api_key,
            "redirect_uri": redirect_uri,
            "code": code,
            "client_secret": client_secret,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rest-api-key", required=True)
    parser.add_argument("--client-secret", required=True)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8766)
    parser.add_argument("--open-browser", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    redirect_uri = f"http://{args.host}:{args.port}/callback"
    oauth_state = secrets.token_urlsafe(24)
    callback_state = CallbackState()
    server = HTTPServer((args.host, args.port), build_handler(callback_state, oauth_state))
    server.timeout = 1

    def serve():
        while not callback_state.done.is_set():
            server.handle_request()

    thread = threading.Thread(target=serve, daemon=True)
    thread.start()

    authorize_url = (
        f"{AUTH_URL}?{urlencode({
            'response_type': 'code',
            'client_id': args.rest_api_key,
            'redirect_uri': redirect_uri,
            'scope': 'talk_message',
            'state': oauth_state,
        })}"
    )

    print("1. Register this Redirect URI in Kakao Developers first:")
    print(redirect_uri)
    print()
    print("2. Open this URL and complete Kakao login/consent:")
    print(authorize_url)
    print()
    if args.open_browser:
        webbrowser.open(authorize_url)

    print("Waiting for callback...")
    deadline = time.time() + 300
    while time.time() < deadline and not callback_state.done.is_set():
        time.sleep(0.2)

    if not callback_state.done.is_set():
        print("Timed out waiting for Kakao callback.")
        return 1

    if not callback_state.code:
        print(
            "Authorization failed:",
            callback_state.error or "unknown_error",
            callback_state.error_description,
        )
        return 1

    token_payload = exchange_code(
        rest_api_key=args.rest_api_key,
        client_secret=args.client_secret,
        redirect_uri=redirect_uri,
        code=callback_state.code,
    )
    print()
    print("Token response:")
    print(json.dumps(token_payload, ensure_ascii=False, indent=2))
    print()
    print("Save this value to GitHub Actions secret KAKAO_REFRESH_TOKEN:")
    print(token_payload.get("refresh_token", ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
