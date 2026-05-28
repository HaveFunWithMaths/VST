import os
import json
import urllib.request
import urllib.error
from http.server import BaseHTTPRequestHandler

APPS_SCRIPT_URL = os.environ.get(
    'APPS_SCRIPT_URL',
    'https://script.google.com/macros/s/AKfycbyJIBkLcYnVQ7HDozxzyyOh6nMmvAipQIBY67DjA7KRjsywRBAkA5O9FN4aRcELuRDe/exec'
)


def call_apps_script(url, method='GET', payload=None):
    current_url = url
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0'
    } if payload is not None else {
        'User-Agent': 'Mozilla/5.0'
    }

    class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, hdrs, newurl):
            return None

    opener = urllib.request.build_opener(NoRedirectHandler())

    for _ in range(10):
        req = urllib.request.Request(current_url, data=data, headers=headers, method=method)
        try:
            response = opener.open(req)
            body = response.read().decode('utf-8')
            if not body or body.strip().startswith('<'):
                raise Exception("Google Apps Script returned an error page instead of JSON. Please redeploy the Apps Script in your Google Sheet.")
            return json.loads(body)
        except urllib.error.HTTPError as e:
            if e.code in (301, 302, 303, 307, 308):
                new_url = e.headers.get('Location')
                if not new_url:
                    raise Exception(f"Redirect without Location header (code {e.code})")
                current_url = new_url
                if e.code == 303:
                    method = 'GET'
                    data = None
                    headers.pop('Content-Type', None)
            else:
                error_body = e.read().decode('utf-8')
                raise Exception(f"HTTP {e.code}: {error_body}")
    raise Exception("Too many redirects")


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length)
        payload = json.loads(post_data.decode('utf-8'))

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()

        try:
            result = call_apps_script(APPS_SCRIPT_URL, method='POST', payload=payload)
            self.wfile.write(json.dumps(result).encode('utf-8'))
        except Exception as e:
            self.wfile.write(json.dumps({'status': 'error', 'message': str(e)}).encode('utf-8'))
