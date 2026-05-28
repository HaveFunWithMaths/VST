import os
import json
import requests

def get_apps_script_url():
    # Paths to search for config.json
    paths_to_check = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json'),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json'),
        'config.json'
    ]
    for p in paths_to_check:
        if os.path.exists(p):
            try:
                with open(p, 'r') as f:
                    config = json.load(f)
                    url = config.get('apps_script_url')
                    if url:
                        return url
            except Exception as e:
                print(f"Error reading {p}: {e}")
    # Fallback to env variable or hardcoded default
    return os.environ.get(
        'APPS_SCRIPT_URL',
        'https://script.google.com/macros/s/AKfycbytLE8PgbHy4eAba2eTUG2BQ6bZLnU9r4-huJ_JmhdtLIcCOPwuPmWtF7tI9MYyEog/exec'
    )

def call_apps_script(url, method='GET', payload=None):
    """
    Call the Google Apps Script web app.
    Uses the `requests` library so that redirect handling (including the
    302 POST redirect that Apps Script issues) is managed correctly.
    The `requests` library preserves the POST body across redirects,
    whereas the stdlib urllib NoRedirectHandler approach broke POST calls.
    """
    headers = {'User-Agent': 'Mozilla/5.0', 'Content-Type': 'application/json'}

    try:
        if method == 'POST':
            response = requests.post(url, json=payload, headers=headers, timeout=60)
        else:
            response = requests.get(url, headers=headers, timeout=30)

        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as e:
        raise Exception(f"HTTP {e.response.status_code}: {e.response.text[:500]}")
    except requests.exceptions.RequestException as e:
        raise Exception(f"Request failed: {str(e)}")

def handle_status(handler_inst):
    handler_inst.send_response(200)
    handler_inst.send_header('Content-Type', 'application/json')
    handler_inst.send_header('Access-Control-Allow-Origin', '*')
    handler_inst.end_headers()

    try:
        url = get_apps_script_url()
        result = call_apps_script(url, method='GET')
        handler_inst.wfile.write(json.dumps(result).encode('utf-8'))
    except Exception as e:
        handler_inst.wfile.write(json.dumps({
            'status': 'error',
            'message': str(e)
        }).encode('utf-8'))

def handle_update(handler_inst):
    # Read POST body
    content_length = int(handler_inst.headers.get('Content-Length', 0))
    post_data = handler_inst.rfile.read(content_length)
    payload = json.loads(post_data.decode('utf-8'))

    handler_inst.send_response(200)
    handler_inst.send_header('Content-Type', 'application/json')
    handler_inst.send_header('Access-Control-Allow-Origin', '*')
    handler_inst.end_headers()

    try:
        url = get_apps_script_url()
        result = call_apps_script(url, method='POST', payload=payload)
        handler_inst.wfile.write(json.dumps(result).encode('utf-8'))
    except Exception as e:
        handler_inst.wfile.write(json.dumps({
            'status': 'error',
            'message': str(e)
        }).encode('utf-8'))
