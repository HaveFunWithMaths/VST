import os
import json
import urllib.request
import urllib.error

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
    # Fallback to env variable or default
    return os.environ.get(
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
    
    # We will manually follow redirects to ensure the method and body are preserved correctly.
    # To do this, we disable automatic redirect following.
    class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, hdrs, newurl):
            return None

    opener = urllib.request.build_opener(NoRedirectHandler())
    
    for redirect_count in range(10):
        req = urllib.request.Request(current_url, data=data, headers=headers, method=method)
        try:
            response = opener.open(req)
            body = response.read().decode('utf-8')
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
                    if 'Content-Type' in headers:
                        del headers['Content-Type']
            else:
                error_body = e.read().decode('utf-8')
                raise Exception(f"HTTP {e.code}: {error_body}")
    raise Exception("Too many redirects")

def handle_status(handler_inst):
    # Enable CORS
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
    # read POST data
    content_length = int(handler_inst.headers.get('Content-Length', 0))
    post_data = handler_inst.rfile.read(content_length)
    payload = json.loads(post_data.decode('utf-8'))
    
    # Enable CORS
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
