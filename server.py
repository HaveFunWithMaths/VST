import http.server
import socketserver
import os
import sys
import webbrowser

# Add api directory to sys.path to allow imports
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'api'))
import helper

PORT = 8000

class CustomHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS headers for development
        self.send_header('Access-Control-Allow-Origin', '*')
        super().end_headers()

    def do_OPTIONS(self):
        if self.path == '/api/update':
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type')
            self.end_headers()
        else:
            super().do_OPTIONS()

    def do_GET(self):
        if self.path == '/api/status':
            helper.handle_status(self)
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/update':
            helper.handle_update(self)
        else:
            self.send_error(404, "Not Found")

def run():
    # Make sure we are in the directory containing this script
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if base_dir:
        os.chdir(base_dir)
    
    handler = CustomHTTPRequestHandler
    
    # Allow port reuse to avoid 'Address already in use' errors on restart
    socketserver.TCPServer.allow_reuse_address = True
    
    try:
        with socketserver.TCPServer(("", PORT), handler) as httpd:
            print("==================================================")
            print("Bank Statement Ledger Updater Server is active!")
            print(f"Open in browser: http://localhost:{PORT}")
            print("==================================================")
            print("Press Ctrl+C to stop the server.")
            
            # Automatically open browser tab
            webbrowser.open(f"http://localhost:{PORT}")
            
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        sys.exit(0)
    except Exception as e:
        print(f"Error starting server: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run()
