from http.server import BaseHTTPRequestHandler
import helper

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        helper.handle_status(self)
