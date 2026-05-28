from http.server import BaseHTTPRequestHandler
import _helper as helper

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        helper.handle_status(self)
