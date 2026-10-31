from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import time
class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/slow': time.sleep(2)
        if self.path == '/disconnect':
            self.connection.close()
            return
        body = b'' if self.path == '/empty' else b'hello'
        self.send_response(200)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        try: self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError): pass
ThreadingHTTPServer(('127.0.0.1', 18761), Handler).serve_forever()
