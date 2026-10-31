from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import time
class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/slow': time.sleep(2)
        if self.path == '/trickle':
            self.send_response(200)
            self.send_header('Content-Length', '100')
            self.end_headers()
            try:
                for _ in range(100):
                    self.wfile.write(b'x')
                    self.wfile.flush()
                    time.sleep(0.03)
            except (BrokenPipeError, ConnectionResetError): pass
            return
        if self.path == '/disconnect':
            self.connection.close()
            return
        body = b'' if self.path == '/empty' else b'hello'
        self.send_response(200)
        self.send_header('Content-Length', str(len(body)))
        try:
            self.end_headers()
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError): pass
if __name__ == '__main__':
    ThreadingHTTPServer(('127.0.0.1', 18761), Handler).serve_forever()
