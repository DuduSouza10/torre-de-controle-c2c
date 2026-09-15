import json
import os
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
INDEX = ROOT / 'templates' / 'index.html'

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/health':
            payload = json.dumps({'status': 'ok'}).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return
        if path in ('/', '/index.html'):
            data = INDEX.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(data)
            return
        return super().do_GET()

    def log_message(self, fmt, *args):
        print(f'{self.address_string()} - {fmt % args}', flush=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8080'))
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    print(f'Torre de Controle C2C rodando na porta {port}', flush=True)
    server.serve_forever()
