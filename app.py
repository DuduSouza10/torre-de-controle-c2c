import csv
import io
import json
import os
import shutil
import threading
import time
from datetime import datetime, timezone
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
INDEX = ROOT / 'templates' / 'index.html'
INITIAL_DATA = ROOT / 'data' / 'initial_dataset.csv'
MAX_UPLOAD_BYTES = int(os.environ.get('MAX_UPLOAD_MB', '50')) * 1024 * 1024

FIELDNAMES = [
    'rm', 'reg', 'pedido', 'coleta', 'estcol', 'envio', 'tipo', 'oprec', 'horaop',
    'baseent', 'uf', 'motivo', 'marca', 'transf', 'interceptado',
    'indenizado', 'finalizado'
]


def choose_data_dir() -> Path:
    configured = os.environ.get('DATA_DIR')
    railway_volume = os.environ.get('RAILWAY_VOLUME_MOUNT_PATH')
    candidates = []
    if configured:
        candidates.append(Path(configured))
    if railway_volume:
        candidates.append(Path(railway_volume))
    if os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('RAILWAY_PROJECT_ID'):
        candidates.append(Path('/data'))
    candidates.append(ROOT / 'data' / 'runtime')

    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / '.write_test'
            probe.write_text('ok', encoding='utf-8')
            probe.unlink(missing_ok=True)
            return candidate
        except OSError:
            continue
    raise RuntimeError('Nenhum diretório gravável disponível para os dados compartilhados.')


DATA_DIR = choose_data_dir()
STATE_CSV = DATA_DIR / 'shared_dataset.csv'
STATE_META = DATA_DIR / 'shared_meta.json'
DATA_LOCK = threading.RLock()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')


def atomic_write_bytes(path: Path, data: bytes) -> None:
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_bytes(data)
    os.replace(tmp, path)


def atomic_write_text(path: Path, text: str) -> None:
    atomic_write_bytes(path, text.encode('utf-8'))


def read_csv_rows(path: Path):
    if not path.exists():
        return []
    with path.open('r', encoding='utf-8-sig', newline='') as fh:
        return [dict(row) for row in csv.DictReader(fh)]


def write_csv_rows(path: Path, rows) -> None:
    buf = io.StringIO(newline='')
    writer = csv.DictWriter(buf, fieldnames=FIELDNAMES, extrasaction='ignore', lineterminator='\n')
    writer.writeheader()
    for row in rows:
        writer.writerow({key: row.get(key, '') or '' for key in FIELDNAMES})
    atomic_write_text(path, buf.getvalue())


def read_meta():
    try:
        return json.loads(STATE_META.read_text(encoding='utf-8'))
    except Exception:
        rows = read_csv_rows(STATE_CSV)
        return {'version': 'unknown', 'updatedAt': None, 'count': len(rows)}


def write_meta(count: int, source: str = 'upload'):
    meta = {
        'version': str(time.time_ns()),
        'updatedAt': utc_now_iso(),
        'count': count,
        'source': source,
    }
    atomic_write_text(STATE_META, json.dumps(meta, ensure_ascii=False))
    return meta


def ensure_shared_dataset():
    with DATA_LOCK:
        if STATE_CSV.exists():
            if not STATE_META.exists():
                write_meta(len(read_csv_rows(STATE_CSV)), source='recovered')
            return
        if not INITIAL_DATA.exists():
            write_csv_rows(STATE_CSV, [])
            write_meta(0, source='empty')
            return
        # Normalize the embedded snapshot into the shared schema on first startup.
        initial_rows = read_csv_rows(INITIAL_DATA)
        write_csv_rows(STATE_CSV, initial_rows)
        write_meta(len(initial_rows), source='initial')


def merge_incoming_csv(csv_text: str):
    reader = csv.DictReader(io.StringIO(csv_text))
    if not reader.fieldnames or 'pedido' not in reader.fieldnames:
        raise ValueError('CSV sem a coluna normalizada "pedido".')

    incoming = []
    for raw in reader:
        pedido = (raw.get('pedido') or '').strip()
        if not pedido:
            continue
        row = {key: (raw.get(key) or '') for key in FIELDNAMES}
        row['pedido'] = pedido
        incoming.append(row)

    with DATA_LOCK:
        current = read_csv_rows(STATE_CSV)
        by_id = {(row.get('pedido') or '').strip(): row for row in current if (row.get('pedido') or '').strip()}
        for row in incoming:
            by_id[row['pedido']] = row
        merged = list(by_id.values())
        write_csv_rows(STATE_CSV, merged)
        meta = write_meta(len(merged), source='upload')
    return meta, len(incoming)


ensure_shared_dataset()


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def send_json(self, status, payload):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/health':
            meta = read_meta()
            self.send_json(200, {'status': 'ok', 'sharedData': meta})
            return

        if path == '/api/meta':
            with DATA_LOCK:
                meta = read_meta()
            self.send_json(200, meta)
            return

        if path == '/api/data':
            with DATA_LOCK:
                ensure_shared_dataset()
                data = STATE_CSV.read_bytes()
                meta = read_meta()
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
            self.send_header('X-Data-Version', str(meta.get('version') or ''))
            self.send_header('X-Updated-At', str(meta.get('updatedAt') or ''))
            self.send_header('X-Data-Count', str(meta.get('count') or 0))
            self.end_headers()
            self.wfile.write(data)
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

    def do_POST(self):
        path = urlparse(self.path).path
        if path != '/api/upload':
            self.send_json(404, {'error': 'Rota não encontrada.'})
            return

        try:
            content_length = int(self.headers.get('Content-Length', '0'))
        except ValueError:
            content_length = 0
        if content_length <= 0:
            self.send_json(400, {'error': 'Upload vazio.'})
            return
        if content_length > MAX_UPLOAD_BYTES:
            self.send_json(413, {'error': f'Upload excede o limite de {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.'})
            return

        try:
            raw = self.rfile.read(content_length)
            text = raw.decode('utf-8-sig')
            meta, received = merge_incoming_csv(text)
            self.send_json(200, {
                'ok': True,
                'received': received,
                'count': meta['count'],
                'version': meta['version'],
                'updatedAt': meta['updatedAt'],
            })
        except UnicodeDecodeError:
            self.send_json(400, {'error': 'Conteúdo enviado não está em UTF-8.'})
        except ValueError as exc:
            self.send_json(400, {'error': str(exc)})
        except Exception as exc:
            print(f'Erro no upload: {exc}', flush=True)
            self.send_json(500, {'error': 'Falha interna ao salvar a base compartilhada.'})

    def log_message(self, fmt, *args):
        print(f'{self.address_string()} - {fmt % args}', flush=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8080'))
    server = ThreadingHTTPServer(('0.0.0.0', port), Handler)
    print(f'Torre de Controle C2C rodando na porta {port}', flush=True)
    print(f'Dados compartilhados em: {STATE_CSV}', flush=True)
    server.serve_forever()
