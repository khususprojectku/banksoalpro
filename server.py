"""Server lokal BankSoalPro dengan penyimpanan gambar soal."""
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from uuid import uuid4
import json

BASE_DIR = Path(__file__).resolve().parent
IMAGE_DIR = BASE_DIR / 'gambarsoal'
IMAGE_DIR.mkdir(exist_ok=True)
MAX_IMAGE_SIZE = 5 * 1024 * 1024
ALLOWED_TYPES = {
    'image/jpeg': '.jpg', 'image/png': '.png', 'image/gif': '.gif',
    'image/webp': '.webp', 'image/svg+xml': '.svg'
}


class BankSoalHandler(SimpleHTTPRequestHandler):
    def _send_json(self, status, payload):
        data = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        if self.path != '/api/gambarsoal':
            self._send_json(HTTPStatus.NOT_FOUND, {'message': 'Endpoint tidak ditemukan.'})
            return

        try:
            content_length = int(self.headers.get('Content-Length', '0'))
        except ValueError:
            self._send_json(HTTPStatus.BAD_REQUEST, {'message': 'Ukuran unggahan tidak valid.'})
            return
        if content_length <= 0 or content_length > MAX_IMAGE_SIZE + 1024 * 64:
            self._send_json(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, {'message': 'Ukuran gambar maksimal 5 MB.'})
            return

        try:
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data'):
                raise ValueError('Format unggahan tidak valid.')
            body = self.rfile.read(content_length)
            message = BytesParser(policy=policy.default).parsebytes(
                f'Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n'.encode('utf-8') + body
            )
            part = next((item for item in message.iter_parts() if item.get_param('name', header='content-disposition') == 'image'), None)
            if part is None:
                raise ValueError('Berkas gambar tidak ditemukan.')
            image_data = part.get_payload(decode=True) or b''
            image_type = part.get_content_type()
            if image_type not in ALLOWED_TYPES or not image_data or len(image_data) > MAX_IMAGE_SIZE:
                raise ValueError('Gunakan gambar PNG, JPG, GIF, WebP, atau SVG dengan ukuran maksimal 5 MB.')

            filename = f'{uuid4().hex}{ALLOWED_TYPES[image_type]}'
            (IMAGE_DIR / filename).write_bytes(image_data)
            self._send_json(HTTPStatus.CREATED, {'url': f'/gambarsoal/{filename}'})
        except (ValueError, StopIteration) as error:
            self._send_json(HTTPStatus.BAD_REQUEST, {'message': str(error)})
        except Exception:
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {'message': 'Gambar tidak dapat disimpan.'})

    def end_headers(self):
        if self.path.startswith('/gambarsoal/'):
            self.send_header('Cache-Control', 'public, max-age=31536000, immutable')
        super().end_headers()


if __name__ == '__main__':
    print('BankSoalPro berjalan di http://localhost:8000')
    ThreadingHTTPServer(('', 8000), BankSoalHandler).serve_forever()
