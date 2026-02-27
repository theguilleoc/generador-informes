#!/usr/bin/env python3
"""HTTP server for the report generator web app.
Run: python app.py
Open: http://127.0.0.1:8080/report.html
"""
import os, sys, json, base64, tempfile, webbrowser, threading, importlib
from http.server import HTTPServer, SimpleHTTPRequestHandler

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import generate_report


class ReportHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_POST(self):
        if self.path == '/generate':
            self.handle_generate()
        else:
            self.send_error(404)

    def handle_generate(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(length)
            data = json.loads(body)

            with tempfile.TemporaryDirectory() as tmpdir:
                input_files = {}
                for slot in ['O1', 'C', 'OCO', 'CO', 'O2']:
                    if slot not in data.get('files', {}):
                        raise ValueError(f'Falta archivo: {slot}')
                    file_bytes = base64.b64decode(data['files'][slot])
                    path = os.path.join(tmpdir, f'{slot}.docx')
                    with open(path, 'wb') as f:
                        f.write(file_bytes)
                    input_files[slot] = path

                template_path = os.path.join(BASE_DIR, 'template.docx')
                output_path = os.path.join(tmpdir, 'INFORME_GENERADO.docx')

                config = data.get('config', {})
                fields = data.get('fields', {})
                importlib.reload(generate_report)
                generate_report.generate(input_files, template_path, output_path, config=config, fields=fields)

                with open(output_path, 'rb') as f:
                    docx_bytes = f.read()

                self.send_response(200)
                self.send_header('Content-Type',
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition',
                    'attachment; filename="INFORME_GENERADO.docx"')
                self.send_header('Content-Length', len(docx_bytes))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(docx_bytes)

        except Exception as e:
            import traceback
            traceback.print_exc()
            msg = json.dumps({'error': str(e)}).encode()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', len(msg))
            self.end_headers()
            self.wfile.write(msg)

    def log_message(self, format, *args):
        if '/generate' in str(args[0]) if args else False:
            super().log_message(format, *args)


def main():
    port = int(os.environ.get('PORT', 8080))
    server = HTTPServer(('0.0.0.0', port), ReportHandler)
    url = f'http://127.0.0.1:{port}/report.html'
    print(f'Servidor corriendo en {url}')
    print('Presiona Ctrl+C para detener')
    if not os.environ.get('RENDER'):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nDeteniendo servidor...')
        server.server_close()


if __name__ == '__main__':
    main()
