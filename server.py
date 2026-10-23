# -*- coding: utf-8 -*-
"""
Servidor local para el Dashboard de Evaluación PAC 2026 - Línea Verde
Permite consultar los datos procesados en tiempo real.
Cualquier cambio en 'Participantes PAC 2026.xlsx' se refleja de inmediato al refrescar.
"""

import os
import sys
import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from parser import parse_excel_data

PORT = 8085
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR.parent / "datos_evaluacion" / "Participantes PAC 2026.xlsx"
if not EXCEL_PATH.exists() and (BASE_DIR / "ParticipantesPAC2026_live.xlsx").exists():
    EXCEL_PATH = BASE_DIR / "ParticipantesPAC2026_live.xlsx"

class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_GET(self):
        if self.path == '/api/data':
            try:
                data = parse_excel_data(str(EXCEL_PATH))
                response_bytes = json.dumps(data, ensure_ascii=False).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(response_bytes)))
                self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
                self.end_headers()
                self.wfile.write(response_bytes)
            except Exception as e:
                err_msg = json.dumps({'error': str(e)}).encode('utf-8')
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(err_msg)
        else:
            super().do_GET()

def run():
    server_address = ('', PORT)
    httpd = ThreadingHTTPServer(server_address, DashboardHandler)
    print("=" * 65)
    print(f"  Dashboard Interactivo PAC 2026 - Línea Verde iniciado")
    print(f"  Accede en tu navegador a: http://localhost:{PORT}")
    print(f"  Archivo Excel vinculado: {EXCEL_PATH}")
    print("=" * 65)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
        httpd.server_close()

if __name__ == '__main__':
    run()
