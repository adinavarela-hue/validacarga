"""Configuración común de la suite.

- Levanta la app con un servidor local (o usa la demo publicada si se define VC_URL).
- Entrega la página abierta como Page Object (`app`) y falla la prueba si hubo errores de JavaScript.
- VC_LIBS (opcional): carpeta con xlsx.full.min.js y exceljs.min.js para correr sin acceso al CDN.
"""
import functools
import http.server
import os
import pathlib
import sys
import threading

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from paginas.validacarga import ValidaCarga  # noqa: E402

RAIZ = pathlib.Path(__file__).parent.parent
DATOS = pathlib.Path(__file__).parent / 'datos'
EJEMPLOS = RAIZ / 'ejemplos'


class _Silencioso(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass


@pytest.fixture(scope='session')
def url_app():
    if os.environ.get('VC_URL'):
        yield os.environ['VC_URL']
        return
    servidor = http.server.ThreadingHTTPServer(('127.0.0.1', 0), functools.partial(_Silencioso, directory=str(RAIZ)))
    threading.Thread(target=servidor.serve_forever, daemon=True).start()
    yield f'http://127.0.0.1:{servidor.server_port}/'
    servidor.shutdown()


def _librerias_locales(page):
    libs = os.environ.get('VC_LIBS')
    if not libs:
        return
    js = 'application/javascript'
    page.route('**/ajax/libs/xlsx/**', lambda r: r.fulfill(path=f'{libs}/xlsx.full.min.js', content_type=js))
    page.route('**/ajax/libs/exceljs/**', lambda r: r.fulfill(path=f'{libs}/exceljs.min.js', content_type=js))
    page.route('**/fonts.g*/**', lambda r: r.abort())


@pytest.fixture
def errores_js(page):
    errores = []
    page.on('pageerror', lambda e: errores.append(str(e)))
    return errores


@pytest.fixture
def app(page, url_app, errores_js):
    _librerias_locales(page)
    vc = ValidaCarga(page).abrir(url_app)
    yield vc
    assert not errores_js, f'Errores de JavaScript en la página: {errores_js}'


@pytest.fixture
def app_limpia(app):
    """App sin plantilla ni archivo (modo Plantilla oficial)."""
    app.limpiar()
    return app
