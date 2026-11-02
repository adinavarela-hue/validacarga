"""F12 Privacidad · F13 Diseño adaptable · R8 Rendimiento."""
import time
from urllib.parse import urlparse

import pytest
from playwright.sync_api import expect
from conftest import DATOS, EJEMPLOS

PERMITIDOS = {'cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com'}


def test_cp054_el_archivo_no_sale_de_la_computadora(app, url_app, tmp_path):
    peticiones = []
    app.page.on('request', lambda r: peticiones.append((r.method, r.url, r.post_data)))
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    app.descargar('btn-descargar-reporte', tmp_path)
    app.descargar('btn-descargar-base', tmp_path)
    propio = urlparse(url_app).hostname
    for metodo, url, cuerpo in peticiones:
        u = urlparse(url)
        if u.scheme in ('blob', 'data'):
            continue  # descargas generadas dentro del navegador
        assert metodo == 'GET' and not cuerpo, f'Envío de datos: {metodo} {url}'
        assert u.hostname in PERMITIDOS | {propio}, f'Petición a un dominio no esperado: {url}'


def test_cp055_vista_de_celular_sin_desbordes(app):
    app.page.set_viewport_size({'width': 390, 'height': 844})
    assert app.page.evaluate('document.documentElement.scrollWidth <= window.innerWidth')


@pytest.mark.lento
@pytest.mark.parametrize('filas, limite_s', [(5000, 5), (20000, 15)], ids=['cp056_5000_filas', 'cp056_20000_filas'])
def test_cp056_rendimiento(app, filas, limite_s):
    inicio = time.perf_counter()
    app.subir_archivo(DATOS / f'masivo_{filas}.xlsx')
    expect(app.t('resumen-total')).to_have_text(f'{filas:,}'.replace(',', ''), timeout=limite_s * 1000)
    duracion = time.perf_counter() - inicio
    assert duracion < limite_s, f'{filas} filas tardaron {duracion:.1f} s (límite {limite_s} s)'
    assert app.resumen()['error'] == filas // 50
    expect(app.page.locator('.nota-tabla')).to_contain_text('Se muestran las primeras 500 filas')
