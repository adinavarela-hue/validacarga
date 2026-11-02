"""F05 Carga del archivo de datos (con la plantilla demo cargada)."""
import pytest
from playwright.sync_api import expect
from conftest import DATOS, EJEMPLOS

DATOS_DEMO = EJEMPLOS / 'datos_alta_usuarios_demo.xlsx'


@pytest.mark.parametrize('ruta', [DATOS_DEMO, DATOS / 'datos_alta_usuarios_demo.xls'], ids=['cp020_xlsx', 'cp023_xls'])
def test_archivo_excel(app, ruta):
    app.subir_archivo(ruta)
    assert app.resumen() == {'total': 9, 'correcto': 3, 'advertencia': 0, 'error': 6}


@pytest.mark.parametrize('archivo', ['csv_utf8_con_bom.csv', 'csv_utf8_sin_bom.csv'], ids=['cp021_con_bom', 'cp022_sin_bom'])
def test_csv_conserva_acentos(app, archivo):
    """Si los acentos se leyeran mal, "Región 5" y "Michoacán" no coincidirían con las listas y la fila saldría con error."""
    app.subir_archivo(DATOS / archivo)
    assert app.resumen() == {'total': 1, 'correcto': 1, 'advertencia': 0, 'error': 0}


def test_cp024_rechaza_pdf(app):
    app.subir_archivo(DATOS / 'documento.pdf', esperar=False)
    expect(app.t('mensaje')).to_contain_text('"documento.pdf" no es un Excel ni un CSV')
    expect(app.t('tabla-resultados')).to_have_count(0)


def test_cp025_archivo_vacio(app):
    app.subir_archivo(DATOS / 'archivo_vacio.xlsx', esperar=False)
    expect(app.t('mensaje')).to_contain_text('"archivo_vacio.xlsx" está vacío')


def test_cp026_titulos_arriba_numeracion_igual_a_excel(app):
    app.subir_archivo(DATOS / 'datos_con_titulo.xlsx')
    expect(app.page.locator('.titulo-res p')).to_contain_text('encabezados en la fila 4')
    assert sorted(app.estados()) == [5, 6]


def test_cp027_elige_la_pestana_por_sus_encabezados(app):
    app.subir_archivo(EJEMPLOS / 'datos_varias_pestanas_demo.xlsx')
    assert app.opcion_elegida('sel-hoja-archivo') == '2 · Altas noviembre'
    expect(app.t('archivo-actual')).to_contain_text('(elegida por sus encabezados)')
    assert app.resumen()['total'] == 9


def test_cp028_cambiar_de_pestana_a_mano(app):
    app.subir_archivo(EJEMPLOS / 'datos_varias_pestanas_demo.xlsx')
    app.t('sel-hoja-archivo').select_option('2')
    expect(app.t('aviso-archivo-err')).to_contain_text('Faltan columnas obligatorias')
    expect(app.t('archivo-actual')).not_to_contain_text('(elegida por sus encabezados)')
    app.t('sel-hoja-archivo').select_option('1')
    expect(app.t('resumen-total')).to_have_text('9')


def test_cp029_columnas_obligatorias_faltantes(app):
    app.subir_archivo(DATOS / 'columnas_faltantes.xlsx')
    expect(app.t('aviso-archivo-err')).to_contain_text('Faltan columnas obligatorias')
    expect(app.t('btn-descargar-validas')).to_be_disabled()


def test_cp029_fila_incompleta_no_aparece_como_correcta(app):
    """Regresión de VAL-1: antes la fila aparecía como "Correcta" aunque faltaban 6 columnas obligatorias."""
    app.subir_archivo(DATOS / 'columnas_faltantes.xlsx')
    assert app.estado_fila(2) == 'error'
    assert 'Fila incompleta: tu archivo no trae 5 columnas obligatorias' in app.texto_fila(2)
    assert app.resumen() == {'total': 1, 'correcto': 0, 'advertencia': 0, 'error': 1}
