"""F01 Carga de plantilla · F02 Detección de reglas · F04 Revisión de la plantilla."""
import re

import pytest
from playwright.sync_api import expect
from conftest import DATOS, EJEMPLOS

DEMO = EJEMPLOS / 'plantilla_oficial_demo.xlsx'


# ------------------------------------------------------------------ F01
def test_cp001_plantilla_xlsx_valida(app_limpia):
    app_limpia.subir_plantilla(DEMO)
    expect(app_limpia.t('nombre-plantilla')).to_have_text(DEMO.name)
    expect(app_limpia.t('plantilla-actual')).to_contain_text('Encabezados en la fila 1 · 9 columnas')
    assert len(app_limpia.reglas()) == 9
    expect(app_limpia.t('input-archivo')).to_be_enabled()


def test_cp002_rechaza_csv_como_plantilla(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'csv_utf8_sin_bom.csv')
    expect(app_limpia.t('mensaje')).to_contain_text('La plantilla debe ser .xlsx')
    assert app_limpia.reglas() == []
    expect(app_limpia.t('input-archivo')).to_be_disabled()


def test_cp003_plantilla_vacia(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_vacia.xlsx')
    expect(app_limpia.t('mensaje')).to_contain_text(re.compile('vacía|No encontré la fila de encabezados'))
    assert app_limpia.reglas() == []


@pytest.mark.parametrize('archivo', ['plantilla_instrucciones.xlsx', 'plantilla_hoja_vacia_primero.xlsx'],
                         ids=['cp004_instrucciones_primero', 'cp005_hoja_vacia_primero'])
def test_elige_la_pestana_del_formato(app_limpia, archivo):
    app_limpia.subir_plantilla(DATOS / archivo)
    expect(app_limpia.t('sel-hoja-plantilla')).to_be_visible()
    assert app_limpia.opcion_elegida('sel-hoja-plantilla') == '2 · Formato'
    assert len(app_limpia.reglas()) == 9


def test_cp006_cambiar_a_pestana_sin_encabezados_y_regresar(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_instrucciones.xlsx')
    app_limpia.t('sel-hoja-plantilla').select_option('0')
    expect(app_limpia.t('error-plantilla')).to_contain_text('No encontré la fila de encabezados')
    assert app_limpia.reglas() == []
    app_limpia.t('sel-hoja-plantilla').select_option('1')
    expect(app_limpia.t('reglas-detectadas')).to_be_visible()
    assert len(app_limpia.reglas()) == 9


def test_cp007_plantilla_con_titulo_arriba(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_titulo_x14.xlsx')
    expect(app_limpia.t('plantilla-actual')).to_contain_text('Encabezados en la fila 3 · 5 columnas')
    assert [r[0] for r in app_limpia.reglas()] == ['Sucursal', 'Fecha de Venta', 'IMEI', 'Teléfono', 'Canal']


# ------------------------------------------------------------------ F02
@pytest.mark.parametrize('columna, chip, valores', [
    ('Tipo de Canal*', 'Lista · 3', 'Detalle de Canal · Cadenas Comerciales · Masivo'),
    ('Región*', 'Lista · 3', 'Región 4 · Región 5 · Región 9'),
    ('Estado*', 'Depende de Región', None),
    ('Usuario*', '4–20 caracteres', None),
], ids=['cp008_lista_escrita', 'cp009_lista_nombre_definido', 'cp011_cascada', 'cp012_longitud'])
def test_reglas_de_la_plantilla_demo(app, columna, chip, valores):
    _, chips, titulos, _ = app.regla(columna)
    assert chip in chips
    if valores:
        assert valores in titulos


def test_cp012_valores_unicos(app):
    assert 'Sin repetidos' in app.regla('Usuario*')[1]


def test_cp010_lista_moderna_x14_y_cp013_largo_exacto(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_titulo_x14.xlsx')
    expect(app_limpia.t('reglas-detectadas')).to_be_visible()
    _, chips, titulos, _ = app_limpia.regla('Sucursal')
    assert 'Lista · 3' in chips and 'SUC CENTRO · SUC PLAZA · SUC NORTE' in titulos
    assert '15 caracteres' in app_limpia.regla('IMEI')[1]


def test_cp014_reglas_no_reconocidas(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_reglas_no_reconocidas.xlsx')
    expect(app_limpia.t('reglas-detectadas')).to_be_visible()
    for columna in ('Sucursal*', 'Producto*'):
        assert 'Regla no reconocida' in app_limpia.regla(columna)[1]
    assert app_limpia.revision_plantilla().count('no puede revisar esta regla') == 2


def test_cp015_plantilla_sin_listas_ni_marcas(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_sin_listas.xlsx')
    expect(app_limpia.t('reglas-detectadas')).to_contain_text('Sin marcas de obligatorio: se tomaron todas')
    reglas = app_limpia.reglas()
    assert all(r[3] for r in reglas), 'todas las columnas deben quedar obligatorias'
    assert 'Texto libre' in app_limpia.regla('Sucursal')[1]
    assert 'Fecha?' in app_limpia.regla('Fecha')[1]


# ------------------------------------------------------------------ F04
def test_cp017_fila_de_muestra_no_pasa_sus_listas(app):
    texto = app.revision_plantilla()
    assert '2 hallazgos' in texto
    assert 'La fila de muestra 2 de la propia plantilla no pasa sus reglas' in texto and 'Tipo de Canal' in texto
    assert 'No hay lista de Estado para: Región 4' in texto


def test_cp018_opciones_con_espacios_sobrantes(app_limpia):
    app_limpia.subir_plantilla(DATOS / 'plantilla_titulo_x14.xlsx')
    expect(app_limpia.t('revision-plantilla')).to_contain_text('opciones con espacios sobrantes ("SUC NORTE ")')


def test_cp019_archivo_acumulado_como_plantilla(app_limpia):
    app_limpia.subir_plantilla(EJEMPLOS / 'base_ventas_demo.xlsx')
    expect(app_limpia.t('revision-plantilla')).to_contain_text('Esta plantilla ya trae datos (más de 50 filas)')
    assert 'fila de muestra' not in app_limpia.revision_plantilla()
