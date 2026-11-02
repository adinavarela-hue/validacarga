"""F09 Descargas: se abre cada Excel descargado con openpyxl y se revisan celdas, colores y pestañas."""
import zipfile

import openpyxl
from playwright.sync_api import expect
from conftest import DATOS, EJEMPLOS

ROJO, AMBAR = 'FFF8D3CF', 'FFFCEBC6'


def relleno(celda):
    return celda.fill.fgColor.rgb if celda.fill and celda.fill.fill_type else None


def test_cp044_reporte_de_errores(app, tmp_path):
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    incidencias = app.incidencias_en_pantalla()
    archivo = app.descargar('btn-descargar-reporte', tmp_path)
    assert archivo.name == 'datos_alta_usuarios_demo_reporte_errores.xlsx'
    filas = list(openpyxl.load_workbook(archivo).active.iter_rows(values_only=True))
    assert filas[0] == ('Fila', 'Estado de la fila', 'Severidad', 'Columna', 'Valor', 'Motivo', 'Cómo corregirlo')
    assert len(filas) - 1 == incidencias


def test_cp045_filas_validas_en_csv_con_acentos(app, tmp_path):
    app.subir_archivo(DATOS / 'csv_utf8_con_bom.csv')
    archivo = app.descargar('btn-descargar-validas', tmp_path)
    assert archivo.suffix == '.csv'
    contenido = archivo.read_bytes()
    assert contenido.startswith(b'\xef\xbb\xbf'), 'sin BOM, Excel muestra mal los acentos'
    texto = contenido.decode('utf-8-sig')
    assert 'José' in texto and 'Núñez' in texto and '3311110000' in texto
    assert len(texto.strip().splitlines()) == 2  # encabezado + 1 fila válida


def _agregar_ventas(app, plantilla, tmp_path):
    app.limpiar()
    app.subir_plantilla(plantilla)
    expect(app.t('reglas-detectadas')).to_be_visible()
    app.t('obl-0').uncheck()  # la columna "NO" no viene en las ventas nuevas
    app.subir_archivo(EJEMPLOS / 'ventas_nuevas_demo.xlsx')
    return app.descargar('btn-descargar-base', tmp_path)


def test_cp046_agregar_a_archivo_acumulado(app, tmp_path):
    salida = _agregar_ventas(app, EJEMPLOS / 'base_ventas_demo.xlsx', tmp_path)
    assert salida.name == 'base_ventas_demo_actualizado.xlsx'
    wb = openpyxl.load_workbook(salida)
    assert wb.sheetnames == ['Ventas', 'Catalogos', 'Resumen', 'ValidaCarga Revisión', 'ValidaCarga Resumen']
    v = wb['Ventas']
    assert v['B401'].value == 'EJECUTIVO 401', 'la última fila original no debe cambiar'
    assert [v[f'B{n}'].value for n in (402, 403)] == ['NANCY GABRIELA BARBA PEREZ', 'MAYRA GONZALEZ ORTEGA']
    assert v['A402'].value == 'KREDIYA', 'se conserva la columna prellenada'
    assert v['F402'].value == 2410 and v['F402'].number_format == v['F401'].number_format
    assert v['I402'].is_date
    assert relleno(v['G403']) == ROJO and v['G403'].comment is not None   # "SUC JAL ATOTONILCO " con espacio
    assert relleno(v['G404']) == ROJO and relleno(v['B404']) == ROJO     # sucursal fuera de lista y ejecutivo vacío
    assert relleno(v['I404']) == AMBAR                                    # 31/09/2026
    assert wb['Resumen']['B1'].value == '=SUM(Ventas!F:F)'
    assert [str(dv.sqref) for dv in v.data_validations.dataValidation] == ['G2:G5000']


def test_cp047_agregar_dos_veces_sin_duplicar(app, tmp_path):
    primera = _agregar_ventas(app, EJEMPLOS / 'base_ventas_demo.xlsx', tmp_path)
    segunda = _agregar_ventas(app, primera, tmp_path)  # se llama ..._actualizado_actualizado.xlsx
    wb = openpyxl.load_workbook(segunda)
    v = wb['Ventas']
    assert [v[f'B{n}'].value for n in (405, 406)] == ['NANCY GABRIELA BARBA PEREZ', 'MAYRA GONZALEZ ORTEGA']
    assert wb.sheetnames.count('ValidaCarga Revisión') == 1
    with zipfile.ZipFile(segunda) as z:
        hoja = z.read('xl/worksheets/sheet1.xml').decode()
    assert hoja.count('<dataValidation ') == 1, 'la lista de sucursales no debe duplicarse'


def test_cp049_sin_conexion_para_excel_con_colores(app):
    app.page.route('**/ajax/libs/exceljs/**', lambda r: r.abort())
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    app.t('btn-descargar-base').click()
    expect(app.t('estado-descarga')).to_contain_text('No se pudo cargar el generador de Excel con colores')
    expect(app.t('btn-descargar-base')).to_be_enabled()


def test_cp050_archivo_marcado_en_ejemplos(app, tmp_path):
    app.modo('ejemplos')
    app.perfil('recargas')
    app.probar_ejemplo()
    archivo = app.descargar('btn-descargar-marcado', tmp_path)
    wb = openpyxl.load_workbook(archivo)
    assert wb.sheetnames == ['Datos', 'Resumen']
    masters = [c for fila in wb['Datos'].iter_rows(min_row=2) for c in fila if c.value == 'Master']
    assert masters and all(relleno(c) == ROJO and c.comment for c in masters)
