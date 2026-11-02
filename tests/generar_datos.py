"""Genera los archivos de datos de prueba de tests/datos/. Todos los datos son ficticios.
Uso: python tests/generar_datos.py"""
import datetime, pathlib, random, shutil, zipfile
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

D = pathlib.Path(__file__).parent / "datos"
D.mkdir(exist_ok=True)
EJ = pathlib.Path(__file__).parent.parent / "ejemplos"
ENC_DEMO = ['Usuario*', 'Nombre*', 'Apellido Paterno*', 'Correo*', 'Región*', 'Estado*', 'Tipo de Alumno*', 'Tipo de Canal*', 'Teléfono']
ENC_REC = ['Folio', 'Fecha', 'Número de Agente', 'Nombre de Agente', 'Número Destino', 'Monto', 'Movimientos en la Bolsa de', 'Estatus']
ENC_POR = ['Número', 'Tipo de Movimiento', 'Fecha de Solicitud', 'Fecha de Ejecución', 'Proveedor Donador', 'Proveedor Receptor']


def libro(encabezados, filas, ruta, titulo='Datos'):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = titulo
    ws.append(encabezados)
    for f in filas:
        ws.append(f)
    for fila in ws.iter_rows(min_row=2):          # todo como texto: el dato se ve tal cual
        for c in fila:
            if isinstance(c.value, str):
                c.number_format = '@'
    wb.save(D / ruta)


def fila_ok(usuario='DEMO0001', **cambios):
    base = dict(zip(ENC_DEMO, [usuario, 'Ana', 'Lugo', 'alugo@demo.mx', 'Región 5', 'Jalisco', 'Distribuidor', 'Detalle de Canal', '3311110000']))
    base.update({k.replace('_', ' '): v for k, v in cambios.items()})
    return [base[h] for h in ENC_DEMO]


# --- Plantillas especiales -------------------------------------------------------
demo = EJ / 'plantilla_oficial_demo.xlsx'
wb = openpyxl.load_workbook(demo); ins = wb.create_sheet('Instrucciones', 0)
ins['A1'] = 'Llena la pestaña Formato'; ins['A2'] = 'No cambies los encabezados'
wb.save(D / 'plantilla_instrucciones.xlsx')
wb = openpyxl.load_workbook(demo); wb.create_sheet('Hoja1', 0); wb.save(D / 'plantilla_hoja_vacia_primero.xlsx')
libro(['Sucursal', 'Vendedor', 'Monto', 'Fecha'], [], 'plantilla_sin_listas.xlsx')
openpyxl.Workbook().save(D / 'plantilla_vacia.xlsx')

wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Formato'
ws.append(['Sucursal*', 'Producto*'])
c = wb.create_sheet('Listas'); [c.cell(i, 1, v) for i, v in enumerate(['SUC A', 'SUC B'], 1)]
dv = DataValidation(type='list', formula1='OFFSET(Listas!$A$1,0,0,COUNTA(Listas!$A:$A),1)'); dv.add('A2:A500'); ws.add_data_validation(dv)
dv = DataValidation(type='custom', formula1='ISNUMBER(SEARCH("KIT",B2))'); dv.add('B2:B500'); ws.add_data_validation(dv)
wb.save(D / 'plantilla_reglas_no_reconocidas.xlsx')

# Plantilla con título arriba, lista moderna (x14) a otra hoja y largo exacto de IMEI
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Carga'
ws['A1'] = 'FORMATO DE ACTUALIZACIÓN DE VENTAS'; ws.append([]); ws.append(['Sucursal', 'Fecha de Venta', 'IMEI', 'Teléfono', 'Canal'])
c = wb.create_sheet('Listas'); [c.cell(i, 1, v) for i, v in enumerate(['SUC CENTRO', 'SUC PLAZA', 'SUC NORTE '], 1)]
wb.save(D / '_base_titulo.xlsx')
z = zipfile.ZipFile(D / '_base_titulo.xlsx'); f = {n: z.read(n) for n in z.namelist()}; z.close()
x = f['xl/worksheets/sheet1.xml'].decode()
x = x.replace('</sheetData>', '</sheetData><dataValidations count="2"><dataValidation type="textLength" operator="equal" sqref="C4:C1000"><formula1>15</formula1></dataValidation><dataValidation type="list" sqref="E4:E1000"><formula1>"Detalle,Mayoreo"</formula1></dataValidation></dataValidations>', 1)
x = x.replace('</worksheet>', '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Listas!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>A4:A1000</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>')
f['xl/worksheets/sheet1.xml'] = x.encode()
with zipfile.ZipFile(D / 'plantilla_titulo_x14.xlsx', 'w', zipfile.ZIP_DEFLATED) as z:
    for n, d in f.items():
        z.writestr(n, d)
(D / '_base_titulo.xlsx').unlink()

# --- Datos contra la plantilla demo -------------------------------------------
libro(ENC_DEMO, [fila_ok('ABC'), fila_ok('ABCD'), fila_ok('A' * 20), fila_ok('A' * 21)], 'limites_usuario.xlsx')
libro(ENC_DEMO, [fila_ok(f'TEL{n:02d}', Teléfono=t) for n, t in enumerate(['331111000', '3311110000', '33111100001', ''], 1)], 'limites_telefono.xlsx')
libro(ENC_DEMO, [fila_ok(f'REG{n:02d}', **{'Región*': r}) for n, r in enumerate(['Región 5', 'Region 5', 'región 5', 'Región 5 ', 'Región 7'], 1)], 'catalogo_variantes.xlsx')
libro(ENC_DEMO, [fila_ok(f'CAS{n:02d}', **{'Región*': r, 'Estado*': e}) for n, (r, e) in enumerate([
    ('Región 5', 'Jalisco'), ('Región 5', 'Ciudad de México'), ('Región 9', 'Estado de México'),
    ('Región 9', 'Jalisco'), ('Región 4', 'Michoacán'), ('Region 5', 'Jalisco')], 1)], 'cascada_region_estado.xlsx')
libro(ENC_DEMO, [fila_ok('DUP001'), fila_ok('DUP002'), fila_ok('DUP001'), fila_ok('dup001')], 'usuarios_duplicados.xlsx')
libro(ENC_DEMO, [fila_ok('OBL001', **{'Nombre*': ''})], 'obligatorio_vacio.xlsx')
wb = openpyxl.Workbook(); ws = wb.active; ws['A1'] = 'ALTAS NOVIEMBRE'; ws['A2'] = 'Generado por sistema'
ws.append([]); ws.append(ENC_DEMO); ws.append(fila_ok('TIT001')); ws.append(fila_ok('tit 002'))
wb.save(D / 'datos_con_titulo.xlsx')
libro(['Usuario*', 'Nombre*', 'Correo*'], [['FALT001', 'Ana', 'ana@demo.mx']], 'columnas_faltantes.xlsx')
openpyxl.Workbook().save(D / 'archivo_vacio.xlsx')
(D / 'documento.pdf').write_bytes(b'%PDF-1.4\n% archivo de prueba: no es Excel\n')

csv = 'Usuario*,Nombre*,Apellido Paterno*,Correo*,Región*,Estado*,Tipo de Alumno*,Tipo de Canal*,Teléfono\nCSV0001,José,Núñez,jnunez@demo.mx,Región 5,Michoacán,Distribuidor,Detalle de Canal,3311110000\n'
(D / 'csv_utf8_con_bom.csv').write_bytes('﻿'.encode('utf-8') + csv.encode('utf-8'))
(D / 'csv_utf8_sin_bom.csv').write_bytes(csv.encode('utf-8'))

# --- Perfiles de ejemplo --------------------------------------------------------
r = lambda folio, monto='100', bolsa='Agente', estatus='Exitosa', destino='3319876501', fecha='14/10/2026': [folio, fecha, '3310000001', 'SUC DEMO', destino, monto, bolsa, estatus]
libro(ENC_REC, [r('D1', bolsa='Agente', estatus='Exitosa', destino='3319876501'), r('D2', bolsa='Agente', estatus='Fallida', destino='3319876502'), r('D3', bolsa='Master', estatus='Exitosa', destino='3319876503'), r('D4', bolsa='Master', estatus='Fallida', destino='3319876504')], 'recargas_decision.xlsx')
libro(ENC_REC, [r(f'M{n}', monto=m) for n, m in enumerate(['0', '10', '75', '500', '501', '$1,000', '$1,00', 'cien'], 1)], 'recargas_montos.xlsx')
libro(ENC_REC, [r(f'T{n}', destino=d) for n, d in enumerate(['331987650', '3319876501', '33198765012', '523319876501', '33 1987 6501'], 1)], 'recargas_numeros.xlsx')
libro(ENC_POR, [[n, 'Portabilidad Entrante', '01/10/2026', '02/10/2026', 'Telcel', 'Wimo'] for n in
               ['3320001001', '3320001100-3320001199', '3320001100-99', '33200012XX', '3320001100 a 3320001109']], 'portabilidad_rangos.xlsx')
libro(ENC_POR, [['33200010' + f'{n:02d}', 'Portabilidad Entrante', s, e, 'Telcel', 'Wimo'] for n, (s, e) in enumerate([
    ('29/02/2024', '01/03/2024'), ('29/02/2026', '01/03/2026'), ('31/09/2026', '01/10/2026'), ('2026-10-15', '2026-10-16'),
    ('15-10-2026', '16/10/2026'), ('10/10/2026', '09/10/2026'), ('10/10/2026', '10/10/2026')], 1)], 'portabilidad_fechas.xlsx')

# --- Rendimiento ------------------------------------------------------------------
random.seed(7)
for n in (5000, 20000):
    wb = openpyxl.Workbook(write_only=True); ws = wb.create_sheet('Datos'); ws.append(ENC_DEMO)
    for i in range(n):
        ws.append(fila_ok(f'MAS{i:06d}', **({'Región*': 'Region 5'} if i % 50 == 0 else {})))
    wb.save(D / f'masivo_{n}.xlsx')
# .xls (formato antiguo) a partir del archivo demo, si LibreOffice está instalado
import subprocess
if shutil.which('soffice'):
    subprocess.run(['soffice', '--headless', '--convert-to', 'xls', '--outdir', str(D), str(EJ / 'datos_alta_usuarios_demo.xlsx')], capture_output=True, timeout=120)
print('Listo:', sorted(p.name for p in D.iterdir()))
