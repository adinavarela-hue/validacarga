"""Construye docs/casos-de-prueba.xlsx a partir de tests/casos_de_prueba.py.
Uso: python tests/generar_casos_xlsx.py"""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from casos_de_prueba import CASOS, FUNCIONES, RIESGOS

SALIDA = pathlib.Path(__file__).parent.parent / 'docs' / 'casos-de-prueba.xlsx'
F = 'Arial'
PETROLEO, BLANCO, GRIS, LINEA = '155F78', 'FFFFFF', 'F2F5F6', 'D3DDE0'
AMARILLO = 'FFF4CC'
fuente = lambda **k: Font(name=F, size=k.pop('size', 10), **k)
borde = Border(*(Side(style='thin', color=LINEA),) * 4)
envolver = Alignment(wrap_text=True, vertical='top')
nombre_f = {f: n for f, n, _ in FUNCIONES}
ESTADOS = ['No ejecutado', 'Aprobado', 'Fallido', 'Bloqueado']

wb = Workbook()


def encabezado(ws, fila, textos, anchos=None):
    for i, t in enumerate(textos, 1):
        c = ws.cell(fila, i, t)
        c.font = fuente(bold=True, color=BLANCO)
        c.fill = PatternFill('solid', fgColor=PETROLEO)
        c.alignment = Alignment(wrap_text=True, vertical='center')
        c.border = borde
    if anchos:
        for i, w in enumerate(anchos, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w


def titulo(ws, texto, sub=None):
    ws['A1'] = texto
    ws['A1'].font = fuente(bold=True, size=14, color=PETROLEO)
    if sub:
        ws['A2'] = sub
        ws['A2'].font = fuente(italic=True, color='52636B')


def tabla(ws, fila0, cabeceras, filas, anchos=None):
    encabezado(ws, fila0, cabeceras, anchos)
    for r, valores in enumerate(filas, fila0 + 1):
        for c, v in enumerate(valores, 1):
            cel = ws.cell(r, c, v)
            cel.font = fuente()
            cel.alignment = envolver
            cel.border = borde
    return fila0 + len(filas)


# ------------------------------------------------------------------ Cómo usar
ws = wb.active
ws.title = 'Cómo usar'
titulo(ws, 'ValidaCarga v1.0 · Casos de prueba', 'Autora: Adina Varela · Plan de pruebas: docs/plan-de-pruebas.md · Demo: https://adinavarela-hue.github.io/validacarga/')
filas = [
    ('Pestaña', 'Para qué sirve'),
    ('Resumen', 'Avance de la ejecución por prioridad y estado. Se calcula solo.'),
    ('Casos', f'Los {len(CASOS)} casos de prueba. Al ejecutar, llena solo las columnas en amarillo.'),
    ('Valores límite', 'Valores elegidos en cada frontera y el resultado esperado.'),
    ('Clases de equivalencia', 'Grupos de entradas que la app debe tratar igual.'),
    ('Tablas de decisión', 'Combinaciones de reglas (TD-01 recargas, TD-02 cascada).'),
    ('Trazabilidad', 'Cuántos casos cubren cada función (F01–F13) y cada riesgo (R1–R8) del plan.'),
]
for i, (a, b) in enumerate(filas, 4):
    ws.cell(i, 1, a).font = fuente(bold=(i == 4))
    ws.cell(i, 2, b).font = fuente(bold=(i == 4))
ws['A12'] = 'Columnas que se llenan al ejecutar (en amarillo en "Casos")'
ws['A12'].font = fuente(bold=True)
leyenda = [('Estado', 'No ejecutado / Aprobado / Fallido / Bloqueado (lista desplegable)'),
           ('Resultado obtenido', 'Lo que pasó realmente, sobre todo si no coincide con lo esperado'),
           ('Bug (Jira)', 'Clave del bug creado, por ejemplo VAL-12'),
           ('Fecha de ejecución', 'DD/MM/AAAA')]
for i, (a, b) in enumerate(leyenda, 13):
    ws.cell(i, 1, a).fill = PatternFill('solid', fgColor=AMARILLO)
    ws.cell(i, 1).font = fuente()
    ws.cell(i, 2, b).font = fuente()
ws['A18'] = 'Ejemplo de una ejecución llenada'
ws['A18'].font = fuente(bold=True)
ej = [('ID', 'CP-029'), ('Estado', 'Fallido'), ('Resultado obtenido', 'La fila 2 aparece como "Correcta" aunque faltan 6 columnas obligatorias.'),
      ('Bug (Jira)', 'VAL-3'), ('Fecha de ejecución', '03/11/2026')]
for i, (a, b) in enumerate(ej, 19):
    ws.cell(i, 1, a).font = fuente()
    ws.cell(i, 2, b).font = fuente(italic=True, color='52636B')
ws.column_dimensions['A'].width = 26
ws.column_dimensions['B'].width = 100

# ------------------------------------------------------------------ Casos
wc = wb.create_sheet('Casos')
cab = ['ID', 'Función', 'Riesgo', 'Técnica', 'Prioridad', 'Título', 'Precondición', 'Pasos', 'Datos de prueba',
       'Resultado esperado', 'Automatizable', 'Estado', 'Resultado obtenido', 'Bug (Jira)', 'Fecha de ejecución']
anchos = [9, 24, 9, 20, 10, 34, 34, 42, 34, 60, 13, 14, 40, 12, 13]
encabezado(wc, 1, cab, anchos)
wc.row_dimensions[1].height = 30
for r, (cid, fn, rg, tec, pri, tit, pre, pasos, datos, esp, auto) in enumerate(CASOS, 2):
    valores = [cid, f'{fn} · {nombre_f[fn]}', rg, tec, pri, tit, pre, pasos, datos, esp, auto, 'No ejecutado', '', '', '']
    for c, v in enumerate(valores, 1):
        cel = wc.cell(r, c, v)
        cel.font = fuente(bold=(c == 1))
        cel.alignment = envolver
        cel.border = borde
        if c >= 12:
            cel.fill = PatternFill('solid', fgColor=AMARILLO)
    wc.cell(r, 15).number_format = 'dd/mm/yyyy'
ultima = len(CASOS) + 1
dv = DataValidation(type='list', formula1='"' + ','.join(ESTADOS) + '"', allow_blank=False)
dv.add(f'L2:L{ultima}')
wc.add_data_validation(dv)
colores = {'Aprobado': ('E3F2EA', '1F7A4D'), 'Fallido': ('FAE5E3', 'B4322B'), 'Bloqueado': ('FBF0DA', '935D00')}
for est, (fondo, texto) in colores.items():
    wc.conditional_formatting.add(f'L2:L{ultima}', CellIsRule(operator='equal', formula=[f'"{est}"'],
                                  fill=PatternFill('solid', fgColor=fondo), font=Font(name=F, bold=True, color=texto)))
wc.conditional_formatting.add(f'E2:E{ultima}', CellIsRule(operator='equal', formula=['"Alta"'], font=Font(name=F, bold=True, color='B4322B')))
wc.freeze_panes = 'C2'
wc.page_setup.orientation = 'landscape'
wc.page_setup.fitToWidth = 1
wc.page_setup.fitToHeight = 0
wc.sheet_properties.pageSetUpPr.fitToPage = True
wc.print_title_rows = '1:1'
wc.auto_filter.ref = f'A1:O{ultima}'

# ------------------------------------------------------------------ Valores límite
wv = wb.create_sheet('Valores límite')
titulo(wv, 'Valores límite', 'Se prueba justo antes, en y justo después de cada frontera.')
filas = [
    ('VL-01', 'Usuario* (longitud)', 'Entre 4 y 20 caracteres', '3', 'Error', 'CP-030'),
    ('VL-02', 'Usuario* (longitud)', 'Entre 4 y 20 caracteres', '4', 'Correcta', 'CP-030'),
    ('VL-03', 'Usuario* (longitud)', 'Entre 4 y 20 caracteres', '20', 'Correcta', 'CP-030'),
    ('VL-04', 'Usuario* (longitud)', 'Entre 4 y 20 caracteres', '21', 'Error', 'CP-030'),
    ('VL-05', 'Teléfono', '10 dígitos (deducida por el nombre)', '9 dígitos', 'Advertencia', 'CP-031'),
    ('VL-06', 'Teléfono', '10 dígitos (deducida por el nombre)', '10 dígitos', 'Correcta', 'CP-031'),
    ('VL-07', 'Teléfono', '10 dígitos (deducida por el nombre)', '11 dígitos', 'Advertencia', 'CP-031'),
    ('VL-08', 'Número Destino (recargas)', 'Exactamente 10 dígitos', '9 / 10 / 11 dígitos', 'Error / Correcta / Error', 'CP-039'),
    ('VL-09', 'Monto (recargas)', 'Denominaciones $10 a $500', '0 / 10 / 500 / 501', 'Error / Correcta / Correcta / Error', 'CP-038'),
    ('VL-10', 'Fecha de Solicitud', 'Fecha que exista en el calendario', '29/02/2024 (bisiesto)', 'Correcta', 'CP-041'),
    ('VL-11', 'Fecha de Solicitud', 'Fecha que exista en el calendario', '29/02/2026 / 31/09/2026', 'Error', 'CP-041'),
    ('VL-12', 'Fecha de Ejecución', 'Igual o posterior a la solicitud', 'Un día antes / el mismo día', 'Error / Correcta', 'CP-041'),
    ('VL-13', 'Revisión de muestra', 'Hasta 5 filas = muestra; más = archivo acumulado', '400 filas', 'Archivo acumulado', 'CP-019'),
    ('VL-14', 'Tabla en pantalla', 'Muestra hasta 500 filas', '5,000 y 20,000 filas', '500 filas + nota', 'CP-056'),
]
tabla(wv, 4, ['ID', 'Campo', 'Regla', 'Valor probado', 'Resultado esperado', 'Caso'], filas, [8, 26, 38, 26, 30, 9])

# ------------------------------------------------------------------ Clases de equivalencia
we = wb.create_sheet('Clases de equivalencia')
titulo(we, 'Clases de equivalencia', 'Un valor representativo por clase; si uno pasa, la clase pasa.')
filas = [
    ('CE-01', 'Formato de la plantilla', 'Válida', '.xlsx', 'Se leen reglas', 'CP-001'),
    ('CE-02', 'Formato de la plantilla', 'Inválida', '.csv u otro', 'Mensaje "debe ser .xlsx"', 'CP-002'),
    ('CE-03', 'Formato del archivo de datos', 'Válida', '.xlsx / .xls / .csv con y sin BOM', 'Se valida', 'CP-020 a CP-023'),
    ('CE-04', 'Formato del archivo de datos', 'Inválida', '.pdf', 'Mensaje "no es un Excel ni un CSV"', 'CP-024'),
    ('CE-05', 'Valor de catálogo', 'Válida', 'Idéntico a la lista', 'Correcta', 'CP-032'),
    ('CE-06', 'Valor de catálogo', 'Inválida: parecido', 'Sin acento, minúsculas, espacio final', 'Error con sugerencia exacta', 'CP-032'),
    ('CE-07', 'Valor de catálogo', 'Inválida: ajeno', 'No existe en la lista', 'Error "no está en la lista"', 'CP-032'),
    ('CE-08', 'Pestaña con el formato', 'Primera pestaña', 'Formato en la 1', 'Se elige la 1', 'CP-001'),
    ('CE-09', 'Pestaña con el formato', 'No es la primera', 'Instrucciones o vacía primero', 'Se elige la del formato', 'CP-004, CP-005'),
    ('CE-10', 'Fila de encabezados', 'Fila 1', 'Encabezados arriba', 'Filas numeradas desde 2', 'CP-020'),
    ('CE-11', 'Fila de encabezados', 'Con títulos arriba', 'Encabezados en la fila 3 o 4', 'Se ubican y la numeración coincide con Excel', 'CP-007, CP-026'),
    ('CE-12', 'Número telefónico', 'Válida', '10 dígitos', 'Correcta', 'CP-039'),
    ('CE-13', 'Número telefónico', 'Inválida: con lada 52', '12 dígitos', 'Error "Quitar el prefijo 52"', 'CP-039'),
    ('CE-14', 'Número telefónico', 'Inválida: con separadores', 'Espacios o guiones', 'Error "Dejar solo dígitos"', 'CP-039'),
    ('CE-15', 'Número de portabilidad', 'Individual', '10 dígitos', 'Correcta', 'CP-040'),
    ('CE-16', 'Número de portabilidad', 'Rango por bloque', 'Guion, "a", abreviado o XX', 'Error con número de líneas', 'CP-040'),
    ('CE-17', 'Monto', 'Inválida: formato', '"$1,00", "cien"', 'Error "no es un monto válido"', 'CP-038'),
]
tabla(we, 4, ['ID', 'Entrada', 'Clase', 'Representante', 'Resultado esperado', 'Caso'], filas, [8, 28, 22, 34, 40, 15])

# ------------------------------------------------------------------ Tablas de decisión
wt = wb.create_sheet('Tablas de decisión')
titulo(wt, 'Tablas de decisión')
wt['A3'] = 'TD-01 · Consolidado de recargas: Bolsa × Estatus (bug real 2) → CP-037'
wt['A3'].font = fuente(bold=True)
fin = tabla(wt, 4, ['Regla', 'Movimientos en la Bolsa de', 'Estatus', '¿Se carga?', 'Mensaje esperado'], [
    ('R1', 'Agente', 'Exitosa', 'Sí · Correcta', '—'),
    ('R2', 'Agente', 'Fallida', 'No · Error', 'Una recarga "Fallida" no se carga al consolidado'),
    ('R3', 'Master', 'Exitosa', 'No · Error', 'El ERP solo acepta "Agente"… se descartan en silencio'),
    ('R4', 'Master', 'Fallida', 'No · Error', 'Ambos mensajes'),
], [10, 28, 24, 30, 60])
wt.cell(fin + 2, 1, 'TD-02 · Cascada Región × Estado (bug real 1) → CP-033').font = fuente(bold=True)
tabla(wt, fin + 3, ['Fila', 'Región', 'Estado', 'Resultado esperado', 'Motivo'], [
    ('2', 'Región 5', 'Jalisco', 'Correcta', 'Estado dentro de la lista de Región 5'),
    ('3', 'Región 5', 'Ciudad de México', 'Error en Estado', 'No está en la lista de Región 5'),
    ('4', 'Región 9', 'Estado de México', 'Correcta', 'Estado dentro de la lista de Región 9'),
    ('5', 'Región 9', 'Jalisco', 'Error en Estado', 'No está en la lista de Región 9'),
    ('6', 'Región 4', 'Michoacán', 'Advertencia', 'La plantilla no trae lista de Estado para Región 4'),
    ('7', 'Region 5 (sin acento)', 'Jalisco', 'Error en Región', 'Región inválida: el Estado no se revisa'),
])

# ------------------------------------------------------------------ Trazabilidad
wz = wb.create_sheet('Trazabilidad')
titulo(wz, 'Trazabilidad con el plan de pruebas', 'Los conteos se calculan desde la pestaña "Casos".')
n = len(CASOS) + 1
encabezado(wz, 4, ['Función', 'Nombre', 'Prioridad en el plan', 'Casos', 'Casos de prioridad alta', 'Aprobados', 'Fallidos'], [10, 30, 18, 10, 20, 12, 12])
for i, (f, nom, pri) in enumerate(FUNCIONES, 5):
    vals = [f, nom, pri,
            f'=COUNTIF(Casos!$B$2:$B${n},A{i}&"*")',
            f'=COUNTIFS(Casos!$B$2:$B${n},A{i}&"*",Casos!$E$2:$E${n},"Alta")',
            f'=COUNTIFS(Casos!$B$2:$B${n},A{i}&"*",Casos!$L$2:$L${n},"Aprobado")',
            f'=COUNTIFS(Casos!$B$2:$B${n},A{i}&"*",Casos!$L$2:$L${n},"Fallido")']
    for c, v in enumerate(vals, 1):
        cel = wz.cell(i, c, v); cel.font = fuente(); cel.border = borde
fila_r = 5 + len(FUNCIONES) + 2
encabezado(wz, fila_r, ['Riesgo', 'Descripción', '', 'Casos'])
for i, (r, d) in enumerate(RIESGOS, fila_r + 1):
    for c, v in enumerate([r, d, '', f'=COUNTIF(Casos!$C$2:$C${n},"*"&A{i}&"*")'], 1):
        cel = wz.cell(i, c, v); cel.font = fuente(); cel.border = borde
zero = 'D5:D{}'.format(4 + len(FUNCIONES))
wz.conditional_formatting.add(zero, CellIsRule(operator='equal', formula=['0'], fill=PatternFill('solid', fgColor='FAE5E3')))

# ------------------------------------------------------------------ Resumen
wr = wb.create_sheet('Resumen', 1)
titulo(wr, 'Resumen de la ejecución', 'Se actualiza al cambiar la columna "Estado" en la pestaña Casos.')
encabezado(wr, 4, ['Prioridad'] + ESTADOS + ['Total', 'Avance'], [14, 14, 12, 12, 12, 10, 12])
prioridades = ['Alta', 'Media', 'Baja']
for i, p in enumerate(prioridades, 5):
    wr.cell(i, 1, p)
    for j, est in enumerate(ESTADOS, 2):
        wr.cell(i, j, f'=COUNTIFS(Casos!$E$2:$E${n},$A{i},Casos!$L$2:$L${n},{wr.cell(4, j).coordinate})')
    wr.cell(i, 6, f'=COUNTIF(Casos!$E$2:$E${n},$A{i})')
    wr.cell(i, 7, f'=IF(F{i}=0,0,(F{i}-B{i})/F{i})')
t = 5 + len(prioridades)
wr.cell(t, 1, 'Total')
for j in range(2, 7):
    col = wr.cell(4, j).column_letter
    wr.cell(t, j, f'=SUM({col}5:{col}{t - 1})')
wr.cell(t, 7, f'=IF(F{t}=0,0,(F{t}-B{t})/F{t})')
for r in range(5, t + 1):
    for c in range(1, 8):
        cel = wr.cell(r, c); cel.font = fuente(bold=(r == t or c == 1)); cel.border = borde
    wr.cell(r, 7).number_format = '0%'
wr.cell(t + 2, 1, 'Casos automatizables').font = fuente(bold=True)
wr.cell(t + 2, 2, f'=COUNTIF(Casos!$K$2:$K${n},"Sí")').font = fuente()
wr.cell(t + 3, 1, 'Criterio de salida').font = fuente(bold=True)
wr.cell(t + 3, 2, f'=IF(AND(F5>0,C5=F5),"Cumplido: 100% de prioridad alta aprobado","Pendiente: faltan casos de prioridad alta por aprobar")').font = fuente()
wr.column_dimensions['B'].width = 14

for hoja in wb.worksheets:
    hoja.sheet_view.showGridLines = hoja.title in ('Casos',)
SALIDA.parent.mkdir(exist_ok=True)
wb.save(SALIDA)
print('Guardado', SALIDA, len(CASOS), 'casos')
