"""Registra en docs/casos-de-prueba.xlsx el resultado de una corrida (columnas Estado, Resultado obtenido, Bug, Fecha).
Uso: python tests/registrar_resultados.py"""
import datetime
import pathlib
import openpyxl

ARCHIVO = pathlib.Path(__file__).parent.parent / 'docs' / 'casos-de-prueba.xlsx'
FECHA = datetime.date(2026, 11, 2)
CORRIDA = 'Prueba automatizada aprobada: Windows 11 · Chrome y GitHub Actions (59 aprobadas, 1 fallo esperado).'

# Casos con algo distinto a "Aprobado sin bug"
ESPECIALES = {
    'CP-029': ('Fallido', 'Aviso de columnas faltantes y descarga deshabilitada: OK. Pero la fila 2 aparece como "Correcta" y el resumen cuenta 1 correcta.', 'VAL-1'),
    'CP-046': ('Aprobado', CORRIDA + ' Antes de la corrección duplicaba la lista de sucursales.', 'VAL-3'),
    'CP-047': ('Aprobado', CORRIDA + ' Una sola validación G2:G5000.', 'VAL-3'),
    'CP-057': ('Aprobado', CORRIDA + ' Regresión del error de JavaScript al cambiar de modo.', 'VAL-2'),
}
MANUALES = {'CP-036', 'CP-048'}  # se ejecutan a mano; quedan "No ejecutado" hasta hacerlo

wb = openpyxl.load_workbook(ARCHIVO)
ws = wb['Casos']
for fila in ws.iter_rows(min_row=2):
    cid = fila[0].value
    if not cid or cid in MANUALES:
        continue
    estado, obtenido, bug = ESPECIALES.get(cid, ('Aprobado', CORRIDA, ''))
    fila[11].value, fila[12].value, fila[13].value, fila[14].value = estado, obtenido, bug, FECHA
wb.save(ARCHIVO)
print('Resultados registrados en', ARCHIVO)
