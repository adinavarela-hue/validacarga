"""F06 Validación fila por fila · F08 Resumen y filtros. Técnicas: valores límite, equivalencia y tabla de decisión."""
import pytest
from conftest import DATOS, EJEMPLOS

C, A, E = 'correcto', 'advertencia', 'error'

CASOS = [
    # id, archivo, estado esperado por fila de Excel, {fila: texto que debe aparecer}
    ('cp030_limites_usuario', 'limites_usuario.xlsx', {2: E, 3: C, 4: C, 5: E},
     {2: 'Tiene 3 caracteres; la plantilla pide entre 4 y 20 caracteres', 5: 'Tiene 21 caracteres'}),
    ('cp031_limites_telefono', 'limites_telefono.xlsx', {2: A, 3: C, 4: A, 5: C},
     {2: 'No parece un número de 10 dígitos', 4: 'No parece un número de 10 dígitos'}),
    ('cp032_variantes_de_catalogo', 'catalogo_variantes.xlsx', {2: C, 3: E, 4: E, 5: E, 6: E},
     {3: 'Se parece a "Región 5"', 4: 'Se parece a "Región 5"', 5: 'Se parece a "Región 5"', 6: 'no está en la lista'}),
    ('cp033_tabla_decision_cascada', 'cascada_region_estado.xlsx', {2: C, 3: E, 4: C, 5: E, 6: A, 7: E},
     {3: '"Ciudad de México" no está en la lista de la plantilla para Región "Región 5"',
      5: '"Jalisco" no está en la lista de la plantilla para Región "Región 9"',
      6: 'no trae la lista de Estado para Región "Región 4"'}),
    ('cp034_duplicados_sin_distinguir_mayusculas', 'usuarios_duplicados.xlsx', {2: C, 3: C, 4: E, 5: E},
     {4: 'ya aparece en la fila 2', 5: 'ya aparece en la fila 2'}),
    ('cp035_obligatorio_vacio', 'obligatorio_vacio.xlsx', {2: E}, {2: 'Campo obligatorio vacío'}),
]


@pytest.mark.parametrize('archivo, esperado, textos', [c[1:] for c in CASOS], ids=[c[0] for c in CASOS])
def test_validacion_por_fila(app, archivo, esperado, textos):
    app.subir_archivo(DATOS / archivo)
    assert app.estados() == esperado
    for fila, texto in textos.items():
        assert texto in app.texto_fila(fila), f'Fila {fila}: falta "{texto}"'


def test_cp035_muestra_valor_vacio(app):
    app.subir_archivo(DATOS / 'obligatorio_vacio.xlsx')
    assert '(vacío)' in app.texto_fila(2)


def test_cp032_fila_7_de_cascada_no_revisa_estado(app):
    """Si la Región es inválida, el Estado no se revisa (no hay contra qué lista compararlo)."""
    app.subir_archivo(DATOS / 'cascada_region_estado.xlsx')
    texto = app.texto_fila(7)
    assert 'Región*' in texto and 'Estado*' not in texto


def test_cp043_totales_y_filtros_cuadran(app):
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    r = app.resumen()
    assert r['correcto'] + r['advertencia'] + r['error'] == r['total']
    for estado in ('correcto', 'advertencia', 'error'):
        app.t(f'filtro-{estado}').click()
        assert app.page.locator('tbody.grupo').count() == r[estado], estado
    app.t('filtro-todas').click()
    assert app.page.locator('tbody.grupo').count() == r['total']
