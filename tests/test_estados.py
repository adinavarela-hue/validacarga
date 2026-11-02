"""F03 Ajuste de obligatorios · F10 Limpiar · F11 Tema · regresión de cambio de modo (transición de estados)."""
from playwright.sync_api import expect
from conftest import DATOS, EJEMPLOS


def test_cp016_desmarcar_y_marcar_obligatoria(app):
    app.subir_archivo(DATOS / 'obligatorio_vacio.xlsx')
    assert app.estado_fila(2) == 'error'
    app.t('obl-1').uncheck()           # Nombre*
    expect(app.t('fila-2')).to_have_attribute('data-estado', 'correcto')
    app.t('obl-1').check()
    expect(app.t('fila-2')).to_have_attribute('data-estado', 'error')


def test_cp051_limpiar_borra_todo(app):
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    app.t('obl-1').uncheck()
    app.limpiar()
    expect(app.t('reglas-detectadas')).to_have_count(0)
    expect(app.t('tabla-resultados')).to_have_count(0)
    expect(app.t('input-archivo')).to_be_disabled()
    expect(app.t('btn-limpiar')).to_be_disabled()


def test_cp052_volver_a_subir_lo_mismo_despues_de_limpiar(app):
    app.t('obl-1').uncheck()
    app.limpiar()
    app.subir_plantilla(EJEMPLOS / 'plantilla_oficial_demo.xlsx')
    app.subir_archivo(EJEMPLOS / 'datos_alta_usuarios_demo.xlsx')
    assert app.resumen() == {'total': 9, 'correcto': 3, 'advertencia': 0, 'error': 6}
    expect(app.t('obl-1')).to_be_checked()


def test_cp053_tema_claro_se_recuerda(app):
    app.t('btn-tema').click()
    expect(app.page.locator('html')).to_have_attribute('data-theme', 'light')
    app.page.reload()
    expect(app.page.locator('html')).to_have_attribute('data-theme', 'light')


def test_cp057_regresion_cambiar_de_modo_con_plantilla_cargada(app):
    """Bug encontrado al construir: el panel de reglas fallaba al pasar a Ejemplos con una plantilla cargada.
    El fixture `app` falla la prueba si aparece cualquier error de JavaScript."""
    for _ in range(2):
        app.modo('ejemplos')
        app.modo('plantilla')
    assert len(app.reglas()) == 9
