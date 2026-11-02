"""F07 Perfiles de ejemplo: las reglas que nacieron de los tres bugs reales."""
import pytest
from conftest import DATOS

C, E = 'correcto', 'error'

CASOS = [
    ('cp037_decision_bolsa_por_estatus', 'recargas', 'recargas_decision.xlsx', {2: C, 3: E, 4: E, 5: E},
     {3: 'Una recarga "Fallida" no se carga', 4: 'El ERP solo acepta "Agente"', 5: 'El ERP solo acepta "Agente"'}),
    ('cp038_montos', 'recargas', 'recargas_montos.xlsx', {2: E, 3: C, 4: E, 5: C, 6: E, 7: E, 8: E, 9: E},
     {2: 'no es una denominación', 4: 'no es una denominación', 6: 'no es una denominación', 7: 'no es una denominación',
      8: 'no es un monto válido', 9: 'no es un monto válido'}),
    ('cp039_numero_de_10_digitos', 'recargas', 'recargas_numeros.xlsx', {2: E, 3: C, 4: E, 5: E, 6: E},
     {2: 'tiene 9 dígitos', 4: 'tiene 11 dígitos', 5: 'Quitar el prefijo 52: 3319876501', 6: 'Dejar solo dígitos: 3319876501'}),
    ('cp040_rangos_por_bloque', 'portabilidad', 'portabilidad_rangos.xlsx', {2: C, 3: E, 4: E, 5: E, 6: E},
     {3: '(100 números)', 4: '(100 números)', 5: '(100 números)', 6: '(10 números)'}),
    ('cp041_fechas', 'portabilidad', 'portabilidad_fechas.xlsx', {2: C, 3: E, 4: E, 5: C, 6: E, 7: E, 8: C},
     {3: '"29/02/2026" no es una fecha válida', 4: '"31/09/2026" no es una fecha válida', 6: '"15-10-2026" no es una fecha válida',
      7: 'es anterior a la solicitud'}),
]


@pytest.mark.parametrize('perfil, archivo, esperado, textos', [c[1:] for c in CASOS], ids=[c[0] for c in CASOS])
def test_perfil(app, perfil, archivo, esperado, textos):
    app.modo('ejemplos')
    app.perfil(perfil)
    app.subir_archivo(DATOS / archivo)
    assert app.estados() == esperado
    for fila, texto in textos.items():
        assert texto in app.texto_fila(fila), f'Fila {fila}: falta "{texto}"'


def test_cp037_master_y_fallida_muestra_ambos_motivos(app):
    app.modo('ejemplos')
    app.perfil('recargas')
    app.subir_archivo(DATOS / 'recargas_decision.xlsx')
    texto = app.texto_fila(5)
    assert 'solo acepta "Agente"' in texto and '"Fallida" no se carga' in texto


def test_cp042_ejemplo_del_perfil_de_usuarios(app):
    app.modo('ejemplos')
    app.perfil('usuarios')
    app.probar_ejemplo()
    assert app.resumen() == {'total': 10, 'correcto': 2, 'advertencia': 1, 'error': 7}
