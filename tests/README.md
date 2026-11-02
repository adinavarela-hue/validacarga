# Pruebas automatizadas de ValidaCarga

Suite de **Playwright + pytest** que cubre 55 de los 57 casos de [`docs/casos-de-prueba.xlsx`](../docs/casos-de-prueba.xlsx)
(CP-036 y CP-048 son manuales). Cada prueba lleva en su nombre el ID del caso, por ejemplo `test_cp030_...`.

## Estructura

```
tests/
  conftest.py            Servidor local, página abierta (fixture `app`) y falla si hay errores de JavaScript
  paginas/validacarga.py Page Object: cómo se usa la página (subir, leer resumen, descargar…)
  test_plantilla.py      F01, F02, F04 · plantilla oficial y reglas detectadas
  test_archivo.py        F05 · formatos, títulos, pestañas, columnas faltantes
  test_validacion.py     F06, F08 · valores límite, equivalencia, tabla de decisión, filtros
  test_perfiles.py       F07 · reglas de los 3 bugs reales
  test_descargas.py      F09 · abre cada Excel descargado y revisa celdas, colores y pestañas
  test_estados.py        F03, F10, F11 · obligatorios, limpiar, tema, regresión de cambio de modo
  test_no_funcionales.py F12, F13, R8 · privacidad (red), vista de celular, rendimiento
  casos_de_prueba.py     Fuente de los casos → `python tests/generar_casos_xlsx.py`
  generar_datos.py       Crea los archivos de tests/datos/
```

## Cómo correrla en Windows

```powershell
py -3.12 -m pip install -r requirements-dev.txt
py -3.12 -m playwright install chromium
py -3.12 -m pytest --html=reporte/reporte.html --self-contained-html
```

- Solo las rápidas: `pytest -m "not lento"` · Un caso: `pytest -k cp046`
- Ver el navegador mientras corre: `pytest --headed --slowmo 300 -k cp037`
- Contra la demo publicada: `set VC_URL=https://adinavarela-hue.github.io/validacarga/` y luego `pytest`
- Usar tu Chrome en lugar del Chromium de Playwright: `pytest --browser-channel chrome`

## Bugs conocidos

Un bug abierto se documenta con `xfail(strict=True)`: la suite sigue en verde y, cuando el bug se corrige,
la prueba falla a propósito para avisar que hay que quitar la marca.

Así se manejó **VAL-1** (fila incompleta aparecía como "Correcta"): se registró con `xfail`, se corrigió en
`js/validador.js` y la prueba `test_cp029_fila_incompleta_no_aparece_como_correcta` quedó como regresión normal.
No hay bugs abiertos.

## Resultado de referencia

02/11/2026 · Chrome · 60 de 60 aprobadas · ~45 s · dos corridas seguidas en verde.
