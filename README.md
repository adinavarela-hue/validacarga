# ValidaCarga

[![Pruebas](https://github.com/adinavarela-hue/validacarga/actions/workflows/pruebas.yml/badge.svg)](https://github.com/adinavarela-hue/validacarga/actions/workflows/pruebas.yml)

Validador de archivos de carga masiva (Excel / CSV) que revisa cada fila **antes** de subirla a un portal y explica qué falla, en qué columna, por qué y cómo corregirlo.

Proyecto de portafolio QA de **Adina Varela**. Nació de tres bugs reales que encontré en mi trabajo en telecomunicaciones:

| Perfil | Bug real que lo originó |
|---|---|
| Alta masiva de usuarios | La plantilla oficial de un portal de capacitación fallaba en los campos de catálogo (Región, Estado, Ciudad, Municipio, Tipo de Alumno, Tipo de Canal), incluso con su propia fila de muestra. |
| Consolidado de recargas | Un ERP descartaba en silencio las recargas con "Movimientos en la Bolsa de" = "Master": 60 recargas nunca entraron y no hubo ningún mensaje de error. |
| Reporte de portabilidad | El formato regulatorio exigía una línea por número individual y el archivo traía rangos por bloque. |

> Los perfiles, catálogos y datos de ejemplo son ficticios.

## Cómo se usa

**Modo principal: plantilla oficial** (sirve para cualquier portal)

1. Sube la plantilla oficial en blanco que te da el portal (`.xlsx`). ValidaCarga lee de ella:
   - las columnas y cuáles son obligatorias (marcadas con `*`; se pueden ajustar),
   - las listas desplegables (escritas en la validación, de un rango, de otra hoja o de un nombre definido),
   - las listas dependientes en cascada (`INDIRECTO` / `INDIRECT`),
   - reglas de número, longitud de texto, fecha y valores sin repetir.
2. Revisa los hallazgos sobre la propia plantilla: por ejemplo, si su fila de muestra no pasa sus propias listas.
3. Sube tu archivo `.xlsx`, `.xls` o `.csv`. Puede traer filas de título arriba; ValidaCarga ubica los encabezados.
4. Revisa el resultado fila por fila (**correcta**, **advertencia** o **error**) y descarga:
   - el **reporte de errores**,
   - el **archivo solo con filas válidas**,
   - **Agregar a mi archivo**: tu plantilla (o tu archivo acumulado, aunque ya tenga cientos de filas) con las filas nuevas escritas debajo de la última, solo en las columnas compatibles y con el formato de la fila de arriba. Lo que pasó queda normal; lo que falló, en rojo con una nota del motivo; las advertencias, en ámbar. Se conservan las demás pestañas, fórmulas y listas desplegables, y se agregan las pestañas "ValidaCarga Revisión" y "ValidaCarga Resumen". Lo genera ExcelJS, que se carga solo al presionar el botón.
   - En **Ejemplos**, el **archivo marcado**: el formato del perfil con los errores en rojo y notas.

**Ejemplos:** tres perfiles programados a partir de los bugs reales, con reglas que una plantilla no puede expresar (como "Master" vs. "Agente").

Todo corre en el navegador: el archivo no sale de tu computadora.

## Estructura

```
index.html          Página
css/estilos.css     Estilos
js/validador.js     Motor de validación (sin dependencias; corre en navegador y en Node)
js/perfiles.js      Perfiles de ejemplo: columnas, catálogos, reglas y archivos de muestra
js/plantilla.js     Lector de plantillas oficiales: deduce reglas de las validaciones de Excel
js/demo-plantilla.js  Plantilla de demostración embebida
js/app.js           Interfaz: lectura de archivos (SheetJS), tabla y descargas
ejemplos/           Plantilla oficial de demostración y archivo de datos para probar
build.py            Genera dist/ValidaCarga.html, la app en un solo archivo
```

Para abrirla sin servidor: `python build.py` y abre `dist/ValidaCarga.html` con doble clic.

## Demo

https://adinavarela-hue.github.io/validacarga/

## Pruebas

- [Plan de pruebas](docs/plan-de-pruebas.md) — alcance, riesgos, técnicas, criterios y gestión de defectos.
- [Casos de prueba](docs/casos-de-prueba.xlsx) — 57 casos con valores límite, clases de equivalencia, tablas de decisión y trazabilidad a funciones y riesgos. Se generan desde `tests/casos_de_prueba.py` con `python tests/generar_casos_xlsx.py`.
- Datos de prueba en `tests/datos/` (se regeneran con `python tests/generar_datos.py`).
- [Pruebas automatizadas](tests/README.md) con Playwright + pytest: 55 casos automatizados, Page Object, revisión de los Excel descargados y de la red. Corren solas en GitHub Actions en cada cambio.

## Bugs encontrados

Registrados en Jira (espacio **ValidaCarga**, clave VAL) con pasos, resultado esperado, resultado obtenido, severidad y el caso de prueba que los detectó.

| Bug | Resumen | Severidad | Caso | Estado |
|---|---|---|---|---|
| VAL-1 | Fila con columnas obligatorias faltantes aparece como "Correcta" | Alta | CP-029 | Corregido y verificado (lo detectó la prueba `xfail` estricta) |
| VAL-2 | Error de JavaScript al pasar a "Ejemplos" con una plantilla cargada | Media | CP-057 | Corregido y verificado |
| VAL-3 | "Agregar a mi archivo" duplica las listas desplegables y pierde las de formato moderno | Alta | CP-046, CP-047 | Corregido y verificado |

**Última corrida (02/11/2026):** 55 de 57 casos aprobados y 2 manuales pendientes; suite automatizada 60 de 60. Primera corrida: 54 aprobados y 1 fallido (VAL-1). Resultados por caso en [`docs/casos-de-prueba.xlsx`](docs/casos-de-prueba.xlsx).
