# ValidaCarga · Plan de pruebas

01/11/2026 · Adina Varela

## 1. Identificación

| Dato | Valor |
| --- | --- |
| Producto | ValidaCarga v1.0 — validador web de archivos Excel/CSV antes de cargarlos a un portal |
| Demo | [adinavarela-hue.github.io/validacarga](https://adinavarela-hue.github.io/validacarga/) |
| Repositorio | [github.com/adinavarela-hue/validacarga](https://github.com/adinavarela-hue/validacarga) |
| Responsable de QA | Adina Varela |
| Tipo de aplicación | Página estática (HTML/JS) sin servidor; el archivo se procesa en el navegador |
| Origen | Tres bugs reales: catálogos de una plantilla de carga masiva, recargas "Master" descartadas en silencio por un ERP y rangos por bloque en un reporte de portabilidad |

## 2. Objetivo

Confirmar que ValidaCarga clasifica cada fila correctamente (correcta, advertencia o error), explica el motivo y genera descargas que no dañan ni alteran el archivo del usuario. El riesgo principal es un **falso negativo**: dejar pasar una fila que el portal rechazará o, peor, descartará en silencio, como pasó con las 60 recargas "Master".

Criterio de éxito: 100% de los casos de prioridad alta aprobados, cero defectos críticos o altos abiertos y la suite automatizada en verde.

## 3. Alcance

Se prueban 13 funciones; la prioridad sale del riesgo de la sección 4.

| ID | Función | Qué se verifica | Prioridad |
| --- | --- | --- | --- |
| F01 | Carga de plantilla oficial | Acepta .xlsx, rechaza otros formatos, ubica encabezados (fila 1 o con título), elige la pestaña correcta y permite cambiarla | Alta |
| F02 | Detección de reglas | Listas (escritas, rango, otra hoja, nombre definido, formato x14), cascada INDIRECTO, número, longitud, fecha, únicos, obligatorios con \* | Alta |
| F03 | Ajuste de obligatorios | Las casillas cambian la validación al instante | Media |
| F04 | Revisión de la plantilla | Fila de muestra que no pasa sus listas, opciones con espacios, listas dependientes faltantes, archivo acumulado | Alta |
| F05 | Carga del archivo de datos | .xlsx, .xls, .csv (UTF-8 con y sin BOM), filas de título, varias pestañas, archivo vacío | Alta |
| F06 | Validación fila por fila | Estado, columna, valor, motivo y corrección correctos; número de fila igual al de Excel | Alta |
| F07 | Perfiles de ejemplo | Reglas de los 3 bugs: catálogos en cascada, "Agente" vs. "Master", rangos por bloque | Alta |
| F08 | Resumen y filtros | Totales cuadran con la tabla; filtros por estado | Media |
| F09 | Descargas | Reporte de errores, filas válidas, "Agregar a mi archivo" y archivo marcado | Alta |
| F10 | Limpiar y empezar de nuevo | Borra plantilla, reglas, archivo y resultados; permite volver a subir el mismo archivo | Media |
| F11 | Tema oscuro/claro | Cambia y se recuerda al recargar | Baja |
| F12 | Privacidad | El archivo no se envía a ningún servidor (solo se cargan librerías de CDN) | Alta |
| F13 | Diseño adaptable | Sin desbordes a 390 px de ancho | Baja |

**Fuera de alcance:** la carga real a los portales destino; rendimiento con más de 20,000 filas; plantillas .xls; pruebas de seguridad más allá de F12; accesibilidad completa (solo foco visible y etiquetas); el funcionamiento interno de SheetJS y ExcelJS.

## 4. Riesgos del producto

Los tres riesgos altos guían el orden de ejecución: se prueban primero.

| ID | Riesgo | Prob. | Impacto | Mitigación en pruebas |
| --- | --- | --- | --- | --- |
| R1 | Falso negativo: una fila con error sale como correcta (catálogo con acento, espacio o mayúscula distinta) | Media | Alto | Clases de equivalencia sobre cada catálogo; casos de los 3 bugs reales |
| R2 | "Agregar a mi archivo" daña la base acumulada (pierde pestañas, fórmulas, listas o formato) | Media | Alto | Abrir la salida en Excel y LibreOffice; comparar pestañas, fórmulas y validaciones antes y después |
| R3 | El archivo se envía fuera de la computadora | Baja | Alto | Prueba automatizada que registra las peticiones de red durante la carga |
| R4 | Pérdida de datos al leer o descargar: ceros a la izquierda, números de 15 dígitos, fechas | Media | Medio | Valores límite en IMEI, teléfono y fechas; revisión de tipos en la salida |
| R5 | Reglas de la plantilla no reconocidas (fórmulas, OFFSET, validaciones personalizadas) | Alta | Medio | Plantillas de prueba por tipo de regla; la app debe avisar "Regla no reconocida" |
| R6 | Elige mal la pestaña o la fila de encabezados | Media | Medio | Archivos con instrucciones primero, pestañas vacías y títulos arriba |
| R7 | Sin conexión al CDN, la app no carga o no descarga | Baja | Medio | Prueba con red bloqueada: mensaje claro y el resto funciona |
| R8 | Archivos grandes vuelven lenta la página | Media | Bajo | Medición con 5,000 y 20,000 filas |

## 5. Enfoque y técnicas

Pruebas basadas en riesgo: manuales para diseñar y explorar, automatizadas para la regresión de lo que ya pasó.

| Técnica | Dónde se aplica | Ejemplo |
| --- | --- | --- |
| Valores límite | Longitudes y rangos | Teléfono de 9, 10 y 11 dígitos; usuario de 3, 4, 20 y 21 caracteres; edad 17, 18, 99, 100 |
| Clases de equivalencia | Formatos y catálogos | Archivo .xlsx / .xls / .csv / .pdf; valor exacto / con acento distinto / con espacio / fuera de lista |
| Tablas de decisión | Reglas combinadas | Región × Estado (cascada); Bolsa "Agente"/"Master" × Estatus "Exitosa"/"Fallida" |
| Transición de estados | Flujo de la pantalla | Sin plantilla → con plantilla → con archivo → limpiar; cambio de modo y de pestaña |
| Pruebas negativas | Entradas inválidas | Plantilla vacía, CSV como plantilla, pestaña sin encabezados, fecha 31/09 |
| Exploratorias | Por sesiones de 45 min | Archivos reales anonimizados de ventas y altas |
| Regresión automatizada | Playwright + pytest | Flujo principal, 3 perfiles, descargas (se abre el Excel y se revisan celdas y colores), red sin envíos |
| Compatibilidad y visual | Navegadores y anchos | Chrome, Edge, Firefox; 1280 px y 390 px; tema oscuro y claro |

Los casos de prueba detallados van en un archivo aparte (Fase 3), con ID, precondición, pasos, datos, resultado esperado y trazabilidad a F01–F13 y R1–R8.

## 6. Ambiente, herramientas y datos

| Elemento | Detalle |
| --- | --- |
| Sistema | Windows 11 (laptop de pruebas); Android para la vista de celular |
| Navegadores | Chrome, Edge y Firefox, versión estable vigente |
| Versiones de la app | Demo publicada en GitHub Pages y versión de un solo archivo (dist/ValidaCarga.html) |
| Automatización | Python 3.12, pytest, Playwright (Chromium) y openpyxl para leer los Excel descargados |
| Verificación de salidas | Microsoft Excel y LibreOffice Calc |
| Gestión | Jira (casos y bugs) y GitHub (código, pruebas y README) |

**Datos de prueba** (todos ficticios, en la carpeta ejemplos/ del repositorio):

- Plantilla oficial demo: listas, cascada Región → Estado, formato x14, fila de muestra con error.
- Datos de alta de usuarios con 6 errores sembrados.
- Archivo con 3 pestañas y archivo acumulado de 400 filas (ventas) más filas nuevas.
- Por crear en la Fase 3: archivos de valores límite, plantilla sin listas, CSV con y sin BOM, archivo de 20,000 filas.

## 7. Criterios

**Entrada** (se empieza a ejecutar cuando):

- La versión a probar está publicada en GitHub Pages y etiquetada en el repositorio.
- Los casos de prioridad alta están escritos y revisados.
- Los datos de prueba de la sección 6 existen.

**Salida** (se da por terminada cuando):

- 100% de los casos de prioridad alta ejecutados y aprobados; al menos 90% de los de prioridad media.
- Cero defectos críticos o altos abiertos; los medios y bajos tienen decisión documentada en Jira.
- La suite de Playwright corre en verde dos veces seguidas.

**Suspensión** (se detiene y se reporta si):

- La página no carga o no lee ningún archivo.
- "Agregar a mi archivo" produce un Excel que no abre o que pierde datos de la base.

Se reanuda cuando la corrección esté publicada; se repite el bloque afectado y la regresión automatizada.

## 8. Gestión de defectos

Cada bug se registra en Jira con título, ambiente, pasos, resultado esperado, resultado obtenido, evidencia (captura o archivo) y el caso de prueba que lo encontró.

| Severidad | Definición | Ejemplo en ValidaCarga |
| --- | --- | --- |
| Crítica | Pérdida o daño de datos, o la app no se puede usar | "Agregar a mi archivo" borra una pestaña de la base |
| Alta | Resultado incorrecto en una función principal | Una fila con "Master" sale como correcta |
| Media | Falla con alternativa, o mensaje confuso | El selector de pestaña no elige la correcta, pero se puede cambiar a mano |
| Baja | Detalle visual o de texto | Un texto se corta a 390 px |

La prioridad (qué se corrige primero) se decide aparte de la severidad, según el impacto en el usuario.

Flujo de estados en Jira:

1. Abierto
2. En corrección
3. Listo para probar
4. Verificado y cerrado, o reabierto si la prueba de confirmación falla

Dos defectos reales ya se encontraron durante la construcción y se registrarán como ejemplo: el panel de reglas fallaba al pasar a Ejemplos con una plantilla cargada, y ExcelJS duplicaba las listas desplegables al guardar.

## 9. Entregables y cronograma

Cinco entregables, en este orden; cada uno se revisa antes de empezar el siguiente.

1. **Plan de pruebas** (este documento) — 1 día. Hecho.
2. **Casos de prueba** en Excel, con trazabilidad a F01–F13 y R1–R8 — 2 días.
3. **Pruebas automatizadas** con Playwright + pytest en la carpeta tests/ del repositorio — 3 días.
4. **Ejecución y reporte**: resultados por caso, bugs en Jira y resumen de la corrida — 1 día.
5. **Narrativa para el portafolio** (Título / Desafío / Proceso / Resultado), con capturas y enlaces a la demo, el repositorio y Jira — 1 día.

Total estimado: 8 días de trabajo. Las fechas exactas quedan por definir según la disponibilidad de Adina.
