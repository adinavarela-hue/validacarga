/*
 * ValidaCarga — motor de validación (sin dependencias, corre en navegador y en Node).
 * Recibe: un perfil, los encabezados y las filas del archivo (arreglos de celdas).
 * Devuelve: incidencias a nivel archivo y resultado fila por fila.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Validador = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SEV = { ERROR: 'error', ADVERTENCIA: 'advertencia' };

  /** Quita acentos, colapsa espacios y pasa a minúsculas. */
  function normalizar(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/\s+/g, ' ').trim().toLowerCase();
  }
  /** Igual que normalizar, pero además elimina todos los espacios ("Tipo deCanal" = "Tipo de Canal"). */
  function clave(s) { return normalizar(s).replace(/\s/g, ''); }

  function texto(v) { return v == null ? '' : String(v); }
  function vacio(v) { return texto(v).trim() === ''; }

  /**
   * Revisa un valor contra un catálogo con coincidencia EXACTA (como lo hace el portal).
   * Si no es exacto pero se parece (acentos, mayúsculas, espacios), sugiere el valor correcto.
   */
  function revisarCatalogo(valor, opciones) {
    const v = texto(valor);
    if (opciones.includes(v)) return { ok: true };
    const parecido = opciones.find(o => normalizar(o) === normalizar(v));
    return { ok: false, sugerencia: parecido || null };
  }

  function pad(n) { return String(n).padStart(2, '0'); }

  /** Acepta DD/MM/AAAA, D/M/AAAA o AAAA-MM-DD. Devuelve {fecha, iso} o null si no es una fecha real. */
  function leerFecha(valor) {
    const s = texto(valor).trim();
    let d, m, a;
    let r = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (r) { d = +r[1]; m = +r[2]; a = +r[3]; }
    else if ((r = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) { a = +r[1]; m = +r[2]; d = +r[3]; }
    else return null;
    const f = new Date(Date.UTC(a, m - 1, d));
    if (f.getUTCFullYear() !== a || f.getUTCMonth() !== m - 1 || f.getUTCDate() !== d) return null;
    return { fecha: f, iso: `${a}-${pad(m)}-${pad(d)}`, texto: `${pad(d)}/${pad(m)}/${a}` };
  }

  /** Convierte "$1,000.50" o 100 a número. Devuelve {numero, conFormato} o null. */
  function leerMonto(valor) {
    const s = texto(valor).trim();
    if (s === '') return null;
    const sinSimbolo = s.replace(/^\$\s*/, '').replace(/^MXN\s*/i, '').trim();
    // Miles con coma solo si son grupos de 3 dígitos: "1,000" sí, "1,00" no (ambiguo)
    if (!/^-?(\d+|\d{1,3}(,\d{3})+)(\.\d+)?$/.test(sinSimbolo)) return null;
    const limpio = sinSimbolo.replace(/,/g, '');
    return { numero: parseFloat(limpio), conFormato: limpio !== s };
  }

  /** Empareja los encabezados del archivo con las columnas del perfil. */
  function mapearEncabezados(perfil, encabezados) {
    const incidencias = [];
    const indice = {};
    const usados = new Set();
    perfil.columnas.forEach(col => {
      let i = encabezados.findIndex(h => texto(h) === col.nombre);
      if (i === -1) {
        i = encabezados.findIndex(h => clave(h) === clave(col.nombre));
        if (i !== -1) {
          incidencias.push({
            severidad: SEV.ADVERTENCIA, columna: col.nombre, valor: texto(encabezados[i]),
            motivo: `El encabezado "${texto(encabezados[i])}" no es idéntico a "${col.nombre}". ValidaCarga lo reconoce, pero el portal podría no hacerlo.`,
            correccion: `Renombrar el encabezado a "${col.nombre}" exactamente.`
          });
        }
      }
      if (i === -1) {
        if (col.obligatorio) {
          incidencias.push({
            severidad: SEV.ERROR, columna: col.nombre, valor: '',
            motivo: `Falta la columna obligatoria "${col.nombre}".`,
            correccion: `Agregar la columna "${col.nombre}" en la fila de encabezados.`
          });
        }
        return;
      }
      indice[col.nombre] = i;
      usados.add(i);
    });
    encabezados.forEach((h, i) => {
      if (!usados.has(i) && !vacio(h)) {
        incidencias.push({
          severidad: SEV.ADVERTENCIA, columna: texto(h), valor: '',
          motivo: `La columna "${texto(h)}" no forma parte del formato "${perfil.nombre}".`,
          correccion: 'Eliminarla antes de subir el archivo, o confirmar que el portal la acepta.'
        });
      }
    });
    const fatal = incidencias.some(x => x.severidad === SEV.ERROR);
    return { indice, incidencias, fatal };
  }

  /**
   * Valida un archivo completo.
   * @param {object} perfil  — definición de perfiles.js
   * @param {Array} encabezados — primera fila del archivo
   * @param {Array<Array>} filas — filas de datos (sin encabezado)
   * @param {object} [opciones] — { hoy: Date } para pruebas reproducibles
   */
  function validar(perfil, encabezados, filas, opciones) {
    const hoy = (opciones && opciones.hoy) || new Date();
    const mapa = mapearEncabezados(perfil, encabezados || []);
    const resultado = {
      perfil: perfil.id, archivo: mapa.incidencias, filas: [], resumen: null, fatal: mapa.fatal,
      indice: mapa.indice,
      faltantes: perfil.columnas.filter(c => mapa.indice[c.nombre] === undefined).map(c => c.nombre)
    };

    // Aunque falten columnas, se revisan las que sí vienen (útil para archivos que acumulan información).
    {
      const ctx = { vistos: {}, hoy, utiles: utilidades };
      filas.forEach((celdas, k) => {
        if (!celdas || celdas.every(vacio)) return; // fila en blanco: se ignora
        const datos = {};
        perfil.columnas.forEach(col => {
          const i = mapa.indice[col.nombre];
          datos[col.nombre] = i === undefined ? '' : texto(celdas[i]);
        });
        const numFila = k + ((opciones && opciones.primeraFila) || 2); // fila real en Excel
        const incidencias = [];
        const agregar = (severidad, columna, motivo, correccion) => incidencias.push({
          severidad, columna, valor: columna ? texto(datos[columna]) : '', motivo, correccion
        });

        // Reglas comunes: obligatorios y espacios sobrantes
        perfil.columnas.forEach(col => {
          if (mapa.indice[col.nombre] === undefined) return; // columna ausente: ya se reportó a nivel archivo
          const v = datos[col.nombre];
          if (col.obligatorio && vacio(v)) {
            agregar(SEV.ERROR, col.nombre, 'Campo obligatorio vacío.', `Capturar ${col.nombre}.`);
          } else if (!col.catalogo && v !== v.trim()) {
            agregar(SEV.ADVERTENCIA, col.nombre, 'Tiene espacios al inicio o al final.', 'Quitar los espacios sobrantes.');
          }
        });

        perfil.validarFila(datos, { agregar, SEV, ctx, fila: numFila });
        resultado.filas.push({ fila: numFila, estado: estadoDe(incidencias), incidencias, datos, celdas });
      });
    }

    resultado.resumen = resumir(resultado.filas);
    return resultado;
  }

  function estadoDe(incidencias) {
    if (incidencias.some(i => i.severidad === SEV.ERROR)) return 'error';
    if (incidencias.length) return 'advertencia';
    return 'correcto';
  }

  function resumir(filas) {
    const r = { total: filas.length, correcto: 0, advertencia: 0, error: 0 };
    filas.forEach(f => { r[f.estado]++; });
    r.validas = r.correcto + r.advertencia;
    return r;
  }

  /** Filas para el reporte de errores (una línea por incidencia). */
  function filasReporte(resultado) {
    const salida = [['Fila', 'Estado de la fila', 'Severidad', 'Columna', 'Valor', 'Motivo', 'Cómo corregirlo']];
    resultado.archivo.forEach(i => salida.push(['Encabezados', 'archivo', i.severidad, i.columna, i.valor, i.motivo, i.correccion]));
    resultado.filas.forEach(f => {
      f.incidencias.forEach(i => salida.push([f.fila, f.estado, i.severidad, i.columna, i.valor, i.motivo, i.correccion]));
    });
    return salida;
  }

  /** Encabezado original + solo las filas sin errores (correctas y con advertencia), sin modificar. */
  function filasValidas(resultado, encabezados) {
    if (resultado.fatal) return [encabezados];
    return [encabezados].concat(resultado.filas.filter(f => f.estado !== 'error').map(f => f.celdas));
  }

  /**
   * Encuentra la fila de encabezados entre las primeras 10.
   * Si se dan los encabezados esperados, elige la fila con más coincidencias; si no, la primera con 2+ celdas con texto.
   */
  function buscarEncabezados(filas, esperados) {
    const limite = Math.min(filas.length, 10);
    if (esperados && esperados.length) {
      const claves = new Set(esperados.map(clave));
      let mejor = -1, puntos = 0;
      for (let i = 0; i < limite; i++) {
        const n = (filas[i] || []).filter(h => claves.has(clave(h))).length;
        if (n > puntos) { puntos = n; mejor = i; }
      }
      if (mejor >= 0) return mejor;
    }
    for (let i = 0; i < limite; i++) if ((filas[i] || []).filter(v => !vacio(v)).length >= 2) return i;
    return 0;
  }

  const utilidades = { normalizar, clave, revisarCatalogo, leerFecha, leerMonto, vacio, texto };

  return Object.assign({ SEV, validar, mapearEncabezados, filasReporte, filasValidas, buscarEncabezados }, utilidades);
});
