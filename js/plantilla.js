/*
 * ValidaCarga — lector de plantillas oficiales.
 * Abre la plantilla .xlsx que entrega el portal y deduce sus reglas:
 *   - columnas (fila de encabezados) y obligatorias (marcadas con * u "obligatorio")
 *   - listas desplegables: escritas en la celda, de un rango, de un nombre definido o de otra hoja
 *   - listas dependientes (cascada) con INDIRECTO / INDIRECT
 *   - números enteros/decimales, longitud de texto y fechas con rango
 *   - valores únicos (validación personalizada con CONTAR.SI / COUNTIF)
 * Después revisa la propia plantilla: si su fila de muestra no pasa sus reglas, lo reporta.
 *
 * Uso: Plantilla.leer(XLSX, datos, tipo) → modelo;  Plantilla.perfil(modelo, ajustes) → perfil para Validador.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Plantilla = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ------------------------------------------------------------------ utilidades
  const ENT = { '&quot;': '"', '&amp;': '&', '&lt;': '<', '&gt;': '>', '&apos;': "'" };
  const desXml = s => String(s).replace(/&(quot|amp|lt|gt|apos);/g, m => ENT[m]).replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n));
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const texto = v => v == null ? '' : String(v);

  function utf8(bytes) {
    if (typeof TextDecoder !== 'undefined') return new TextDecoder('utf-8').decode(bytes);
    return Buffer.from(bytes).toString('utf8');
  }

  function archivoZip(XLSX, zip, ruta) {
    const e = XLSX.CFB.find(zip, '/' + ruta.replace(/^\/+/, ''));
    return e && e.content ? utf8(e.content) : '';
  }

  function attrs(s) {
    const o = {};
    String(s).replace(/([\w:]+)="([^"]*)"/g, (_, k, v) => { o[k.replace(/^.*:/, '')] = desXml(v); return ''; });
    return o;
  }

  /** "Hoja!$A$1:$B$5" → { hoja, rango } (la hoja puede venir entre comillas simples). */
  function partirRef(ref) {
    const s = String(ref).replace(/^=/, '').trim();
    const m = s.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
    if (m) return { hoja: (m[1] || m[2]).replace(/''/g, "'"), rango: m[3].replace(/\$/g, '') };
    return { hoja: null, rango: s.replace(/\$/g, '') };
  }

  function valoresDeRango(XLSX, wb, hoja, rango) {
    const ws = wb.Sheets[hoja];
    if (!ws || !/^[A-Z]{1,3}\d*(:[A-Z]{1,3}\d*)?$/i.test(rango)) return null;
    let r;
    try {
      // Columnas completas ("A:A") se acotan al área usada de la hoja
      const usado = XLSX.utils.decode_range(ws['!ref'] || 'A1');
      const [a, b] = rango.split(':');
      const ini = XLSX.utils.decode_cell(/\d/.test(a) ? a : a + (usado.s.r + 1));
      const fin = XLSX.utils.decode_cell(b ? (/\d/.test(b) ? b : b + (usado.e.r + 1)) : (/\d/.test(a) ? a : a + (usado.e.r + 1)));
      r = { s: ini, e: { r: Math.min(fin.r, usado.e.r), c: fin.c } };
    } catch (e) { return null; }
    const vals = [];
    for (let i = r.s.r; i <= r.e.r; i++) {
      for (let j = r.s.c; j <= r.e.c; j++) {
        const cel = ws[XLSX.utils.encode_cell({ r: i, c: j })];
        if (cel && cel.v != null && texto(cel.v) !== '') vals.push(cel.w != null && cel.t === 'n' ? texto(cel.v) : texto(cel.v));
      }
    }
    return vals;
  }

  function serialAFecha(XLSX, n) {
    const p = XLSX.SSF.parse_date_code(+n);
    return p ? new Date(Date.UTC(p.y, p.m - 1, p.d)) : null;
  }

  // ------------------------------------------------------------------ lectura
  /**
   * @param XLSX  la librería SheetJS
   * @param datos ArrayBuffer/Uint8Array (navegador) o Buffer (Node)
   * @param tipo  'array' (por defecto) o 'buffer'
   */
  /**
   * Hoja que más probablemente es el formato: la primera visible que no parezca de instrucciones o catálogos.
   */
  function hojaPorDefecto(XLSX, wb) {
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    const visibles = wb.SheetNames.map((n, i) => ({ n, i })).filter(x => !meta[x.i] || !meta[x.i].Hidden);
    const auxiliar = /instrucc|l[eé]eme|ayuda|indicac|cat[aá]logo|listas?$|readme|help/i;
    // Tiene fila de encabezados: alguna de las 10 primeras filas con 2+ celdas con texto
    const conEncabezados = nombre => {
      const ws = wb.Sheets[nombre];
      if (!ws || !ws['!ref']) return false;
      const rg = XLSX.utils.decode_range(ws['!ref']);
      for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 9); r++) {
        let n = 0;
        for (let c = rg.s.c; c <= rg.e.c; c++) { const x = ws[XLSX.utils.encode_cell({ r, c })]; if (x && String(x.v == null ? '' : x.v).trim()) n++; }
        if (n >= 2) return true;
      }
      return false;
    };
    const orden = visibles.filter(x => !auxiliar.test(x.n)).concat(visibles.filter(x => auxiliar.test(x.n)));
    const candidata = orden.find(x => conEncabezados(x.n)) || orden[0];
    return candidata ? candidata.n : wb.SheetNames[0];
  }

  /** Lista de pestañas: [{ nombre, oculta }] */
  function hojas(wb) {
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    return wb.SheetNames.map((n, i) => ({ nombre: n, oculta: !!(meta[i] && meta[i].Hidden) }));
  }

  function leer(XLSX, datos, tipo, opciones) {
    tipo = tipo || 'array';
    const wb = XLSX.read(datos, { type: tipo, cellNF: true });
    let zip = null;
    try { zip = XLSX.CFB.read(datos, { type: tipo }); } catch (e) { zip = null; }

    // Hoja del formato: la primera visible
    const pedida = opciones && opciones.hoja;
    const nombreHoja = pedida && wb.SheetNames.includes(pedida) ? pedida : hojaPorDefecto(XLSX, wb);
    const iHoja = wb.SheetNames.indexOf(nombreHoja);
    const ws = wb.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) throw new Error(`La pestaña "${nombreHoja}" de la plantilla está vacía. Elige otra pestaña.`);

    // Fila de encabezados: la primera (de las 10 primeras) con al menos 2 celdas con texto
    const rg = XLSX.utils.decode_range(ws['!ref']);
    const celda = (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x && x.v != null ? texto(x.v) : ''; };
    let filaEnc = -1;
    for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 9) && filaEnc < 0; r++) {
      let n = 0;
      for (let c = rg.s.c; c <= rg.e.c; c++) if (celda(r, c).trim()) n++;
      if (n >= 2) filaEnc = r;
    }
    if (filaEnc < 0) throw new Error(`No encontré la fila de encabezados en las primeras 10 filas de la pestaña "${nombreHoja}". Elige otra pestaña.`);

    const columnas = [];
    const porIndice = {};
    for (let c = rg.s.c; c <= rg.e.c; c++) {
      const h = celda(filaEnc, c);
      if (!h.trim()) continue;
      const col = {
        nombre: h, indice: c, letra: XLSX.utils.encode_col(c),
        marcadaObligatoria: /\*|\(obligatori[oa]\)|obligatori[oa]\b/i.test(h),
        lista: null, dependiente: null, numero: null, largo: null, fecha: null, unico: false,
        inferidas: [], noResueltas: []
      };
      columnas.push(col);
      porIndice[c] = col;
    }

    // Filas de muestra que trae la propia plantilla
    const muestra = [];
    for (let r = filaEnc + 1; r <= rg.e.r && muestra.length < 50; r++) {
      const fila = columnas.map(col => {
        const x = ws[XLSX.utils.encode_cell({ r, c: col.indice })];
        if (!x || x.v == null) return '';
        if (x.t === 'n' && x.z && XLSX.SSF.is_date(x.z)) {
          const p = XLSX.SSF.parse_date_code(x.v);
          return `${String(p.d).padStart(2, '0')}/${String(p.m).padStart(2, '0')}/${p.y}`;
        }
        return texto(x.v);
      });
      if (fila.some(v => v.trim() !== '')) muestra.push({ fila: r + 1, celdas: fila });
    }

    // Nombres definidos (para listas y cascadas)
    const nombres = {};
    ((wb.Workbook && wb.Workbook.Names) || []).forEach(n => {
      if (n && n.Name && n.Ref && !/^_xlnm\./i.test(n.Name)) nombres[norm(n.Name)] = { nombre: n.Name, ref: n.Ref };
    });

    const avisos = [];
    const validacionesX14 = []; // formato moderno de Excel: se guardan para reescribirlas al agregar filas
    const resolverRef = (ref, hojaPorDefecto) => {
      const limpio = String(ref).replace(/^=/, '').trim();
      const nd = nombres[norm(limpio)];
      const destino = nd ? nd.ref : limpio;
      const { hoja, rango } = partirRef(destino);
      const vals = valoresDeRango(XLSX, wb, hoja || hojaPorDefecto, rango);
      return vals ? { valores: vals, fuente: nd ? `nombre "${nd.nombre}"` : `${hoja || hojaPorDefecto}!${rango}` } : null;
    };

    // Validaciones de datos (clásicas y extendidas x14)
    if (zip) {
      const rutaHoja = rutaDeHoja(XLSX, zip, nombreHoja) || `xl/worksheets/sheet${iHoja + 1}.xml`;
      const xml = archivoZip(XLSX, zip, rutaHoja);
      const re = /<(x14:)?dataValidation\b([^>]*)>([\s\S]*?)<\/(?:x14:)?dataValidation>/g;
      let m;
      while ((m = re.exec(xml))) {
        const a = attrs(m[2]);
        const cuerpo = m[3];
        const f1m = cuerpo.match(/<(?:x14:)?formula1>\s*(?:<xm:f>)?([\s\S]*?)(?:<\/xm:f>)?\s*<\/(?:x14:)?formula1>/);
        const f2m = cuerpo.match(/<(?:x14:)?formula2>\s*(?:<xm:f>)?([\s\S]*?)(?:<\/xm:f>)?\s*<\/(?:x14:)?formula2>/);
        const sqm = cuerpo.match(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/);
        const sqref = a.sqref || (sqm ? sqm[1] : '');
        const f1 = f1m ? desXml(f1m[1]).trim() : '';
        const f2 = f2m ? desXml(f2m[1]).trim() : '';
        const tipoV = a.type || 'any';
        if (m[1]) validacionesX14.push({ sqref, type: tipoV, operator: a.operator, f1, f2, allowBlank: a.allowBlank === '1' || a.allowBlank === 'true' });
        columnasDeSqref(XLSX, sqref, filaEnc).forEach(ci => {
          const col = porIndice[ci];
          if (!col) return;
          aplicarValidacion(XLSX, col, tipoV, a.operator || 'between', f1, f2, { resolverRef, hoja: nombreHoja, porIndice });
        });
      }
    } else {
      avisos.push('No pude abrir el interior de la plantilla; solo se usarán los encabezados.');
    }

    // Reglas inferidas por el nombre de la columna (solo si no hay una regla explícita)
    columnas.forEach(col => {
      const n = norm(col.nombre);
      if (!col.lista && !col.dependiente) {
        if (/correo|e-?mail/.test(n)) col.inferidas.push('correo');
        else if (/(telefono|celular|movil|msisdn|numero de linea|\bdn\b)/.test(n)) col.inferidas.push('telefono');
      }
      if (!col.fecha && /fecha/.test(n)) col.inferidas.push('fecha');
    });

    const hayMarcas = columnas.some(c => c.marcadaObligatoria);
    columnas.forEach(c => { c.obligatorio = hayMarcas ? c.marcadaObligatoria : true; });

    const modelo = {
      hoja: nombreHoja, hojas: hojas(wb), filaEncabezados: filaEnc + 1, columnas, muestra, avisos, validacionesX14,
      obligatoriosPorMarca: hayMarcas,
      listaDependiente: (nombreValor) => {
        const candidatos = [nombreValor, nombreValor.replace(/ /g, '_'), nombreValor.replace(/[^\p{L}\p{N}_]/gu, '_'), nombreValor.replace(/[^\p{L}\p{N}]/gu, '')];
        for (const k of candidatos) {
          const nd = nombres[norm(k)];
          if (nd) { const r = resolverRef(nd.nombre, nombreHoja); if (r) return r.valores; }
        }
        return null;
      }
    };
    modelo.revision = revisar(modelo);
    return modelo;
  }

  function rutaDeHoja(XLSX, zip, nombreHoja) {
    const wbx = archivoZip(XLSX, zip, 'xl/workbook.xml');
    const rels = archivoZip(XLSX, zip, 'xl/_rels/workbook.xml.rels');
    let rid = null;
    wbx.replace(/<sheet\b([^>]*)\/?>/g, (_, at) => { const a = attrs(at); if (a.name === nombreHoja) rid = a.id; return ''; });
    if (!rid) return null;
    let destino = null;
    rels.replace(/<Relationship\b([^>]*)\/?>/g, (_, at) => { const a = attrs(at); if (a.Id === rid) destino = a.Target; return ''; });
    if (!destino) return null;
    return destino.startsWith('/') ? destino.slice(1) : 'xl/' + destino.replace(/^\.\//, '');
  }

  /** "D2:D500 F2" → índices de columna que aplican a filas de datos. */
  function columnasDeSqref(XLSX, sqref, filaEnc) {
    const cols = new Set();
    String(sqref).split(/\s+/).filter(Boolean).forEach(parte => {
      const [a, b] = parte.split(':');
      const ca = a.match(/^([A-Z]{1,3})(\d*)$/i), cb = (b || a).match(/^([A-Z]{1,3})(\d*)$/i);
      if (!ca || !cb) return;
      const filaFin = cb[2] ? +cb[2] - 1 : Infinity;
      if (filaFin <= filaEnc) return; // solo cubre encabezados
      for (let c = XLSX.utils.decode_col(ca[1]); c <= XLSX.utils.decode_col(cb[1]); c++) cols.add(c);
    });
    return [...cols];
  }

  const num = s => { const n = parseFloat(String(s).replace(/^=/, '')); return isFinite(n) && /^=?-?\d+(\.\d+)?$/.test(String(s).trim()) ? n : null; };

  function aplicarValidacion(XLSX, col, tipo, operador, f1, f2, ctx) {
    const f1s = f1.replace(/^=/, '');
    if (tipo === 'list') {
      if (/^"[\s\S]*"$/.test(f1s)) {
        col.lista = { valores: f1s.slice(1, -1).split(',').map(s => s.trim()).filter(Boolean), fuente: 'lista escrita en la validación' };
      } else if (/^INDIRECT[O]?\s*\(/i.test(f1s)) {
        const ref = f1s.match(/\$?([A-Z]{1,3})\$?\d+/i);
        const padre = ref ? ctx.porIndice[XLSX.utils.decode_col(ref[1].toUpperCase())] : null;
        if (padre && padre !== col) col.dependiente = { padre: padre.nombre, formula: f1s };
        else col.noResueltas.push(`Lista dependiente no reconocida: ${f1s}`);
      } else {
        const r = ctx.resolverRef(f1s, ctx.hoja);
        if (r) col.lista = { valores: r.valores, fuente: r.fuente };
        else col.noResueltas.push(`Lista con origen no reconocido: ${f1s}`);
      }
    } else if (tipo === 'whole' || tipo === 'decimal' || tipo === 'textLength') {
      const a = num(f1), b = num(f2);
      if (a === null) { col.noResueltas.push(`Regla de ${tipo === 'textLength' ? 'longitud' : 'número'} con fórmula: ${f1}`); return; }
      const regla = { operador, a, b };
      if (tipo === 'textLength') col.largo = regla;
      else col.numero = Object.assign({ entero: tipo === 'whole' }, regla);
    } else if (tipo === 'date') {
      const a = num(f1), b = num(f2);
      if (a === null) { col.noResueltas.push(`Regla de fecha con fórmula: ${f1}`); col.fecha = { soloFormato: true }; return; }
      col.fecha = { operador, a: serialAFecha(XLSX, a), b: b === null ? null : serialAFecha(XLSX, b) };
    } else if (tipo === 'custom') {
      if (/^(COUNTIF|CONTAR\.SI)\s*\(/i.test(f1s) && /[<=]=?\s*1\s*$/.test(f1s)) col.unico = true;
      else col.noResueltas.push(`Validación personalizada: ${f1s}`);
    }
  }

  // ------------------------------------------------------------------ perfil
  function cumple(op, v, a, b) {
    switch (op) {
      case 'between': return v >= a && v <= b;
      case 'notBetween': return v < a || v > b;
      case 'equal': return v === a;
      case 'notEqual': return v !== a;
      case 'greaterThan': return v > a;
      case 'lessThan': return v < a;
      case 'greaterThanOrEqual': return v >= a;
      case 'lessThanOrEqual': return v <= a;
      default: return true;
    }
  }
  function describir(op, a, b, f, exacto) {
    f = f || (x => x);
    return {
      between: `entre ${f(a)} y ${f(b)}`, notBetween: `fuera del rango ${f(a)}–${f(b)}`, equal: exacto ? `exactamente ${f(a)}` : `igual a ${f(a)}`,
      notEqual: `distinto de ${f(a)}`, greaterThan: `mayor que ${f(a)}`, lessThan: `menor que ${f(a)}`,
      greaterThanOrEqual: `de ${f(a)} o más`, lessThanOrEqual: `de ${f(a)} o menos`
    }[op] || '';
  }
  const etq = n => String(n).replace(/\s*\*+\s*$/, '').replace(/\s*\(obligatori[oa]\)\s*$/i, '');
  const fmtFecha = d => d ? `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}` : '';

  function revisarLista(api, columna, valor, opciones, contexto) {
    const u = api.ctx.utiles;
    if (u.vacio(valor)) return false;
    const r = u.revisarCatalogo(valor, opciones);
    if (r.ok) return true;
    if (r.sugerencia) {
      api.agregar(api.SEV.ERROR, columna,
        `Se parece a "${r.sugerencia}", pero no es idéntico (acentos, mayúsculas o espacios). El portal compara texto exacto y rechaza la fila.`,
        `Escribir exactamente "${r.sugerencia}".`);
    } else {
      api.agregar(api.SEV.ERROR, columna, `"${valor}" no está en la lista de la plantilla${contexto ? ' ' + contexto : ''}.`,
        `Usar uno de: ${opciones.slice(0, 6).join(', ')}${opciones.length > 6 ? '…' : ''}.`);
    }
    return false;
  }

  /**
   * Convierte el modelo en un perfil para Validador.validar.
   * @param ajustes { obligatorios: { [nombreColumna]: boolean } } — cambios hechos por el usuario
   */
  function perfil(modelo, ajustes) {
    const obl = (ajustes && ajustes.obligatorios) || {};
    const cols = modelo.columnas;
    return {
      id: 'plantilla',
      nombre: 'Según plantilla oficial',
      destino: modelo.archivo || 'Plantilla',
      columnas: cols.map(c => ({
        nombre: c.nombre,
        obligatorio: obl[c.nombre] !== undefined ? obl[c.nombre] : c.obligatorio,
        catalogo: !!(c.lista || c.dependiente)
      })),
      validarFila(d, api) {
        const { SEV, agregar, ctx } = api;
        const u = ctx.utiles;
        const validas = {};
        cols.forEach(c => {
          const v = d[c.nombre];
          if (u.vacio(v)) return;
          let ok = true;

          if (c.lista) ok = revisarLista(api, c.nombre, v, c.lista.valores);

          if (c.dependiente) {
            const pv = d[c.dependiente.padre];
            if (!u.vacio(pv) && validas[c.dependiente.padre] !== false) {
              const opciones = modelo.listaDependiente(pv);
              if (opciones) ok = revisarLista(api, c.nombre, v, opciones, `para ${etq(c.dependiente.padre)} "${pv}"`);
              else agregar(SEV.ADVERTENCIA, c.nombre, `La plantilla no trae la lista de ${etq(c.nombre)} para ${etq(c.dependiente.padre)} "${pv}", así que no se pudo revisar.`,
                'Confirmar el valor directamente en el portal.');
            }
          }

          if (c.numero) {
            const n = u.leerMonto(v);
            const desc = `${c.numero.entero ? 'un número entero' : 'un número'} ${describir(c.numero.operador, c.numero.a, c.numero.b)}`;
            if (!n || (c.numero.entero && !Number.isInteger(n.numero)) || !cumple(c.numero.operador, n.numero, c.numero.a, c.numero.b)) {
              ok = false;
              agregar(SEV.ERROR, c.nombre, `"${v}" no cumple la regla de la plantilla: debe ser ${desc}.`, `Capturar ${desc}.`);
            }
          }

          if (c.largo) {
            const L = v.trim().length;
            if (!cumple(c.largo.operador, L, c.largo.a, c.largo.b)) {
              ok = false;
              agregar(SEV.ERROR, c.nombre, `Tiene ${L} caracteres; la plantilla pide ${describir(c.largo.operador, c.largo.a, c.largo.b, null, true)} caracteres.`, 'Ajustar la longitud del texto.');
            }
          }

          if (c.fecha || c.inferidas.includes('fecha')) {
            const f = u.leerFecha(v);
            if (!f) {
              ok = false;
              agregar(c.fecha ? SEV.ERROR : SEV.ADVERTENCIA, c.nombre, `"${v}" no es una fecha válida${c.fecha ? '' : ' (la columna se llama "fecha")'}.`,
                'Usar el formato DD/MM/AAAA con una fecha que exista.');
            } else if (c.fecha && c.fecha.a && !cumple(c.fecha.operador, f.fecha.getTime(), c.fecha.a.getTime(), c.fecha.b && c.fecha.b.getTime())) {
              ok = false;
              agregar(SEV.ERROR, c.nombre, `La fecha ${f.texto} está fuera de lo que pide la plantilla: ${describir(c.fecha.operador, c.fecha.a, c.fecha.b, fmtFecha)}.`, 'Revisar la fecha.');
            }
          }

          if (c.inferidas.includes('correo') && !/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(v.trim())) {
            agregar(SEV.ADVERTENCIA, c.nombre, 'No parece un correo válido (revisión por el nombre de la columna).', 'Usar el formato nombre@dominio.com.');
          }
          if (c.inferidas.includes('telefono') && !/^\d{10}$/.test(v.trim())) {
            agregar(SEV.ADVERTENCIA, c.nombre, 'No parece un número de 10 dígitos (revisión por el nombre de la columna).', 'Capturar 10 dígitos, sin espacios ni lada internacional.');
          }

          if (c.unico) {
            const vis = ctx.vistos.unico = ctx.vistos.unico || {};
            const k = c.nombre + '|' + v.trim().toUpperCase();
            if (vis[k]) {
              ok = false;
              agregar(SEV.ERROR, c.nombre, `Valor repetido: ya aparece en la fila ${vis[k]}. La plantilla exige valores únicos en esta columna.`, 'Dejar una sola fila con este valor.');
            } else vis[k] = api.fila;
          }
          validas[c.nombre] = ok;
        });
      }
    };
  }

  // ------------------------------------------------------------------ revisión de la propia plantilla
  function revisar(modelo) {
    const hallazgos = [];
    modelo.columnas.forEach(c => {
      if (c.lista) {
        const conEspacios = c.lista.valores.filter(v => v !== v.trim());
        if (conEspacios.length) hallazgos.push({ severidad: 'advertencia', columna: c.nombre,
          motivo: `La lista de la plantilla tiene opciones con espacios sobrantes (${conEspacios.map(v => `"${v}"`).join(', ')}).`,
          correccion: 'Si el portal compara contra la versión sin espacios, estas opciones nunca van a pasar. Reportarlo al dueño del portal.' });
        const vistos = new Set(), rep = new Set();
        c.lista.valores.forEach(v => { (vistos.has(v) ? rep : vistos).add(v); });
        if (rep.size) hallazgos.push({ severidad: 'advertencia', columna: c.nombre, motivo: `La lista tiene opciones repetidas: ${[...rep].join(', ')}.`, correccion: 'Sin impacto en la carga, pero indica un catálogo descuidado.' });
      }
      if (c.dependiente) {
        const padre = modelo.columnas.find(p => p.nombre === c.dependiente.padre);
        if (padre && padre.lista) {
          const sinLista = padre.lista.valores.filter(v => !modelo.listaDependiente(v));
          if (sinLista.length) hallazgos.push({ severidad: 'advertencia', columna: c.nombre,
            motivo: `No hay lista de ${etq(c.nombre)} para: ${sinLista.slice(0, 5).join(', ')}${sinLista.length > 5 ? '…' : ''}.`,
            correccion: `Quien elija esas opciones de ${etq(padre.nombre)} no podrá llenar ${etq(c.nombre)}.` });
        }
      }
      c.noResueltas.forEach(t => hallazgos.push({ severidad: 'advertencia', columna: c.nombre, motivo: `${t}. ValidaCarga no puede revisar esta regla.`, correccion: 'Revisarla a mano.' }));
    });
    return hallazgos;
  }

  return { leer, perfil, hojas, etiqueta: etq };
});
