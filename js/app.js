/* ValidaCarga — interfaz. Lee la plantilla y el archivo en el navegador, llama al motor y pinta resultados. */
(function () {
  'use strict';

  const MAX_FILAS_TABLA = 500;
  const $ = sel => document.querySelector(sel);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  const estado = {
    modo: 'plantilla',      // 'plantilla' | 'ejemplos'
    perfilId: 'usuarios',   // perfil de ejemplo elegido
    plantilla: null,        // { nombre, modelo, esDemo }
    obligatorios: {},       // ajustes del usuario sobre la plantilla
    archivo: null,          // { nombre, tipo, crudas, esMuestra }
    preparado: null,        // { encabezados, filas, primeraFila }
    resultado: null,
    revisionMuestra: null,  // resultado de validar la fila de muestra de la plantilla
    filtro: 'todas',
    mensaje: '',
    mensajeError: false
  };

  // ---------------------------------------------------------------- Perfil vigente
  function perfilActual() {
    if (estado.modo === 'ejemplos') return Perfiles.porId[estado.perfilId];
    if (!estado.plantilla || !estado.plantilla.modelo) return null;
    return Plantilla.perfil(estado.plantilla.modelo, { obligatorios: estado.obligatorios });
  }

  // ---------------------------------------------------------------- Lectura de archivos
  const pad = n => String(n).padStart(2, '0');

  function hojaAFilas(ws) {
    if (!ws || !ws['!ref']) return [];
    const rg = XLSX.utils.decode_range(ws['!ref']);
    const filas = [];
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      const fila = [];
      for (let c = rg.s.c; c <= rg.e.c; c++) fila.push(valorCelda(ws[XLSX.utils.encode_cell({ r, c })]));
      filas.push(fila);
    }
    return filas;
  }

  function valorCelda(cel) {
    if (!cel || cel.v == null) return '';
    if (cel.t === 'd' && cel.v instanceof Date) return `${pad(cel.v.getDate())}/${pad(cel.v.getMonth() + 1)}/${cel.v.getFullYear()}`;
    if (cel.t === 'n' && cel.z && XLSX.SSF.is_date(cel.z)) {
      const p = XLSX.SSF.parse_date_code(cel.v);
      return `${pad(p.d)}/${pad(p.m)}/${p.y}`;
    }
    if (cel.t === 'b') return cel.v ? 'VERDADERO' : 'FALSO';
    return String(cel.v);
  }

  function extension(nombre) { return (nombre.split('.').pop() || '').toLowerCase(); }

  async function leerArchivo(file) {
    const ext = extension(file.name);
    if (!['xlsx', 'xls', 'csv'].includes(ext)) throw new Error(`"${file.name}" no es un Excel ni un CSV. Usa un archivo .xlsx, .xls o .csv.`);
    let libro;
    if (ext === 'csv') {
      const texto = await file.text();
      libro = XLSX.read(texto.replace(/^﻿/, ''), { type: 'string', raw: true });
    } else {
      libro = XLSX.read(await file.arrayBuffer(), { type: 'array', cellNF: true });
    }
    const meta = (libro.Workbook && libro.Workbook.Sheets) || [];
    const pestanas = libro.SheetNames.map((n, i) => ({ nombre: n, oculta: !!(meta[i] && meta[i].Hidden), crudas: hojaAFilas(libro.Sheets[n]) }));
    if (!pestanas.some(h => tieneDatos(h.crudas))) throw new Error(`"${file.name}" está vacío.`);
    return { nombre: file.name, tipo: ext === 'csv' ? 'csv' : 'xlsx', pestanas, hoja: 0, hojaManual: false, crudas: [], esMuestra: false };
  }

  const tieneDatos = filas => filas.some(f => f.some(v => String(v).trim() !== ''));

  /** Cuántos encabezados esperados aparecen en la mejor fila de las 10 primeras. */
  function puntaje(filas, esperados) {
    if (!esperados || !esperados.length) return 0;
    const claves = new Set(esperados.map(Validador.clave));
    let mejor = 0;
    for (let i = 0; i < Math.min(filas.length, 10); i++) {
      mejor = Math.max(mejor, (filas[i] || []).filter(h => claves.has(Validador.clave(h))).length);
    }
    return mejor;
  }

  /** Elige la pestaña cuyos encabezados más se parecen a lo que espera el perfil; si empatan, la primera visible con datos. */
  function elegirPestana(archivo, perfil) {
    if (!archivo.pestanas) return;
    if (!archivo.hojaManual) {
      const esperados = perfil ? perfil.columnas.map(c => c.nombre) : null;
      let mejor = -1, puntos = -1;
      archivo.pestanas.forEach((h, i) => {
        if (!tieneDatos(h.crudas)) return;
        const p = puntaje(h.crudas, esperados) + (h.oculta ? -0.5 : 0);
        if (p > puntos) { puntos = p; mejor = i; }
      });
      archivo.hoja = mejor < 0 ? 0 : mejor;
    }
    archivo.crudas = archivo.pestanas[archivo.hoja].crudas;
  }

  function opcionesPestanas(lista, actual) {
    return lista.map((h, i) => `<option value="${i}" ${i === actual ? 'selected' : ''}>${i + 1} · ${esc(h.nombre)}${h.oculta ? ' (oculta)' : ''}</option>`).join('');
  }

  /** Ubica la fila de encabezados del archivo según lo que espera el perfil vigente. */
  function preparar(archivo, perfil) {
    const esperados = perfil ? perfil.columnas.map(c => c.nombre) : null;
    const i = Validador.buscarEncabezados(archivo.crudas, esperados);
    const encabezados = (archivo.crudas[i] || []).slice();
    while (encabezados.length && String(encabezados[encabezados.length - 1]).trim() === '') encabezados.pop();
    return {
      encabezados,
      filas: archivo.crudas.slice(i + 1).map(f => f.slice(0, encabezados.length)),
      primeraFila: i + 2,
      filaEncabezados: i + 1
    };
  }

  // ---------------------------------------------------------------- Plantilla
  function base64ABytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  /** Lee la plantilla (en la pestaña pedida o la más probable). Si esa pestaña no sirve, deja elegir otra. */
  function usarPlantilla(nombre, bytes, esDemo, hoja) {
    estado.obligatorios = {};
    estado.revisionMuestra = null;
    try {
      const modelo = Plantilla.leer(XLSX, bytes, 'array', { hoja });
      modelo.archivo = nombre;
      estado.plantilla = { nombre, bytes, esDemo, modelo, hojas: modelo.hojas, hoja: modelo.hoja, error: '' };
      revisarMuestraPlantilla();
    } catch (err) {
      const hojas = Plantilla.hojas(XLSX.read(bytes, { type: 'array' }));
      if (hojas.length < 2) throw err;
      estado.plantilla = { nombre, bytes, esDemo, modelo: null, hojas, hoja: hoja || null, error: err.message };
    }
  }

  async function recibirPlantilla(file) {
    if (!file) return;
    try {
      if (extension(file.name) !== 'xlsx') {
        throw new Error(`La plantilla debe ser .xlsx (recibí "${file.name}"). Si el portal la entrega en otro formato, ábrela en Excel y guárdala como .xlsx.`);
      }
      usarPlantilla(file.name, new Uint8Array(await file.arrayBuffer()), false);
      estado.mensaje = ''; estado.mensajeError = false;
      validarActual();
    } catch (err) {
      estado.mensaje = err.message || 'No se pudo leer la plantilla.';
      estado.mensajeError = true;
      pintarTodo();
    }
  }

  /** Valida la fila de muestra que trae la propia plantilla (así se detectó el bug real). */
  function revisarMuestraPlantilla() {
    if (!estado.plantilla || !estado.plantilla.modelo) { estado.revisionMuestra = null; return; }
    const m = estado.plantilla.modelo;
    const perfil = Plantilla.perfil(m, { obligatorios: estado.obligatorios });
    // Con más de 5 filas ya no es una muestra: es un archivo que acumula información
    estado.revisionMuestra = m.muestra.length && m.muestra.length <= 5
      ? Validador.validar(perfil, m.columnas.map(c => c.nombre), m.muestra.map(x => x.celdas), { primeraFila: m.muestra[0].fila })
      : null;
  }

  const esAcumulado = () => !!(estado.plantilla && estado.plantilla.modelo && estado.plantilla.modelo.muestra.length > 5);

  function cargarDemo() {
    usarPlantilla(DemoPlantilla.nombre, base64ABytes(DemoPlantilla.base64), true);
    estado.archivo = { nombre: DemoPlantilla.datos.nombre, tipo: 'xlsx', crudas: DemoPlantilla.datos.filas.map(f => f.slice()), esMuestra: true };
    estado.archivo.pestanas = [{ nombre: 'Formato', oculta: false, crudas: estado.archivo.crudas }]; estado.archivo.hoja = 0;
    estado.mensaje = ''; estado.mensajeError = false; estado.filtro = 'todas';
    validarActual();
  }

  function quitarPlantilla() {
    estado.plantilla = null;
    estado.obligatorios = {};
    estado.revisionMuestra = null;
    estado.resultado = null;
    if (estado.archivo && estado.archivo.esMuestra) estado.archivo = null;
    estado.mensaje = ''; estado.mensajeError = false;
    $('#input-plantilla').value = '';
    pintarTodo();
  }

  // ---------------------------------------------------------------- Archivo de datos
  async function recibirArchivo(file) {
    if (!file) return;
    try {
      estado.archivo = await leerArchivo(file);
      estado.mensaje = '';
      estado.mensajeError = false;
      estado.filtro = 'todas';
      validarActual();
    } catch (err) {
      estado.mensaje = err.message || 'No se pudo leer el archivo.';
      estado.mensajeError = true;
      pintarTodo();
    }
  }

  function cargarMuestraEjemplo() {
    const m = Perfiles.porId[estado.perfilId].muestra;
    estado.archivo = { nombre: m.archivo, tipo: 'xlsx', crudas: [m.encabezados.slice()].concat(m.filas.map(f => f.slice())), esMuestra: true };
    estado.archivo.pestanas = [{ nombre: 'Datos', oculta: false, crudas: estado.archivo.crudas }]; estado.archivo.hoja = 0;
    estado.mensaje = ''; estado.mensajeError = false; estado.filtro = 'todas';
    validarActual();
  }

  function validarActual() {
    const perfil = perfilActual();
    estado.resultado = null;
    estado.preparado = null;
    if (estado.archivo) elegirPestana(estado.archivo, perfil);
    if (perfil && estado.archivo) {
      estado.preparado = preparar(estado.archivo, perfil);
      estado.resultado = Validador.validar(perfil, estado.preparado.encabezados, estado.preparado.filas, { primeraFila: estado.preparado.primeraFila });
    }
    pintarTodo();
  }

  /** Empieza de cero: borra plantilla, reglas detectadas, ajustes, archivo y resultados. */
  function limpiar() {
    estado.plantilla = null;
    estado.obligatorios = {};
    estado.revisionMuestra = null;
    estado.archivo = null;
    estado.preparado = null;
    estado.resultado = null;
    estado.filtro = 'todas';
    estado.mensaje = ''; estado.mensajeError = false;
    $('#input-archivo').value = '';
    $('#input-plantilla').value = '';
    pintarTodo();
    (estado.modo === 'plantilla' ? $('#zona-plantilla') : $('#zona')).focus();
  }

  // ---------------------------------------------------------------- Pintado: panel izquierdo
  function pintarTodo() {
    pintarModo();
    pintarPlantilla();
    pintarArchivo();
    pintarResultados();
  }

  function pintarModo() {
    document.querySelectorAll('[data-modo]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.modo === estado.modo)));
    $('#panel-plantilla').hidden = estado.modo !== 'plantilla';
    $('#panel-ejemplos').hidden = estado.modo !== 'ejemplos';
    $('#btn-plantilla').hidden = estado.modo !== 'ejemplos';
    $('#btn-muestra').textContent = estado.modo === 'plantilla' ? 'Probar con la plantilla de ejemplo' : 'Probar con archivo de ejemplo';
  }

  function pintarPerfiles() {
    $('#perfiles').innerHTML = Perfiles.lista.map(p => `
      <label class="perfil" data-testid="perfil-${p.id}">
        <input type="radio" name="perfil" value="${p.id}" id="perfil-${p.id}" ${p.id === estado.perfilId ? 'checked' : ''}>
        <strong>${esc(p.nombre)}</strong>
        <span class="destino">${esc(p.destino)}</span>
        <span class="detalle">${esc(p.descripcion)}
          <span class="origen"><b>Bug real:</b> ${esc(p.origen)}</span>
        </span>
      </label>`).join('');
  }

  function chips(c) {
    const out = [];
    const et = Plantilla.etiqueta;
    if (c.lista) out.push(`<span class="chip" title="${esc(c.lista.valores.join(' · '))}">Lista · ${c.lista.valores.length}</span>`);
    if (c.dependiente) out.push(`<span class="chip">Depende de ${esc(et(c.dependiente.padre))}</span>`);
    if (c.numero) out.push(`<span class="chip">${c.numero.entero ? 'Entero' : 'Número'}${c.numero.operador === 'between' ? ` ${c.numero.a}–${c.numero.b}` : ''}</span>`);
    if (c.largo) out.push(`<span class="chip">${c.largo.operador === 'between' ? `${c.largo.a}–${c.largo.b}` : c.largo.a} caracteres</span>`);
    if (c.fecha) out.push('<span class="chip">Fecha</span>');
    if (c.unico) out.push('<span class="chip">Sin repetidos</span>');
    c.inferidas.forEach(t => out.push(`<span class="chip tenue" title="Deducida por el nombre de la columna">${{ correo: 'Correo', telefono: '10 dígitos', fecha: 'Fecha' }[t]}?</span>`));
    if (c.noResueltas.length) out.push(`<span class="chip aviso-chip" title="${esc(c.noResueltas.join(' / '))}">Regla no reconocida</span>`);
    return out.join('') || '<span class="chip tenue">Texto libre</span>';
  }

  function pintarPlantilla() {
    const caja = $('#plantilla-actual');
    const p = estado.plantilla;
    $('#zona-plantilla').hidden = !!p;
    if (!p) { caja.innerHTML = ''; return; }
    const m = p.modelo;
    const perfil = m ? Plantilla.perfil(m, { obligatorios: estado.obligatorios }) : null; // siempre el de la plantilla, aunque se esté viendo Ejemplos
    const iHoja = p.hojas.findIndex(h => h.nombre === p.hoja);
    const selector = p.hojas.length > 1 ? `
      <label class="selector-hoja">
        <span>Pestaña con el formato</span>
        <select id="sel-hoja-plantilla" data-testid="sel-hoja-plantilla">
          ${iHoja < 0 ? '<option value="" selected disabled>Elige una pestaña</option>' : ''}${opcionesPestanas(p.hojas, iHoja)}
        </select>
      </label>` : '';
    const cabeza = `
      <div class="plantilla-cabeza">
        <div class="min0">
          ${p.esDemo ? '<span class="chip-muestra">Ejemplo</span> ' : ''}<span class="mono" data-testid="nombre-plantilla">${esc(p.nombre)}</span>
          ${m ? `<div class="sub">Encabezados en la fila ${m.filaEncabezados} · ${m.columnas.length} columnas</div>` : ''}
        </div>
        <button type="button" class="enlace" id="btn-quitar-plantilla" data-testid="btn-quitar-plantilla">Cambiar</button>
      </div>${selector}`;
    const conectar = () => {
      $('#btn-quitar-plantilla').addEventListener('click', quitarPlantilla);
      const sel = $('#sel-hoja-plantilla');
      if (sel) sel.addEventListener('change', () => {
        usarPlantilla(p.nombre, p.bytes, p.esDemo, p.hojas[+sel.value].nombre);
        if (estado.archivo) estado.archivo.hojaManual = false;
        validarActual();
      });
    };
    if (!m) {
      caja.innerHTML = cabeza + `<div class="aviso err" data-testid="error-plantilla">${esc(p.error)}</div>`;
      conectar();
      return;
    }
    caja.innerHTML = cabeza + `
      <div class="reglas" data-testid="reglas-detectadas">
        <div class="reglas-titulo">
          <span>Reglas detectadas</span>
          <span class="sub">${m.obligatoriosPorMarca ? 'Obligatorias: las marcadas con *' : 'Sin marcas de obligatorio: se tomaron todas'}</span>
        </div>
        ${m.columnas.map((c, i) => `
          <div class="regla" data-testid="regla-${i}">
            <label class="regla-col">
              <input type="checkbox" data-obl="${esc(c.nombre)}" data-testid="obl-${i}" ${perfil.columnas[i].obligatorio ? 'checked' : ''}>
              <span title="Marcar si la columna es obligatoria">${esc(c.nombre)}</span>
            </label>
            <div class="chips">${chips(c)}</div>
          </div>`).join('')}
      </div>`;
    conectar();
    caja.querySelectorAll('[data-obl]').forEach(chk => chk.addEventListener('change', () => {
      estado.obligatorios[chk.dataset.obl] = chk.checked;
      revisarMuestraPlantilla();
      validarActual();
    }));
  }

  function pintarArchivo() {
    const a = estado.archivo;
    const bloqueado = estado.modo === 'plantilla' && !estado.plantilla;
    $('#zona').classList.toggle('bloqueada', bloqueado);
    $('#input-archivo').disabled = bloqueado;
    $('#zona-texto').textContent = bloqueado ? 'Primero sube la plantilla oficial' : 'Arrastra tu Excel o CSV aquí';
    const varias = a && a.pestanas && a.pestanas.length > 1;
    $('#archivo-actual').innerHTML = a
      ? `<div class="archivo-nombre">${a.esMuestra ? '<span class="chip-muestra">Ejemplo</span>' : ''}<span class="mono" data-testid="nombre-archivo">${esc(a.nombre)}</span></div>
         ${varias ? `<label class="selector-hoja">
           <span>Pestaña a revisar${a.hojaManual ? '' : ' <em>(elegida por sus encabezados)</em>'}</span>
           <select id="sel-hoja-archivo" data-testid="sel-hoja-archivo">${opcionesPestanas(a.pestanas, a.hoja)}</select>
         </label>` : ''}`
      : '';
    const sel = $('#sel-hoja-archivo');
    if (sel) sel.addEventListener('change', () => {
      a.hoja = +sel.value; a.hojaManual = true; estado.filtro = 'todas';
      validarActual();
    });
    $('#btn-limpiar').disabled = !a && !estado.mensaje && !estado.plantilla;
  }

  // ---------------------------------------------------------------- Pintado: resultados
  function bloqueRevisionPlantilla() {
    if (estado.modo !== 'plantilla' || !estado.plantilla || !estado.plantilla.modelo) return '';
    const m = estado.plantilla.modelo;
    const items = [];
    const rm = estado.revisionMuestra;
    if (esAcumulado()) items.push(`<li><b>Esta plantilla ya trae datos</b> (${m.muestra.length >= 50 ? 'más de 50' : m.muestra.length} filas): se toma como archivo acumulado. Con "Agregar a mi archivo" las filas nuevas se escriben debajo de la última.</li>`);
    if (rm && !rm.fatal) {
      rm.filas.filter(f => f.estado === 'error').forEach(f => {
        f.incidencias.filter(i => i.severidad === 'error').forEach(i => items.push(
          `<li><b>La fila de muestra ${f.fila} de la propia plantilla no pasa sus reglas</b> en ${esc(Plantilla.etiqueta(i.columna))}: ${esc(i.motivo)} <span class="como">Si copias la muestra tal cual, el portal la va a rechazar.</span></li>`));
      });
    }
    m.revision.forEach(h => items.push(`<li><b>${esc(Plantilla.etiqueta(h.columna))}:</b> ${esc(h.motivo)} <span class="como">${esc(h.correccion)}</span></li>`));
    m.avisos.forEach(a => items.push(`<li>${esc(a)}</li>`));
    if (!items.length) return `<div class="aviso ok" data-testid="revision-plantilla">La plantilla oficial no tiene inconsistencias propias.</div>`;
    return `<details class="aviso warn" data-testid="revision-plantilla" open>
      <summary><strong>Revisión de la plantilla oficial · ${items.length} ${items.length === 1 ? 'hallazgo' : 'hallazgos'}</strong></summary>
      <ul>${items.join('')}</ul></details>`;
  }

  function pintarResultados() {
    const cont = $('#resultados');
    const r = estado.resultado;
    const perfil = perfilActual();
    const msg = estado.mensaje
      ? `<div class="aviso ${estado.mensajeError ? 'err' : 'warn'}" data-testid="mensaje">${esc(estado.mensaje)}</div>` : '';
    const revision = bloqueRevisionPlantilla();

    if (!perfil) {
      cont.innerHTML = msg + `<div class="vacio-tabla" data-testid="estado-vacio">
        <strong>Sube la plantilla oficial del portal.</strong><br>
        Es el Excel que el portal te da para llenar. ValidaCarga lee de ahí las columnas y las listas desplegables, y con eso revisa tu archivo.
        Tu información no sale de esta computadora.</div>`;
      return;
    }
    if (!r || estado.mensajeError) {
      cont.innerHTML = `<div class="titulo-res"><h2>${esc(perfil.nombre)}</h2></div>` + msg + revision +
        `<div class="vacio-tabla" data-testid="estado-vacio"><strong>Listo para un archivo nuevo.</strong><br>Sube tu Excel o CSV para revisarlo${estado.modo === 'plantilla' ? ' contra la plantilla' : `: ${esc(perfil.nombre)}`}.</div>`;
      return;
    }

    const s = r.resumen;
    const pct = n => s.total ? (n / s.total * 100).toFixed(2) : 0;
    const errArchivo = r.archivo.filter(i => i.severidad === 'error');
    const warnArchivo = r.archivo.filter(i => i.severidad !== 'error');
    const avisoArchivo = (lista, clase, titulo) => lista.length ? `
      <div class="aviso ${clase}" data-testid="aviso-archivo-${clase}">
        <strong>${titulo}</strong>
        <ul>${lista.map(i => `<li>${esc(i.motivo)} <span class="como">${esc(i.correccion)}</span></li>`).join('')}</ul>
      </div>` : '';
    const filtros = [['todas', 'Todas', s.total], ['error', 'Con error', s.error], ['advertencia', 'Con advertencia', s.advertencia], ['correcto', 'Correctas', s.correcto]];
    const pre = estado.preparado;

    cont.innerHTML = `
      <div class="titulo-res">
        <h2>${esc(perfil.nombre)}</h2>
        <p>${estado.archivo.pestanas && estado.archivo.pestanas.length > 1 ? `Pestaña "${esc(estado.archivo.pestanas[estado.archivo.hoja].nombre)}" · ` : ''}${s.total.toLocaleString('es-MX')} filas revisadas${pre.filaEncabezados > 1 ? ` · encabezados en la fila ${pre.filaEncabezados}` : ''}</p>
      </div>
      ${msg}${revision}
      ${avisoArchivo(errArchivo, 'err', 'Faltan columnas obligatorias. Se revisaron solo las columnas que sí trae tu archivo; descarga el archivo marcado para ver en rojo lo que falta.')}
      ${avisoArchivo(warnArchivo, 'warn', 'Observaciones en los encabezados')}
      ${`
      <div class="marcador" data-testid="resumen">
        <div class="cifra"><b data-testid="resumen-total">${s.total}</b><span>Filas</span></div>
        <div class="cifra ok"><b data-testid="resumen-correcto">${s.correcto}</b><span>Correctas</span></div>
        <div class="cifra warn"><b data-testid="resumen-advertencia">${s.advertencia}</b><span>Con advertencia</span></div>
        <div class="cifra err"><b data-testid="resumen-error">${s.error}</b><span>Con error</span></div>
      </div>
      <div class="barra" role="img" aria-label="${s.correcto} correctas, ${s.advertencia} con advertencia, ${s.error} con error">
        <i style="width:${pct(s.correcto)}%;background:var(--ok)"></i><i style="width:${pct(s.advertencia)}%;background:var(--warn)"></i><i style="width:${pct(s.error)}%;background:var(--err)"></i>
      </div>
      <div class="descargas">
        <button class="boton principal" type="button" id="btn-reporte" data-testid="btn-descargar-reporte">Descargar reporte de errores</button>
        ${estado.modo === 'plantilla'
          ? `<button class="boton" type="button" id="btn-marcado" data-testid="btn-descargar-base" title="Tu plantilla o archivo acumulado con las filas nuevas abajo: lo que pasó queda normal y lo que falló en rojo con nota. La revisión va en pestañas nuevas.">Agregar a mi archivo</button>`
          : `<button class="boton" type="button" id="btn-marcado" data-testid="btn-descargar-marcado" title="Excel con las columnas del formato, los errores en rojo y una nota con el motivo">Descargar archivo marcado</button>`}
        <button class="boton" type="button" id="btn-validas" data-testid="btn-descargar-validas" ${s.validas && !r.fatal ? '' : 'disabled'} ${r.fatal ? 'title="No disponible: faltan columnas obligatorias"' : ''}>Descargar filas válidas (${r.fatal ? 0 : s.validas})</button>
      </div>
      <div class="estado-msg" id="estado-descarga" data-testid="estado-descarga" role="status"></div>
      <div class="filtros" role="group" aria-label="Filtrar filas">
        ${filtros.map(([id, txt, n]) => `<button type="button" class="filtro" data-filtro="${id}" data-testid="filtro-${id}" aria-pressed="${estado.filtro === id}">${txt}<span class="n">${n}</span></button>`).join('')}
      </div>
      ${tabla(r)}`}
    `;
    cont.querySelectorAll('.filtro').forEach(b => b.addEventListener('click', () => { estado.filtro = b.dataset.filtro; pintarResultados(); }));
    const br = $('#btn-reporte'); if (br) br.addEventListener('click', descargarReporte);
    const bv = $('#btn-validas'); if (bv) bv.addEventListener('click', descargarValidas);
    const bm = $('#btn-marcado'); if (bm) bm.addEventListener('click', descargarMarcado);
  }

  function tabla(r) {
    const filas = r.filas.filter(f => estado.filtro === 'todas' || f.estado === estado.filtro);
    if (!filas.length) return '<div class="tabla-caja"><p class="vacio-tabla">No hay filas en esta vista.</p></div>';
    const visibles = filas.slice(0, MAX_FILAS_TABLA);
    const etiqueta = { correcto: 'Correcta', advertencia: 'Advertencia', error: 'Error' };
    const cuerpo = visibles.map(f => {
      const inc = f.incidencias.length ? f.incidencias : [null];
      return `<tbody class="grupo" data-testid="fila-${f.fila}" data-fila="${f.fila}" data-estado="${f.estado}">` + inc.map((i, k) => `
        <tr>
          ${k === 0 ? `<td class="num" rowspan="${inc.length}">${f.fila}</td><td rowspan="${inc.length}"><span class="estado ${f.estado}">${etiqueta[f.estado]}</span></td>` : ''}
          ${i ? `
            <td class="col">${esc(i.columna)}<br><span class="sev ${i.severidad}">${i.severidad}</span></td>
            <td class="val">${i.valor === '' ? '<em>(vacío)</em>' : esc(mostrarEspacios(i.valor))}</td>
            <td>${esc(i.motivo)}</td>
            <td class="como">${esc(i.correccion)}</td>` : `
            <td class="col">—</td><td class="val"><em>—</em></td><td>Sin observaciones.</td><td class="como">Lista para cargar.</td>`}
        </tr>`).join('') + '</tbody>';
    }).join('');
    const nota = filas.length > MAX_FILAS_TABLA
      ? `<p class="nota-tabla">Se muestran las primeras ${MAX_FILAS_TABLA} filas de ${filas.length.toLocaleString('es-MX')}. El reporte descargable incluye todas.</p>` : '';
    return `<div class="tabla-caja"><table data-testid="tabla-resultados">
      <thead><tr><th>Fila</th><th>Estado</th><th>Columna</th><th>Valor</th><th>Motivo</th><th>Cómo corregirlo</th></tr></thead>
      ${cuerpo}</table></div>${nota}`;
  }

  /** Hace visibles los espacios al inicio/fin para que se entienda el error. */
  const mostrarEspacios = v => String(v).replace(/^ +| +$/g, m => '·'.repeat(m.length));

  // ---------------------------------------------------------------- Descargas
  const base = () => estado.archivo.nombre.replace(/\.[^.]+$/, '');

  function avisarDescarga(texto, error) {
    const el = $('#estado-descarga');
    if (el) { el.textContent = texto; el.classList.toggle('err', !!error); }
  }

  function descargar(fn, nombre) {
    try { fn(); avisarDescarga(`Descargado: ${nombre}`); }
    catch (e) { avisarDescarga('Este navegador bloqueó la descarga. Abre ValidaCarga en una pestaña normal e inténtalo de nuevo.', true); }
  }

  function guardarXlsx(aoa, hoja, nombre, anchos) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (anchos) ws['!cols'] = anchos.map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, hoja);
    XLSX.writeFile(wb, nombre);
  }

  function guardarCsv(aoa, nombre) {
    const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(aoa));
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }); // BOM para que Excel respete acentos
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function descargarReporte() {
    const nombre = `${base()}_reporte_errores.xlsx`;
    descargar(() => guardarXlsx(Validador.filasReporte(estado.resultado), 'Reporte', nombre, [10, 14, 12, 26, 28, 70, 60]), nombre);
  }

  function descargarValidas() {
    const aoa = Validador.filasValidas(estado.resultado, estado.preparado.encabezados);
    const nombre = `${base()}_filas_validas.${estado.archivo.tipo}`;
    descargar(() => estado.archivo.tipo === 'csv' ? guardarCsv(aoa, nombre) : guardarXlsx(aoa, 'Datos', nombre), nombre);
  }

  // ---------------------------------------------------------------- Archivo marcado (colores con ExcelJS)
  const EXCELJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
  let cargaExcelJS = null;

  /** Carga ExcelJS solo cuando se necesita (SheetJS gratuito no escribe colores de celda). */
  function cargarExcelJS() {
    if (window.ExcelJS) return Promise.resolve();
    if (!cargaExcelJS) {
      cargaExcelJS = new Promise((ok, falla) => {
        const sc = document.createElement('script');
        sc.src = EXCELJS_URL;
        sc.onload = () => window.ExcelJS ? ok() : falla(new Error('sin ExcelJS'));
        sc.onerror = () => { cargaExcelJS = null; falla(new Error('sin conexión')); };
        document.head.appendChild(sc);
      });
    }
    return cargaExcelJS;
  }

  const COLOR = {
    encabezado: 'FF155F78', faltante: 'FFC0392B', blanco: 'FFFFFFFF',
    error: 'FFF8D3CF', errorTexto: 'FF8E1B13', advertencia: 'FFFCEBC6', advertenciaTexto: 'FF6B4500',
    vacioOpcional: 'FFF0F3F4', correcto: 'FFDDF0E4', correctoTexto: 'FF1F5C3A'
  };
  const relleno = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

  async function descargarMarcado() {
    const boton = $('#btn-marcado');
    boton.disabled = true;
    avisarDescarga('Preparando el archivo marcado…');
    try {
      await cargarExcelJS();
    } catch (e) {
      boton.disabled = false;
      avisarDescarga('No se pudo cargar el generador de Excel con colores. Revisa tu conexión a internet e inténtalo de nuevo.', true);
      return;
    }
    try {
      const acumular = estado.modo === 'plantilla';
      const nombre = acumular ? `${estado.plantilla.nombre.replace(/\.[^.]+$/, '')}_actualizado.xlsx` : `${base()}_marcado.xlsx`;
      const buffer = acumular ? await construirBaseActualizada() : await construirMarcado();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob); a.download = nombre;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 1500);
      avisarDescarga(`Descargado: ${nombre}`);
    } catch (e) {
      avisarDescarga('Este navegador bloqueó la descarga. Abre ValidaCarga en una pestaña normal e inténtalo de nuevo.', true);
    } finally {
      boton.disabled = false;
    }
  }

  /**
   * Hoja "Datos": columnas en el orden de la plantilla, llenas con lo compatible del archivo.
   * Rojo = error o columna que tu archivo no trae · ámbar = advertencia. Cada celda marcada lleva una nota con el motivo.
   * Hoja "Resumen": totales, columnas faltantes/no incluidas y la leyenda de colores.
   */
  async function construirMarcado() {
    const perfil = perfilActual();
    const r = estado.resultado;
    const pre = estado.preparado;
    const cols = perfil.columnas;
    const libro = new ExcelJS.Workbook();
    libro.creator = 'ValidaCarga';
    const hoja = libro.addWorksheet('Datos', { views: [{ state: 'frozen', ySplit: 1 }] });

    const extra = ['Fila original', 'Estado', 'Observaciones'];
    const encabezado = hoja.addRow(cols.map(c => c.nombre).concat(extra));
    encabezado.font = { bold: true, color: { argb: COLOR.blanco } };
    encabezado.eachCell((cel, n) => {
      const col = cols[n - 1];
      const falta = col && r.faltantes.includes(col.nombre);
      cel.fill = relleno(falta ? COLOR.faltante : COLOR.encabezado);
      cel.alignment = { vertical: 'middle', wrapText: true };
      if (falta) cel.note = `Tu archivo no trae esta columna${col.obligatorio ? ' y es obligatoria' : ''}. Se dejó vacía.`;
    });
    encabezado.height = 30;

    const etiqueta = { correcto: 'Correcta', advertencia: 'Advertencia', error: 'Error', incompleta: 'Incompleta' };
    r.filas.forEach(f => {
      const porCol = {};
      const sueltas = [];
      f.incidencias.forEach(i => {
        if (i.columna && cols.some(c => c.nombre === i.columna)) (porCol[i.columna] = porCol[i.columna] || []).push(i);
        else sueltas.push(i);
      });
      const valores = cols.map(c => f.datos[c.nombre]);
      // Si faltan columnas obligatorias, la fila no se puede cargar aunque lo que trae esté bien
      const oblFaltantes = cols.filter(c => c.obligatorio && r.faltantes.includes(c.nombre)).map(c => c.nombre);
      const estadoFila = f.estado === 'error' ? 'error' : (oblFaltantes.length ? 'incompleta' : f.estado);
      const obs = f.incidencias.map(i => `${i.columna ? i.columna + ': ' : ''}${i.motivo}`)
        .concat(oblFaltantes.length ? [`Falta llenar: ${oblFaltantes.join(', ')}`] : []).join(' | ');
      const fila = hoja.addRow(valores.concat([f.fila, etiqueta[estadoFila], obs]));
      cols.forEach((c, k) => {
        const cel = fila.getCell(k + 1);
        cel.numFmt = '@'; // conserva el texto tal cual (ceros a la izquierda, números largos)
        const inc = porCol[c.nombre];
        if (r.faltantes.includes(c.nombre)) {
          cel.fill = relleno(c.obligatorio ? COLOR.error : COLOR.vacioOpcional);
          if (c.obligatorio && f === r.filas[0]) cel.note = `Columna obligatoria que no viene en tu archivo. Hay que llenarla.`;
        } else if (inc) {
          const esError = inc.some(i => i.severidad === 'error');
          cel.fill = relleno(esError ? COLOR.error : COLOR.advertencia);
          cel.font = { color: { argb: esError ? COLOR.errorTexto : COLOR.advertenciaTexto }, bold: esError };
          cel.note = inc.map(i => `${i.motivo}\nCómo corregirlo: ${i.correccion}`).join('\n\n');
        }
      });
      const celEstado = fila.getCell(cols.length + 2);
      const tono = { correcto: [COLOR.correcto, COLOR.correctoTexto], advertencia: [COLOR.advertencia, COLOR.advertenciaTexto], incompleta: [COLOR.advertencia, COLOR.advertenciaTexto], error: [COLOR.error, COLOR.errorTexto] }[estadoFila];
      celEstado.fill = relleno(tono[0]);
      celEstado.font = { bold: true, color: { argb: tono[1] } };
      fila.getCell(cols.length + 3).alignment = { wrapText: true, vertical: 'top' };
    });

    hoja.columns.forEach((columna, k) => {
      const nombre = k < cols.length ? cols[k].nombre : extra[k - cols.length];
      columna.width = nombre === 'Observaciones' ? 70 : Math.min(40, Math.max(12, nombre.length + 4));
    });
    hoja.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length + extra.length } };

    // Resumen
    const res = libro.addWorksheet('Resumen');
    res.columns = [{ width: 42 }, { width: 70 }];
    const titulo = res.addRow(['ValidaCarga · archivo marcado']);
    titulo.font = { bold: true, size: 14 };
    res.addRow([]);
    const nuevas = (pre.encabezados || []).filter((h, i) => String(h).trim() && !Object.values(r.indice).includes(i));
    const s = r.resumen;
    [
      ['Archivo revisado', estado.archivo.nombre + (estado.archivo.pestanas && estado.archivo.pestanas.length > 1 ? ` (pestaña "${estado.archivo.pestanas[estado.archivo.hoja].nombre}")` : '')],
      ['Revisado contra', estado.modo === 'plantilla' ? `Plantilla ${estado.plantilla.nombre}` : `Ejemplo: ${perfil.nombre}`],
      ['Generado', new Date().toLocaleString('es-MX')],
      ['Filas', s.total], ['Correctas', s.correcto], ['Con advertencia', s.advertencia], ['Con error', s.error],
      ['Columnas de la plantilla que tu archivo no trae', r.faltantes.join(', ') || 'Ninguna'],
      ['Columnas de tu archivo que no se incluyeron', nuevas.join(', ') || 'Ninguna']
    ].forEach(([k, v]) => {
      const fila = res.addRow([k, v]);
      fila.getCell(1).font = { bold: true };
      fila.getCell(2).alignment = { wrapText: true, horizontal: 'left' };
    });
    res.addRow([]);
    res.addRow(['Colores']).font = { bold: true };
    [
      [COLOR.faltante, 'Encabezado rojo: la columna no viene en tu archivo'],
      [COLOR.error, 'Celda roja: error (o columna obligatoria que no viene). Pasa el mouse para ver el motivo'],
      [COLOR.advertencia, 'Celda ámbar: advertencia, revisar antes de subir'],
      [COLOR.vacioOpcional, 'Celda gris: columna opcional que tu archivo no trae'],
      [COLOR.advertencia, 'Estado "Incompleta": lo que trae la fila está bien, pero faltan columnas obligatorias por llenar']
    ].forEach(([c, t]) => {
      const fila = res.addRow(['', t]);
      fila.getCell(1).fill = relleno(c);
    });
    return libro.xlsx.writeBuffer();
  }

  /** Convierte el texto validado al tipo de la columna en la base (número o fecha), para que se integre a lo que ya hay. */
  function convertirValor(v, referencia) {
    const t = String(v == null ? '' : v);
    if (t.trim() === '') return null;
    const ref = referencia && typeof referencia === 'object' && !(referencia instanceof Date) && 'result' in referencia ? referencia.result : referencia;
    const f = Validador.leerFecha(t);
    if (f && (ref instanceof Date || ref == null)) return f.fecha; // fecha real de Excel
    if (typeof ref === 'number' || ref == null) {
      const n = Validador.leerMonto(t);
      const ceroIzq = /^0\d/.test(t.trim());
      const largo = t.replace(/\D/g, '').length > 15;
      if (n && !ceroIzq && !largo && (typeof ref === 'number' || /^-?\d+(\.\d+)?$/.test(t.trim()))) return n.numero;
    }
    return t;
  }

  /**
   * ExcelJS guarda las validaciones celda por celda y al reescribirlas genera rangos duplicados
   * (ordena "G10" antes que "G2"). Aquí se vuelven a juntar en rangos continuos por columna.
   */
  function compactarValidaciones(ws) {
    const modelo = ws.dataValidations && ws.dataValidations.model;
    if (!modelo) return;
    const grupos = new Map();
    Object.entries(modelo).forEach(([dir, v]) => {
      const m = dir.match(/^([A-Z]+)(\d+)$/);
      if (!m) { grupos.set('rango|' + dir, { rango: dir, v }); return; }
      const k = JSON.stringify(v) + '|' + m[1];
      if (!grupos.has(k)) grupos.set(k, { col: m[1], v, filas: [] });
      grupos.get(k).filas.push(+m[2]);
    });
    const nuevo = {};
    grupos.forEach(g => {
      if (g.rango) { nuevo[g.rango] = g.v; return; }
      g.filas.sort((a, b) => a - b);
      let ini = g.filas[0], prev = ini;
      const cerrar = () => { nuevo[ini === prev ? `${g.col}${ini}` : `${g.col}${ini}:${g.col}${prev}`] = g.v; };
      g.filas.slice(1).forEach(f => { if (f === prev + 1) prev = f; else { cerrar(); ini = prev = f; } });
      cerrar();
    });
    ws.dataValidations.model = nuevo;
  }

  function copiarEstilo(estilo) { return estilo ? JSON.parse(JSON.stringify(estilo)) : {}; }

  /**
   * Abre la plantilla (o archivo acumulado) original y agrega abajo las filas de tu archivo, solo en las columnas compatibles.
   * Conserva las demás pestañas. Agrega "ValidaCarga Revisión" y "ValidaCarga Resumen".
   */
  async function construirBaseActualizada() {
    const p = estado.plantilla;
    const m = p.modelo;
    const r = estado.resultado;
    const perfil = Plantilla.perfil(m, { obligatorios: estado.obligatorios });
    const oblig = Object.fromEntries(perfil.columnas.map(c => [c.nombre, c.obligatorio]));
    const libro = new ExcelJS.Workbook();
    await libro.xlsx.load(p.bytes);
    const hoja = libro.getWorksheet(m.hoja);
    if (!hoja) throw new Error('No encontré la pestaña del formato en la plantilla.');

    // Última fila con datos en las columnas que se van a llenar
    const escribibles = m.columnas.filter(c => r.indice[c.nombre] !== undefined);
    const conValor = v => v !== null && v !== undefined && String(typeof v === 'object' && 'result' in v ? v.result : v).trim() !== '';
    let ultima = m.filaEncabezados;
    hoja.eachRow({ includeEmpty: false }, (fila, n) => {
      if (n > ultima && escribibles.some(c => conValor(fila.getCell(c.indice + 1).value))) ultima = n;
    });
    const modelo = ultima > m.filaEncabezados ? hoja.getRow(ultima) : null;
    const primeraNueva = ultima + 1;

    const revision = [];
    r.filas.forEach((f, k) => {
      const n = primeraNueva + k;
      const fila = hoja.getRow(n);
      const porCol = {};
      f.incidencias.forEach(i => { (porCol[i.columna] = porCol[i.columna] || []).push(i); });
      m.columnas.forEach(c => {
        const cel = fila.getCell(c.indice + 1);
        const ref = modelo ? modelo.getCell(c.indice + 1) : null;
        const presente = r.indice[c.nombre] !== undefined;
        if (!presente && !(oblig[c.nombre])) return; // columnas que tu archivo no trae: no se tocan (pueden tener fórmulas)
        if (ref && ref.style) cel.style = copiarEstilo(ref.style);
        if (presente) cel.value = convertirValor(f.datos[c.nombre], ref ? ref.value : null);
        if (cel.value instanceof Date && !cel.numFmt) cel.numFmt = 'dd/mm/yyyy';
        const inc = porCol[c.nombre];
        if (!presente) {
          cel.fill = relleno(COLOR.error);
          cel.note = 'Falta: tu archivo no trae esta columna y es obligatoria.';
        } else if (inc) {
          const esError = inc.some(i => i.severidad === 'error');
          cel.fill = relleno(esError ? COLOR.error : COLOR.advertencia);
          cel.font = Object.assign({}, cel.font || {}, { color: { argb: esError ? 'FFC0392B' : COLOR.advertenciaTexto }, bold: esError });
          cel.note = inc.map(i => `${i.motivo}\nCómo corregirlo: ${i.correccion}`).join('\n\n');
        }
      });
      f.incidencias.forEach(i => revision.push([n, f.fila, { error: 'Error', advertencia: 'Advertencia' }[i.severidad], i.columna, i.valor, i.motivo, i.correccion]));
      r.faltantes.filter(c => oblig[c]).forEach(c => revision.push([n, f.fila, 'Falta', c, '', 'Tu archivo no trae esta columna obligatoria.', 'Llenarla en la base.']));
    });

    // Pestañas de ValidaCarga (se reemplazan si ya existían de una corrida anterior)
    ['ValidaCarga Revisión', 'ValidaCarga Resumen'].forEach(nom => { const h = libro.getWorksheet(nom); if (h) libro.removeWorksheet(h.id); });

    const rev = libro.addWorksheet('ValidaCarga Revisión', { views: [{ state: 'frozen', ySplit: 1 }] });
    rev.columns = [
      { header: 'Fila en la base', width: 14 }, { header: 'Fila en tu archivo', width: 16 }, { header: 'Tipo', width: 13 },
      { header: 'Columna', width: 24 }, { header: 'Valor', width: 26 }, { header: 'Motivo', width: 70 }, { header: 'Cómo corregirlo', width: 55 }
    ];
    rev.getRow(1).font = { bold: true, color: { argb: COLOR.blanco } };
    rev.getRow(1).eachCell(c => { c.fill = relleno(COLOR.encabezado); });
    revision.forEach(fila => {
      const nueva = rev.addRow(fila);
      const tipo = nueva.getCell(3);
      tipo.fill = relleno(fila[2] === 'Advertencia' ? COLOR.advertencia : COLOR.error);
      nueva.getCell(6).alignment = { wrapText: true, vertical: 'top' };
      nueva.getCell(7).alignment = { wrapText: true, vertical: 'top' };
    });
    if (!revision.length) rev.addRow(['', '', 'Sin observaciones: todas las filas nuevas pasaron.']);
    rev.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 7 } };

    const res = libro.addWorksheet('ValidaCarga Resumen');
    res.columns = [{ width: 44 }, { width: 70 }];
    res.addRow(['ValidaCarga · filas agregadas']).font = { bold: true, size: 14 };
    res.addRow([]);
    const s = r.resumen;
    const nuevas = (estado.preparado.encabezados || []).filter((h, i) => String(h).trim() && !Object.values(r.indice).includes(i));
    [
      ['Archivo de origen', estado.archivo.nombre + (estado.archivo.pestanas && estado.archivo.pestanas.length > 1 ? ` (pestaña "${estado.archivo.pestanas[estado.archivo.hoja].nombre}")` : '')],
      ['Pestaña donde se agregaron', `${m.hoja}, filas ${primeraNueva} a ${primeraNueva + r.filas.length - 1}`],
      ['Generado', new Date().toLocaleString('es-MX')],
      ['Filas agregadas', s.total], ['Sin observaciones', s.correcto], ['Con advertencia', s.advertencia], ['Con error', s.error],
      ['Columnas obligatorias que tu archivo no trae', r.faltantes.filter(c => oblig[c]).join(', ') || 'Ninguna'],
      ['Columnas de tu archivo que no se agregaron', nuevas.join(', ') || 'Ninguna']
    ].forEach(([k, v]) => {
      const fila = res.addRow([k, v]);
      fila.getCell(1).font = { bold: true };
      fila.getCell(2).alignment = { wrapText: true, horizontal: 'left' };
    });
    res.addRow([]);
    res.addRow(['Colores en las filas agregadas']).font = { bold: true };
    [[COLOR.error, 'Rojo: la celda falló (o es obligatoria y no venía). Pasa el mouse para ver el motivo'],
     [COLOR.advertencia, 'Ámbar: advertencia, revisar antes de usar']].forEach(([c, t]) => { res.addRow(['', t]).getCell(1).fill = relleno(c); });

    // Reescribe en formato clásico las listas modernas (x14) que ExcelJS no conserva
    (m.validacionesX14 || []).forEach(v => {
      const modelo = hoja.dataValidations.model;
      String(v.sqref).split(/\s+/).filter(Boolean).forEach(rango => {
        const r1 = rango.split(':')[0];
        if (modelo[r1] || modelo[rango]) return; // ya existe
        const regla = { type: v.type, allowBlank: v.allowBlank, showErrorMessage: true, formulae: [v.f1.replace(/^=/, '')].concat(v.f2 ? [v.f2] : []) };
        if (v.operator) regla.operator = v.operator;
        modelo[rango] = regla;
      });
    });
    libro.worksheets.forEach(compactarValidaciones);
    return libro.xlsx.writeBuffer();
  }

  function descargarPlantillaEjemplo() {
    const p = Perfiles.porId[estado.perfilId];
    try { guardarXlsx([p.muestra.encabezados, p.muestra.filas[0]], 'Datos', `plantilla_${p.id}.xlsx`, p.muestra.encabezados.map(h => Math.max(14, h.length + 4))); }
    catch (e) { /* descarga bloqueada por el navegador */ }
  }

  // ---------------------------------------------------------------- Eventos
  function zonaArrastre(zona, alSoltar) {
    zona.tabIndex = -1;
    ['dragenter', 'dragover'].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); if (!zona.classList.contains('bloqueada')) zona.classList.add('arrastrando'); }));
    ['dragleave', 'drop'].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); zona.classList.remove('arrastrando'); }));
    zona.addEventListener('drop', e => { if (!zona.classList.contains('bloqueada')) alSoltar(e.dataTransfer.files[0]); });
  }

  function iniciar() {
    pintarPerfiles();

    document.querySelectorAll('[data-modo]').forEach(b => b.addEventListener('click', () => {
      if (estado.modo === b.dataset.modo) return;
      estado.modo = b.dataset.modo;
      estado.filtro = 'todas';
      if (estado.archivo && estado.archivo.esMuestra) {
        if (estado.modo === 'ejemplos') cargarMuestraEjemplo();
        else if (estado.plantilla && estado.plantilla.esDemo) cargarDemo();
        else { estado.archivo = null; validarActual(); }
      } else validarActual();
    }));

    $('#perfiles').addEventListener('change', e => {
      estado.perfilId = e.target.value;
      if (estado.archivo && estado.archivo.esMuestra) cargarMuestraEjemplo();
      else validarActual();
    });

    const inPlantilla = $('#input-plantilla');
    inPlantilla.addEventListener('change', () => { recibirPlantilla(inPlantilla.files[0]); inPlantilla.value = ''; });
    zonaArrastre($('#zona-plantilla'), recibirPlantilla);

    const input = $('#input-archivo');
    input.addEventListener('change', () => { recibirArchivo(input.files[0]); input.value = ''; });
    zonaArrastre($('#zona'), recibirArchivo);

    $('#btn-muestra').addEventListener('click', () => estado.modo === 'plantilla' ? cargarDemo() : cargarMuestraEjemplo());
    $('#btn-plantilla').addEventListener('click', descargarPlantillaEjemplo);
    $('#btn-limpiar').addEventListener('click', limpiar);

    cargarDemo(); // abre con la plantilla de ejemplo para que se vea qué hace
  }

  if (typeof XLSX === 'undefined') {
    document.getElementById('resultados').innerHTML =
      '<div class="aviso err">No se pudo cargar el lector de Excel. Revisa tu conexión a internet y recarga la página.</div>';
    return;
  }
  iniciar();
})();
