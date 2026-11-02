/*
 * ValidaCarga — perfiles de portal destino.
 * Cada perfil define sus columnas, sus reglas por fila y un archivo de muestra con errores sembrados.
 * Todos los datos son ficticios.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Perfiles = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Catálogos
  // ---------------------------------------------------------------------------

  /** Región → Estado → Ciudad → [Municipios]. Coincidencia exacta, como en el portal. */
  const GEOGRAFIA = {
    'Región 5': {
      'Jalisco': {
        'Guadalajara': ['Guadalajara', 'Zapopan', 'San Pedro Tlaquepaque', 'Tonalá'],
        'Puerto Vallarta': ['Puerto Vallarta']
      },
      'Colima': {
        'Colima': ['Colima', 'Villa de Álvarez'],
        'Manzanillo': ['Manzanillo'],
        'Tecomán': ['Tecomán']
      },
      'Nayarit': { 'Tepic': ['Tepic', 'Xalisco'] },
      'Michoacán': { 'Morelia': ['Morelia'], 'Uruapan': ['Uruapan'] }
    },
    'Región 9': {
      'Ciudad de México': { 'Ciudad de México': ['Benito Juárez', 'Coyoacán', 'Cuauhtémoc', 'Iztapalapa'] },
      'Estado de México': { 'Toluca': ['Toluca', 'Metepec'], 'Naucalpan': ['Naucalpan de Juárez'] }
    }
  };
  const TIPO_ALUMNO = ['Distribuidor', 'Empleado', 'Externo'];
  const TIPO_CANAL = ['Detalle de Canal', 'Cadenas Comerciales', 'Masivo'];

  const BOLSA = ['Agente', 'Master'];
  const DENOMINACIONES = [10, 20, 30, 50, 80, 100, 150, 200, 300, 500];
  const ESTATUS_RECARGA = ['Exitosa', 'Fallida', 'Cancelada'];

  const TIPO_MOVIMIENTO = ['Portabilidad Entrante', 'Portabilidad Saliente', 'Reversión'];
  const PROVEEDORES = ['Telcel', 'AT&T', 'Movistar', 'Bait', 'Wimo', 'Virgin Mobile'];

  // ---------------------------------------------------------------------------
  // Ayudas compartidas
  // ---------------------------------------------------------------------------

  /** Revisa un catálogo y registra el error con sugerencia si el valor solo "se parece". */
  function catalogo(api, columna, valor, opciones, contexto) {
    const u = api.ctx.utiles;
    if (u.vacio(valor)) return false;
    const r = u.revisarCatalogo(valor, opciones);
    if (r.ok) return true;
    if (r.sugerencia) {
      api.agregar(api.SEV.ERROR, columna,
        `Se parece a "${r.sugerencia}", pero no es idéntico (acentos, mayúsculas o espacios). El portal compara texto exacto y rechaza la fila.`,
        `Escribir exactamente "${r.sugerencia}".`);
    } else {
      api.agregar(api.SEV.ERROR, columna,
        `"${valor}" no existe en el catálogo${contexto ? ' ' + contexto : ''}.`,
        `Usar uno de: ${opciones.slice(0, 6).join(', ')}${opciones.length > 6 ? '…' : ''}.`);
    }
    return false;
  }

  /** Registra duplicados: la primera aparición pasa, las siguientes se marcan. */
  function duplicado(api, grupo, llave) {
    const v = api.ctx.vistos;
    v[grupo] = v[grupo] || {};
    if (v[grupo][llave] !== undefined) return v[grupo][llave];
    v[grupo][llave] = api.fila;
    return null;
  }

  /** Número de 10 dígitos. Devuelve true si es válido. */
  function telefono(api, columna, valor) {
    const u = api.ctx.utiles;
    if (u.vacio(valor)) return false;
    const s = String(valor).trim();
    if (/^\d{10}$/.test(s)) return true;
    const soloDigitos = s.replace(/\D/g, '');
    let correccion = 'Capturar el número a 10 dígitos, sin espacios, guiones ni lada internacional.';
    if (soloDigitos.length === 12 && soloDigitos.startsWith('52')) correccion = `Quitar el prefijo 52: ${soloDigitos.slice(2)}.`;
    else if (soloDigitos.length === 10) correccion = `Dejar solo dígitos: ${soloDigitos}.`;
    api.agregar(api.SEV.ERROR, columna,
      `"${s}" no es un número de 10 dígitos (tiene ${soloDigitos.length} dígitos${soloDigitos.length !== s.length ? ' y otros caracteres' : ''}).`,
      correccion);
    return false;
  }

  function fecha(api, columna, valor, { noFutura } = {}) {
    const u = api.ctx.utiles;
    if (u.vacio(valor)) return null;
    const f = u.leerFecha(valor);
    if (!f) {
      api.agregar(api.SEV.ERROR, columna, `"${valor}" no es una fecha válida.`,
        'Usar el formato DD/MM/AAAA con una fecha que exista (ej. 15/10/2026).');
      return null;
    }
    if (noFutura) {
      const h = api.ctx.hoy;
      const hoyUTC = Date.UTC(h.getFullYear(), h.getMonth(), h.getDate());
      if (f.fecha.getTime() > hoyUTC) {
        api.agregar(api.SEV.ERROR, columna, `La fecha ${f.texto} es posterior a hoy.`, 'Revisar el día, mes y año.');
      }
    }
    return f;
  }

  // ---------------------------------------------------------------------------
  // Perfil 1 — Alta masiva de usuarios (portal de capacitación)
  // ---------------------------------------------------------------------------
  const usuarios = {
    id: 'usuarios',
    nombre: 'Alta masiva de usuarios',
    destino: 'Portal de capacitación',
    descripcion: 'Campos obligatorios, catálogos en cascada (Región → Estado → Ciudad → Municipio) y usuarios duplicados.',
    origen: 'La plantilla oficial fallaba en los campos de catálogo, incluso con su propia fila de muestra.',
    columnas: [
      { nombre: 'Usuario', obligatorio: true },
      { nombre: 'Nombre', obligatorio: true },
      { nombre: 'Apellido Paterno', obligatorio: true },
      { nombre: 'Apellido Materno', obligatorio: false },
      { nombre: 'Correo', obligatorio: true },
      { nombre: 'Región', obligatorio: true, catalogo: true },
      { nombre: 'Estado', obligatorio: true, catalogo: true },
      { nombre: 'Ciudad', obligatorio: true, catalogo: true },
      { nombre: 'Municipio', obligatorio: true, catalogo: true },
      { nombre: 'Tipo de Alumno', obligatorio: true, catalogo: true },
      { nombre: 'Tipo de Canal', obligatorio: true, catalogo: true }
    ],
    validarFila(d, api) {
      const { SEV, agregar, ctx } = api;
      const u = ctx.utiles;

      // Usuario: 4 a 20 caracteres, letras y números
      if (!u.vacio(d['Usuario'])) {
        const usr = d['Usuario'].trim();
        if (!/^[A-Za-z0-9]{4,20}$/.test(usr)) {
          agregar(SEV.ERROR, 'Usuario', 'El usuario debe tener de 4 a 20 caracteres, solo letras y números.',
            'Quitar espacios, guiones o acentos y ajustar la longitud.');
        }
        const previo = duplicado(api, 'usuario', usr.toUpperCase());
        if (previo) agregar(SEV.ERROR, 'Usuario', `Usuario duplicado: ya aparece en la fila ${previo}.`,
          'Dejar una sola fila por usuario o asignarle un usuario distinto.');
      }

      ['Nombre', 'Apellido Paterno', 'Apellido Materno'].forEach(c => {
        if (!u.vacio(d[c]) && /[0-9@#$%&*_=+<>]/.test(d[c])) {
          agregar(SEV.ERROR, c, 'Contiene números o símbolos.', 'Capturar solo letras.');
        }
      });

      if (!u.vacio(d['Correo'])) {
        const correo = d['Correo'].trim();
        if (!/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(correo)) {
          agregar(SEV.ERROR, 'Correo', 'El correo no tiene un formato válido.', 'Usar el formato nombre@dominio.com.');
        } else {
          const previo = duplicado(api, 'correo', correo.toLowerCase());
          if (previo) agregar(SEV.ADVERTENCIA, 'Correo', `El mismo correo ya se usó en la fila ${previo}.`,
            'El portal lo acepta, pero las contraseñas de ambos usuarios llegarán al mismo buzón. Confirmar que es intencional.');
        }
      }

      // Cascada geográfica: solo se valida el hijo si el padre es válido
      const regiones = Object.keys(GEOGRAFIA);
      if (catalogo(api, 'Región', d['Región'], regiones)) {
        const estados = Object.keys(GEOGRAFIA[d['Región']]);
        // Estado válido pero de otra región: error de cascada con pista concreta
        const otra = regiones.find(r => r !== d['Región'] && Object.keys(GEOGRAFIA[r]).includes(d['Estado']));
        if (otra) {
          agregar(SEV.ERROR, 'Estado', `${d['Estado']} no pertenece a ${d['Región']}; pertenece a ${otra}.`,
            `Cambiar Región a "${otra}" o corregir el Estado.`);
        } else if (catalogo(api, 'Estado', d['Estado'], estados, `para ${d['Región']}`)) {
          const ciudades = Object.keys(GEOGRAFIA[d['Región']][d['Estado']]);
          if (catalogo(api, 'Ciudad', d['Ciudad'], ciudades, `para ${d['Estado']}`)) {
            const municipios = GEOGRAFIA[d['Región']][d['Estado']][d['Ciudad']];
            catalogo(api, 'Municipio', d['Municipio'], municipios, `para ${d['Ciudad']}`);
          }
        }
      }

      catalogo(api, 'Tipo de Alumno', d['Tipo de Alumno'], TIPO_ALUMNO);
      catalogo(api, 'Tipo de Canal', d['Tipo de Canal'], TIPO_CANAL);
    },
    muestra: {
      archivo: 'muestra_alta_usuarios.xlsx',
      encabezados: ['Usuario', 'Nombre', 'Apellido Paterno', 'Apellido Materno', 'Correo', 'Región', 'Estado', 'Ciudad', 'Municipio', 'Tipo de Alumno', 'Tipo de Canal'],
      filas: [
        ['GDLDEMO01', 'Mariana', 'Ochoa', 'Ruelas', 'capacitacion@distribuidora-demo.mx', 'Región 5', 'Jalisco', 'Guadalajara', 'Zapopan', 'Distribuidor', 'Detalle de Canal'],
        ['GDLDEMO02', 'Luis Fernando', 'Ibarra', '', 'capacitacion@distribuidora-demo.mx', 'Región 5', 'Jalisco', 'Guadalajara', 'Tonalá', 'Distribuidor', 'Detalle de Canal'],
        ['GDLDEMO03', 'Karla', 'Méndez', 'Solís', 'kmendez@distribuidora-demo.mx', 'Region 5', 'Jalisco', 'Guadalajara', 'Tlaquepaque', 'Distribuidor', 'Detalle de canal'],
        ['MZODEMO01', 'Óscar', 'Valdivia', 'Ramos', 'ovaldivia@distribuidora-demo.mx', 'Región 5', 'Colima ', 'Manzanillo', 'Manzanillo', 'Distribuidor', 'Detalle de Canal'],
        ['TCMDEMO01', 'Brenda', 'Cárdenas', 'Lara', 'bcardenas@distribuidora-demo.mx', 'Región 9', 'Colima', 'Tecomán', 'Tecomán', 'Empleado', 'Masivo'],
        ['GDLDEMO01', 'Mariana', 'Ochoa', 'Ruelas', 'mochoa@distribuidora-demo.mx', 'Región 5', 'Jalisco', 'Guadalajara', 'Guadalajara', 'Distribuidor', 'Detalle de Canal'],
        ['TPC DEMO', 'Jorge', 'Partida', 'Núñez', 'jpartida@distribuidora-demo', 'Región 5', 'Nayarit', 'Tepic', 'Xalisco', 'Vendedor', 'Detalle de Canal'],
        ['MLMDEMO01', '', 'Zavala', 'Torres', 'czavala@distribuidora-demo.mx', 'Región 5', 'Michoacán', 'Morelia', 'Morelia', 'Externo', 'Cadenas Comerciales'],
        ['CDMDEMO01', 'Paola', 'Rentería', 'Gil', 'prenteria@distribuidora-demo.mx', 'Región 9', 'Ciudad de México', 'Ciudad de México', 'Coyoacán', 'Empleado', 'Cadenas Comerciales'],
        ['VALDEMO01', 'Ana Sofía', 'Lugo', 'Peña', 'alugo@distribuidora-demo.mx', 'Región 5', 'Colima', 'Colima', 'Villa de Alvarez', 'Distribuidor', 'Detalle de Canal']
      ]
    }
  };

  // ---------------------------------------------------------------------------
  // Perfil 2 — Consolidado de recargas (ERP)
  // ---------------------------------------------------------------------------
  const recargas = {
    id: 'recargas',
    nombre: 'Consolidado de recargas',
    destino: 'ERP de distribución',
    descripcion: '"Agente" vs. "Master", números de 10 dígitos, montos válidos y recargas duplicadas.',
    origen: 'El ERP descartaba en silencio las recargas con bolsa "Master": 60 recargas nunca entraron y no hubo ningún mensaje.',
    columnas: [
      { nombre: 'Folio', obligatorio: true },
      { nombre: 'Fecha', obligatorio: true },
      { nombre: 'Número de Agente', obligatorio: true },
      { nombre: 'Nombre de Agente', obligatorio: true },
      { nombre: 'Número Destino', obligatorio: true },
      { nombre: 'Monto', obligatorio: true },
      { nombre: 'Movimientos en la Bolsa de', obligatorio: true, catalogo: true },
      { nombre: 'Estatus', obligatorio: true, catalogo: true }
    ],
    validarFila(d, api) {
      const { SEV, agregar, ctx } = api;
      const u = ctx.utiles;

      if (!u.vacio(d['Folio'])) {
        const previo = duplicado(api, 'folio', d['Folio'].trim());
        if (previo) agregar(SEV.ERROR, 'Folio', `Folio duplicado: ya aparece en la fila ${previo}.`,
          'Eliminar la fila repetida; si son recargas distintas, verificar el folio en el reporte de origen.');
      }

      const f = fecha(api, 'Fecha', d['Fecha'], { noFutura: true });
      telefono(api, 'Número de Agente', d['Número de Agente']);
      const destinoOk = telefono(api, 'Número Destino', d['Número Destino']);

      let monto = null;
      if (!u.vacio(d['Monto'])) {
        const m = u.leerMonto(d['Monto']);
        if (!m) {
          agregar(SEV.ERROR, 'Monto', `"${d['Monto']}" no es un monto válido (texto, o separador de miles ambiguo).`, 'Capturar solo el número, por ejemplo 100.');
        } else if (!DENOMINACIONES.includes(m.numero)) {
          agregar(SEV.ERROR, 'Monto', `$${m.numero} no es una denominación de recarga válida.`,
            `Usar una de: ${DENOMINACIONES.map(x => '$' + x).join(', ')}.`);
        } else {
          monto = m.numero;
          if (m.conFormato) agregar(SEV.ADVERTENCIA, 'Monto', 'El monto trae símbolo de pesos, comas o espacios.',
            `Dejar solo el número: ${m.numero}.`);
        }
      }

      // La regla que originó el proyecto
      const bolsa = d['Movimientos en la Bolsa de'];
      if (!u.vacio(bolsa)) {
        const r = u.revisarCatalogo(bolsa, BOLSA);
        const valor = r.ok ? bolsa : r.sugerencia;
        if (valor === 'Master') {
          agregar(SEV.ERROR, 'Movimientos en la Bolsa de',
            'El ERP solo acepta "Agente". Las filas con "Master" se descartan en silencio: la recarga no entra y no aparece ningún error.',
            'Cambiar a "Agente" y poner en Número/Nombre de Agente la sucursal que realmente vendió la recarga.');
        } else if (!r.ok) {
          catalogo(api, 'Movimientos en la Bolsa de', bolsa, ['Agente']);
        }
      }

      if (!u.vacio(d['Estatus'])) {
        if (catalogo(api, 'Estatus', d['Estatus'], ESTATUS_RECARGA) && d['Estatus'] !== 'Exitosa') {
          agregar(SEV.ERROR, 'Estatus', `Una recarga "${d['Estatus']}" no se carga al consolidado.`,
            'Quitar la fila del archivo; solo se cargan recargas exitosas.');
        }
      }

      // Posible duplicado con folio distinto: mismo destino, monto y día
      if (destinoOk && monto !== null && f) {
        const previo = duplicado(api, 'recarga', `${d['Número Destino'].trim()}|${monto}|${f.iso}`);
        if (previo) agregar(SEV.ADVERTENCIA, 'Número Destino',
          `Misma recarga que la fila ${previo} (mismo número, monto y fecha) con otro folio.`,
          'Confirmar en el portal de origen que fueron dos recargas distintas antes de cargarlas.');
      }
    },
    muestra: {
      archivo: 'muestra_consolidado_recargas.xlsx',
      encabezados: ['Folio', 'Fecha', 'Número de Agente', 'Nombre de Agente', 'Número Destino', 'Monto', 'Movimientos en la Bolsa de', 'Estatus'],
      filas: [
        ['RC-100231', '14/10/2026', '3310000001', 'SUC DEMO CENTRO', '3319876501', '100', 'Agente', 'Exitosa'],
        ['RC-100232', '14/10/2026', '3310000001', 'SUC DEMO CENTRO', '3319876502', '50', 'Agente', 'Exitosa'],
        ['RC-100233', '14/10/2026', '3310000099', 'ADMIN DEMO', '3319876503', '200', 'Master', 'Exitosa'],
        ['RC-100234', '15/10/2026', '3310000099', 'ADMIN DEMO', '3319876504', '100', 'Master', 'Exitosa'],
        ['RC-100235', '15/10/2026', '3310000002', 'SUC DEMO PLAZA', '331987650', '30', 'Agente', 'Exitosa'],
        ['RC-100236', '15/10/2026', '3310000002', 'SUC DEMO PLAZA', '523319876506', '100', 'Agente', 'Exitosa'],
        ['RC-100237', '16/10/2026', '3310000002', 'SUC DEMO PLAZA', '3319876507', '75', 'Agente', 'Exitosa'],
        ['RC-100238', '16/10/2026', '3310000003', 'SUC DEMO NORTE', '3319876508', '$1,00', 'Agente', 'Exitosa'],
        ['RC-100232', '16/10/2026', '3310000001', 'SUC DEMO CENTRO', '3319876509', '20', 'Agente', 'Exitosa'],
        ['RC-100240', '31/09/2026', '3310000003', 'SUC DEMO NORTE', '3319876510', '100', 'Agente', 'Exitosa'],
        ['RC-100241', '17/10/2026', '3310000003', 'SUC DEMO NORTE', '3319876511', '$150', 'Agente', 'Exitosa'],
        ['RC-100242', '17/10/2026', '3310000001', 'SUC DEMO CENTRO', '3319876501', '100', 'Agente', 'Exitosa'],
        ['RC-100243', '14/10/2026', '3310000001', 'SUC DEMO CENTRO', '3319876501', '100', 'Agente', 'Exitosa'],
        ['RC-100244', '17/10/2026', '3310000003', 'SUC DEMO NORTE', '3319876513', '50', 'Agente', 'Fallida']
      ]
    }
  };

  // ---------------------------------------------------------------------------
  // Perfil 3 — Reporte regulatorio de portabilidad
  // ---------------------------------------------------------------------------

  /** Detecta números capturados como rango/bloque. Devuelve {desde, hasta, cantidad} o null. */
  function detectarRango(valor) {
    const s = String(valor).trim();
    let r = s.match(/^(\d{10})\s*(?:-|–|a|al|hasta)\s*(\d{10})$/i);
    if (r) {
      const a = Number(r[1]), b = Number(r[2]);
      return { desde: r[1], hasta: r[2], cantidad: Math.abs(b - a) + 1 };
    }
    r = s.match(/^(\d{10})\s*(?:-|–|a|al|hasta)\s*(\d{1,4})$/i);
    if (r) { // 3312345600-99 → hasta 3312345699
      const hasta = r[1].slice(0, 10 - r[2].length) + r[2];
      return { desde: r[1], hasta, cantidad: Math.abs(Number(hasta) - Number(r[1])) + 1 };
    }
    r = s.match(/^(\d{6,9})([Xx*]{1,4})$/);
    if (r && (r[1].length + r[2].length) === 10) {
      const n = r[2].length;
      return { desde: r[1] + '0'.repeat(n), hasta: r[1] + '9'.repeat(n), cantidad: Math.pow(10, n) };
    }
    return null;
  }

  const portabilidad = {
    id: 'portabilidad',
    nombre: 'Reporte de portabilidad',
    destino: 'Formato regulatorio',
    descripcion: 'Una línea por número (no rangos por bloque), fechas válidas y tipo de movimiento.',
    origen: 'El formato exigía una línea por número individual, y el archivo traía rangos por bloque.',
    columnas: [
      { nombre: 'Número', obligatorio: true },
      { nombre: 'Tipo de Movimiento', obligatorio: true, catalogo: true },
      { nombre: 'Fecha de Solicitud', obligatorio: true },
      { nombre: 'Fecha de Ejecución', obligatorio: true },
      { nombre: 'Proveedor Donador', obligatorio: true, catalogo: true },
      { nombre: 'Proveedor Receptor', obligatorio: true, catalogo: true }
    ],
    validarFila(d, api) {
      const { SEV, agregar, ctx } = api;
      const u = ctx.utiles;

      if (!u.vacio(d['Número'])) {
        const rango = detectarRango(d['Número']);
        if (rango) {
          agregar(SEV.ERROR, 'Número',
            `Es un rango por bloque (${rango.cantidad.toLocaleString('es-MX')} números). El formato exige una línea por número individual.`,
            `Separar en ${rango.cantidad.toLocaleString('es-MX')} líneas, del ${rango.desde} al ${rango.hasta}, repitiendo los demás datos.`);
        } else if (telefono(api, 'Número', d['Número'])) {
          const previo = duplicado(api, 'numero', d['Número'].trim() + '|' + d['Tipo de Movimiento']);
          if (previo) agregar(SEV.ERROR, 'Número', `El número ya aparece en la fila ${previo} con el mismo tipo de movimiento.`,
            'Dejar una sola línea por número y movimiento.');
        }
      }

      catalogo(api, 'Tipo de Movimiento', d['Tipo de Movimiento'], TIPO_MOVIMIENTO);

      const fs = fecha(api, 'Fecha de Solicitud', d['Fecha de Solicitud'], { noFutura: true });
      const fe = fecha(api, 'Fecha de Ejecución', d['Fecha de Ejecución']);
      if (fs && fe && fe.fecha < fs.fecha) {
        agregar(SEV.ERROR, 'Fecha de Ejecución',
          `La ejecución (${fe.texto}) es anterior a la solicitud (${fs.texto}).`,
          'Revisar ambas fechas; la ejecución debe ser el mismo día o posterior a la solicitud.');
      }

      const don = catalogo(api, 'Proveedor Donador', d['Proveedor Donador'], PROVEEDORES);
      const rec = catalogo(api, 'Proveedor Receptor', d['Proveedor Receptor'], PROVEEDORES);
      if (don && rec && d['Proveedor Donador'] === d['Proveedor Receptor']) {
        agregar(SEV.ERROR, 'Proveedor Receptor', 'El proveedor donador y el receptor son el mismo.',
          'Verificar el operador de origen y el de destino de la portabilidad.');
      }
    },
    muestra: {
      archivo: 'muestra_reporte_portabilidad.xlsx',
      encabezados: ['Número', 'Tipo de Movimiento', 'Fecha de Solicitud', 'Fecha de Ejecución', 'Proveedor Donador', 'Proveedor Receptor'],
      filas: [
        ['3320001001', 'Portabilidad Entrante', '01/10/2026', '02/10/2026', 'Telcel', 'Wimo'],
        ['3320001002', 'Portabilidad Entrante', '01/10/2026', '02/10/2026', 'AT&T', 'Wimo'],
        ['3320001100-3320001199', 'Portabilidad Entrante', '03/10/2026', '04/10/2026', 'Movistar', 'Wimo'],
        ['33200012XX', 'Portabilidad Saliente', '05/10/2026', '06/10/2026', 'Wimo', 'Telcel'],
        ['3320001003', 'Portabilidad entrante', '05/10/2026', '06/10/2026', 'Bait', 'Wimo'],
        ['3320001004', 'Portabilidad Saliente', '08/10/2026', '07/10/2026', 'Wimo', 'AT&T'],
        ['3320001002', 'Portabilidad Entrante', '09/10/2026', '09/10/2026', 'AT&T', 'Wimo'],
        ['3320001001', 'Reversión', '09/10/2026', '09/10/2026', 'Wimo', 'Telcel'],
        ['3320001005', 'Portabilidad Entrante', '29/02/2026', '02/03/2026', 'Telcel', 'Wimo'],
        ['3320001006', 'Portabilidad Entrante', '10/10/2026', '11/10/2026', 'Wimo', 'Wimo'],
        ['3320001007', 'Portabilidad Saliente', '12/10/2026', '13/10/2026', 'Wimo', 'Virgin Mobile']
      ]
    }
  };

  const lista = [usuarios, recargas, portabilidad];
  return {
    lista,
    porId: Object.fromEntries(lista.map(p => [p.id, p])),
    catalogos: { GEOGRAFIA, TIPO_ALUMNO, TIPO_CANAL, BOLSA, DENOMINACIONES, ESTATUS_RECARGA, TIPO_MOVIMIENTO, PROVEEDORES },
    detectarRango
  };
});
