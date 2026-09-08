/**
 * STOCK DE CAUSAS Y RANKING PARA TABLA
 * Backend Google Apps Script — lee y escribe DIRECTAMENTE sobre las hojas reales
 * ya existentes en la planilla (importadas desde el Excel de control):
 *
 *   - Base_Datos : encabezados en la fila 6, datos desde la fila 7 (una fila por causa)
 *   - Stock      : encabezados en la fila 1 (listado de referencia para la MATERIA)
 *   - Config     : prioridad de LIBRO para el ranking (encabezados NIVEL/LIBRO/PUNTAJE)
 *
 * No se reescribe la hoja completa en cada acción: se hacen escrituras puntuales
 * (celda o rango específico) para no arriesgar datos ni formato ya existentes.
 * Las columnas se ubican por NOMBRE de encabezado (no por letra fija), igual que
 * hacía el macro VBA original (BuscarColumna), para tolerar pequeños cambios de orden.
 *
 * IMPORTANTE: cambiar CLAVE antes de desplegar. No es seguridad real (el
 * frontend es público en GitHub Pages), es solo una traba básica.
 */

var CLAVE = 'Causas5689';

var HOJA_BASE = 'Base_Datos';
var HOJA_STOCK = 'Stock';
var HOJA_CONFIG = 'Config';

var FILA_ENCABEZADO_BASE = 6;
var FILA_ENCABEZADO_STOCK = 1;

var PUNTAJES_LIBRO_DEFECTO = [300, 200, 100, 50, 25];

function doGet(e) {
  return responder(procesar((e && e.parameter) || {}));
}

function doPost(e) {
  var datos = {};
  try {
    datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (error) {
    return responder({ ok: false, error: 'Solicitud mal formada.' });
  }
  return responder(procesar(datos));
}

function responder(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function procesar(datos) {
  try {
    if (String(datos.clave || '') !== CLAVE) {
      return { ok: false, error: 'Clave incorrecta o no enviada.' };
    }
    switch (datos.accion) {
      case 'probar':
        return { ok: true, mensaje: 'Servicio operativo', version: 'v4-diagnostico-2026-09-08' };
      case 'diagnostico':
        return diagnostico(datos);
      case 'listar':
        return listarTodo();
      case 'ingestarNuevas':
        return ingestarNuevas(datos.filas || []);
      case 'cargarStock':
        return cargarStock(datos.filas || []);
      case 'actualizarMaterias':
        return actualizarMaterias();
      case 'generarRanking':
        return generarRanking();
      case 'guardarConfig':
        return guardarConfig(datos.prioridad || []);
      case 'confirmarTabla':
        return confirmarTabla(datos.idsFila || []);
      case 'actualizarCausa':
        return actualizarCausa(datos.idFila, datos.cambios || {});
      case 'eliminarFila':
        return eliminarFila(datos.idFila);
      case 'detectarDuplicados':
        return { ok: true, grupos: detectarDuplicados(leerCausas()) };
      case 'reportes':
        return reportes();
      default:
        return { ok: false, error: 'Acción no reconocida: ' + datos.accion };
    }
  } catch (error) {
    return { ok: false, error: String(error && error.message ? error.message : error) };
  }
}

/* ------------------------- normalización y columnas ------------------------- */

function normalizar(texto) {
  return String(texto == null ? '' : texto)
    .toLowerCase()
    .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

/** Busca el índice (0-based) de la primera columna cuyo encabezado normalizado
 *  contiene el texto buscado (también normalizado). */
/** Busca por coincidencia EXACTA primero (evita que "MATERIA" confunda con
 *  "COD MATERIA", que también la contiene como substring); si no hay
 *  coincidencia exacta, recién ahí cae a substring. */
function buscarColumna(headers, buscado) {
  var obj = normalizar(buscado);
  for (var i = 0; i < headers.length; i++) {
    if (normalizar(headers[i]) === obj) return i;
  }
  for (var j = 0; j < headers.length; j++) {
    if (normalizar(headers[j]).indexOf(obj) > -1) return j;
  }
  return -1;
}

function mapaColumnasBase(headers) {
  return {
    rol: buscarColumna(headers, 'rol'),
    anio: buscarColumna(headers, 'ano'),
    libro: buscarColumna(headers, 'libro'),
    fechaIngreso: buscarColumna(headers, 'fecha ingreso'),
    observacion: buscarColumna(headers, 'observacion'),
    fechaRelacion: buscarColumna(headers, 'fecha relacion'),
    materia: buscarColumna(headers, 'materia'),
    complejidad: buscarColumna(headers, 'complejidad'),
    tipo: buscarColumna(headers, 'tipo'),
    estado: buscarColumna(headers, 'estado de la causa'),
    rija: buscarColumna(headers, 'rija'),
    fechaSuspension: buscarColumna(headers, 'fecha de suspension'),
    diasSuspension: buscarColumna(headers, 'cantidad de dias'),
    comentarios: buscarColumna(headers, 'comentarios')
  };
}

function mapaColumnasStock(headers) {
  var libro = buscarColumna(headers, 'tipo libro');
  if (libro === -1) libro = buscarColumna(headers, 'libro');
  return {
    rol: buscarColumna(headers, 'rol'),
    anio: buscarColumna(headers, 'ano'),
    libro: libro,
    materia: buscarColumna(headers, 'materia')
  };
}

function hojaBase() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_BASE);
  if (!h) throw new Error('No se encontró la hoja "' + HOJA_BASE + '".');
  return h;
}
function hojaStock() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_STOCK);
  if (!h) throw new Error('No se encontró la hoja "' + HOJA_STOCK + '".');
  return h;
}
function hojaConfig() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  if (!h) throw new Error('No se encontró la hoja "' + HOJA_CONFIG + '".');
  return h;
}

function formatearValor(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT-3', 'dd-MM-yyyy');
  }
  return v;
}

/* ------------------------------ lectura ------------------------------ */

function leerCausas() {
  var h = hojaBase();
  var ultimaFila = h.getLastRow();
  var ultimaCol = h.getLastColumn();
  if (ultimaFila < FILA_ENCABEZADO_BASE + 1) return [];

  var headers = h.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaCol).getValues()[0];
  var col = mapaColumnasBase(headers);
  if (col.rol === -1 || col.libro === -1 || col.estado === -1) {
    throw new Error('No se reconocen las columnas clave de "Base_Datos" (ROL/LIBRO/Estado). Revisa la fila ' + FILA_ENCABEZADO_BASE + '.');
  }

  var filas = h.getRange(FILA_ENCABEZADO_BASE + 1, 1, ultimaFila - FILA_ENCABEZADO_BASE, ultimaCol).getValues();
  var out = [];
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    if (!f[col.rol] && !f[col.libro]) continue; // fila vacía
    out.push({
      idFila: FILA_ENCABEZADO_BASE + 1 + i,
      rol: formatearValor(f[col.rol]),
      anio: formatearValor(f[col.anio]),
      libro: formatearValor(f[col.libro]),
      fechaIngreso: formatearValor(f[col.fechaIngreso]),
      observacion: formatearValor(f[col.observacion]),
      fechaRelacion: formatearValor(f[col.fechaRelacion]),
      materia: formatearValor(f[col.materia]),
      complejidad: col.complejidad > -1 ? formatearValor(f[col.complejidad]) : '',
      tipo: col.tipo > -1 ? formatearValor(f[col.tipo]) : '',
      estado: formatearValor(f[col.estado]),
      rija: col.rija > -1 ? formatearValor(f[col.rija]) : '',
      fechaSuspension: col.fechaSuspension > -1 ? formatearValor(f[col.fechaSuspension]) : '',
      diasSuspension: col.diasSuspension > -1 ? formatearValor(f[col.diasSuspension]) : '',
      comentarios: col.comentarios > -1 ? formatearValor(f[col.comentarios]) : ''
    });
  }
  return out;
}

function leerStock() {
  var h = hojaStock();
  var ultimaFila = h.getLastRow();
  var ultimaCol = h.getLastColumn();
  if (ultimaFila < FILA_ENCABEZADO_STOCK + 1) return [];
  var headers = h.getRange(FILA_ENCABEZADO_STOCK, 1, 1, ultimaCol).getValues()[0];
  var col = mapaColumnasStock(headers);
  var filas = h.getRange(FILA_ENCABEZADO_STOCK + 1, 1, ultimaFila - FILA_ENCABEZADO_STOCK, ultimaCol).getValues();
  var out = [];
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    if (col.rol === -1 || !f[col.rol]) continue;
    out.push({
      fila: FILA_ENCABEZADO_STOCK + 1 + i,
      rol: formatearValor(f[col.rol]),
      anio: col.anio > -1 ? formatearValor(f[col.anio]) : '',
      libro: col.libro > -1 ? formatearValor(f[col.libro]) : '',
      materia: col.materia > -1 ? formatearValor(f[col.materia]) : ''
    });
  }
  return out;
}

/** Diagnóstico temporal: muestra los encabezados reales y 2 filas de ejemplo
 *  (TODAS las columnas, no solo las mapeadas) de Base_Datos y Stock, para
 *  detectar por qué MATERIA muestra códigos en vez de texto descriptivo.
 *  Si se envía rol/anio/libro, además busca esa fila EXACTA en Stock (con
 *  todas sus columnas) para comparar "COD MATERIA" contra "MATERIA".
 *  Quitar esta acción una vez resuelto el problema. */
function diagnostico(datos) {
  var hb = hojaBase();
  var ultimaColB = hb.getLastColumn();
  var headersBase = hb.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaColB).getValues()[0];
  var muestraBase = hb.getLastRow() >= FILA_ENCABEZADO_BASE + 2
    ? hb.getRange(FILA_ENCABEZADO_BASE + 1, 1, 2, ultimaColB).getValues()
    : [];

  var hs = hojaStock();
  var ultimaColS = hs.getLastColumn();
  var headersStock = hs.getRange(FILA_ENCABEZADO_STOCK, 1, 1, ultimaColS).getValues()[0];
  var colStock = mapaColumnasStock(headersStock);
  var muestraStock = hs.getLastRow() >= FILA_ENCABEZADO_STOCK + 2
    ? hs.getRange(FILA_ENCABEZADO_STOCK + 1, 1, 2, ultimaColS).getValues()
    : [];

  var busqueda = null;
  if (datos && datos.rol) {
    var claveBuscada = claveCruce(datos.rol, datos.anio, datos.libro);
    busqueda = { clave: claveBuscada, encontradaEnStock: false, encontradaEnBase: false };
    var ultimaFilaS = hs.getLastRow();
    if (ultimaFilaS >= FILA_ENCABEZADO_STOCK + 1) {
      var todasStock = hs.getRange(FILA_ENCABEZADO_STOCK + 1, 1, ultimaFilaS - FILA_ENCABEZADO_STOCK, ultimaColS).getValues();
      for (var i = 0; i < todasStock.length; i++) {
        var f = todasStock[i];
        if (claveCruce(f[colStock.rol], f[colStock.anio], f[colStock.libro]) === claveBuscada) {
          busqueda.encontradaEnStock = true;
          busqueda.filaStock = FILA_ENCABEZADO_STOCK + 1 + i;
          busqueda.valoresStock = f.map(formatearValor);
          busqueda.materiaSegunColMapeada = colStock.materia > -1 ? formatearValor(f[colStock.materia]) : null;
          break;
        }
      }
    }
    var causaBase = leerCausas().filter(function(c) { return claveCruce(c.rol, c.anio, c.libro) === claveBuscada; })[0];
    if (causaBase) {
      busqueda.encontradaEnBase = true;
      busqueda.materiaEnBaseDatos = causaBase.materia;
    }
  }

  return {
    ok: true,
    baseDatos: { headers: headersBase, colMapeada: mapaColumnasBase(headersBase), muestra: muestraBase.map(function(f){ return f.map(formatearValor); }) },
    stock: { headers: headersStock, colMapeada: colStock, muestra: muestraStock.map(function(f){ return f.map(formatearValor); }) },
    busqueda: busqueda
  };
}

function claveCruce(rol, anio, libro) {
  return String(rol || '') + '|' + String(anio || '') + '|' + String(libro || '').trim().toUpperCase();
}

function indiceMateriaStock() {
  var indice = {};
  leerStock().forEach(function (s) {
    indice[claveCruce(s.rol, s.anio, s.libro)] = s.materia;
  });
  return indice;
}

function leerConfig() {
  var h = hojaConfig();
  var ultimaFila = h.getLastRow();
  var datos = h.getRange(1, 1, ultimaFila, Math.max(h.getLastColumn(), 3)).getValues();
  var filaEncabezado = -1;
  for (var i = 0; i < datos.length; i++) {
    if (normalizar(datos[i][0]) === 'nivel') { filaEncabezado = i; break; }
  }
  var niveles = [];
  if (filaEncabezado > -1) {
    for (var r = filaEncabezado + 1; r < datos.length; r++) {
      var libro = datos[r][1];
      var puntaje = datos[r][2];
      if (!libro || typeof puntaje !== 'number') break; // fin de la tabla de niveles
      niveles.push({ nivel: niveles.length + 1, libro: String(libro).trim(), puntaje: puntaje });
    }
  }
  if (niveles.length === 0) {
    var libros = ['Laboral - Cobranza', 'Familia', 'Civil', 'Policia Local', '(Ninguno)'];
    niveles = libros.map(function (l, i) { return { nivel: i + 1, libro: l, puntaje: PUNTAJES_LIBRO_DEFECTO[i] }; });
  }
  return { niveles: niveles, filaEncabezado: filaEncabezado };
}

/* ------------------------------- acciones ------------------------------- */

function listarTodo() {
  var cfg = leerConfig();
  return { ok: true, causas: leerCausas(), config: cfg.niveles, stockCount: leerStock().length };
}

function ingestarNuevas(filasNuevas) {
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    var h = hojaBase();
    var ultimaFila = h.getLastRow();
    var ultimaCol = h.getLastColumn();
    var headers = h.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaCol).getValues()[0];
    var col = mapaColumnasBase(headers);

    var existentes = leerCausas();
    var clavesExistentes = {};
    existentes.forEach(function (c) { clavesExistentes[claveCruce(c.rol, c.anio, c.libro)] = true; });

    var indiceStock = indiceMateriaStock();

    var agregadas = 0, omitidas = 0;
    var filasParaEscribir = [];

    filasNuevas.forEach(function (f) {
      var clave = claveCruce(f.rol, f.anio, f.libro);
      if (clavesExistentes[clave]) { omitidas++; return; }
      clavesExistentes[clave] = true;
      agregadas++;
      var fila = new Array(ultimaCol).fill('');
      if (col.rol > -1) fila[col.rol] = f.rol || '';
      if (col.anio > -1) fila[col.anio] = f.anio || '';
      if (col.libro > -1) fila[col.libro] = f.libro || '';
      if (col.fechaIngreso > -1) fila[col.fechaIngreso] = f.fechaIngreso || '';
      if (col.observacion > -1) fila[col.observacion] = f.observacion || '';
      if (col.fechaRelacion > -1) fila[col.fechaRelacion] = f.fechaRelacion || '';
      if (col.materia > -1) fila[col.materia] = indiceStock[clave] || '';
      if (col.estado > -1) fila[col.estado] = 'Disponible';
      filasParaEscribir.push(fila);
    });

    if (filasParaEscribir.length > 0) {
      h.getRange(ultimaFila + 1, 1, filasParaEscribir.length, ultimaCol).setValues(filasParaEscribir);
    }

    return { ok: true, agregadas: agregadas, omitidas: omitidas };
  } finally {
    candado.releaseLock();
  }
}

/** Actualiza (upsert) el listado Stock: si ROL+AÑO+LIBRO ya existe, solo
 *  corrige su MATERIA; si no existe, agrega una fila nueva mínima. Nunca
 *  borra ni reemplaza filas existentes (es una planilla de referencia viva). */
function cargarStock(filasNuevas) {
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    var h = hojaStock();
    var ultimaFila = h.getLastRow();
    var ultimaCol = Math.max(h.getLastColumn(), 4);
    var headers = h.getRange(FILA_ENCABEZADO_STOCK, 1, 1, ultimaCol).getValues()[0];
    var col = mapaColumnasStock(headers);

    if (col.rol === -1 || col.libro === -1 || col.materia === -1) {
      throw new Error('No se reconocen las columnas ROL/LIBRO/MATERIA en "Stock".');
    }

    var indiceFilas = {};
    if (ultimaFila >= FILA_ENCABEZADO_STOCK + 1) {
      var actuales = h.getRange(FILA_ENCABEZADO_STOCK + 1, 1, ultimaFila - FILA_ENCABEZADO_STOCK, ultimaCol).getValues();
      actuales.forEach(function (f, i) {
        if (!f[col.rol]) return;
        indiceFilas[claveCruce(f[col.rol], f[col.anio], f[col.libro])] = FILA_ENCABEZADO_STOCK + 1 + i;
      });
    }

    var actualizadas = 0, agregadas = 0;
    var nuevasFilas = [];

    filasNuevas.forEach(function (f) {
      var clave = claveCruce(f.rol, f.anio, f.libro);
      var filaExistente = indiceFilas[clave];
      if (filaExistente) {
        h.getRange(filaExistente, col.materia + 1).setValue(f.materia || '');
        actualizadas++;
      } else {
        var fila = new Array(ultimaCol).fill('');
        fila[col.rol] = f.rol || '';
        if (col.anio > -1) fila[col.anio] = f.anio || '';
        fila[col.libro] = f.libro || '';
        fila[col.materia] = f.materia || '';
        nuevasFilas.push(fila);
        agregadas++;
      }
    });

    if (nuevasFilas.length > 0) {
      h.getRange(h.getLastRow() + 1, 1, nuevasFilas.length, ultimaCol).setValues(nuevasFilas);
    }

    return { ok: true, actualizadas: actualizadas, agregadas: agregadas };
  } finally {
    candado.releaseLock();
  }
}

function actualizarMaterias() {
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    var h = hojaBase();
    var ultimaFila = h.getLastRow();
    var ultimaCol = h.getLastColumn();
    var headers = h.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaCol).getValues()[0];
    var col = mapaColumnasBase(headers);
    if (col.materia === -1) throw new Error('No se encontró la columna MATERIA en "Base_Datos".');

    var causas = leerCausas();
    var indiceStock = indiceMateriaStock();

    // Una sola lectura y una sola escritura de toda la columna MATERIA
    // (en vez de una llamada por fila, que con miles de filas se corta por tiempo).
    var filaInicio = FILA_ENCABEZADO_BASE + 1;
    var numFilas = ultimaFila - FILA_ENCABEZADO_BASE;
    var rango = h.getRange(filaInicio, col.materia + 1, numFilas, 1);
    var valores = rango.getValues();
    var actualizadas = 0;

    causas.forEach(function (c) {
      var materiaStock = indiceStock[claveCruce(c.rol, c.anio, c.libro)];
      if (materiaStock && materiaStock !== c.materia) {
        valores[c.idFila - filaInicio][0] = materiaStock;
        actualizadas++;
      }
    });

    if (actualizadas > 0) rango.setValues(valores);

    return { ok: true, actualizadas: actualizadas };
  } finally {
    candado.releaseLock();
  }
}

/** Días para reanudar = (fecha de suspensión + días de suspensión) - hoy.
 *  Sin fecha de suspensión => null (el criterio de vencimiento no aplica). */
function diasParaReanudar(causa) {
  if (!causa.fechaSuspension || !causa.diasSuspension) return null;
  var fSusp = parsearFecha(causa.fechaSuspension);
  if (!fSusp) return null;
  var fReanuda = new Date(fSusp.getTime());
  fReanuda.setDate(fReanuda.getDate() + Number(causa.diasSuspension));
  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  var msPorDia = 24 * 60 * 60 * 1000;
  return Math.round((fReanuda.getTime() - hoy.getTime()) / msPorDia);
}

/** Acepta Date, ISO, o dd-MM-yyyy / dd/MM/yyyy. */
function parsearFecha(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') return valor;
  var texto = String(valor || '').trim();
  if (!texto) return null;
  var m = texto.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(texto);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Algoritmo de ranking — replica EXACTO el macro VBA "ActualizarRanking"
 * del Excel de control validado por el usuario:
 *  1) RIJA (texto contiene "rija")             -> +1000
 *  2) MATERIA contiene "pesca" y "acuicultura" -> +800
 *  3) Días para reanudar <= 0 (si aplica)      -> +600
 *  4) LIBRO coincide con nivel de Config       -> +puntaje del nivel (solo el primero que matchea)
 *  5) Desempate: FECHA RELACIÓN más antigua primero
 * Solo participan causas con estado "Disponible".
 */
function generarRanking() {
  var causas = leerCausas().filter(function (c) { return c.estado === 'Disponible'; });
  var prioridad = leerConfig().niveles;

  var conteo = { rija: 0, pesca: 0, colP: 0 };

  var calculadas = causas.map(function (c) {
    var puntaje = 0;
    var motivos = [];

    if (String(c.rija || '').toUpperCase().indexOf('RIJA') > -1) {
      puntaje += 1000;
      motivos.push('Contiene RIJA');
      conteo.rija++;
    }

    var materiaMin = String(c.materia || '').toLowerCase();
    if (materiaMin.indexOf('pesca') > -1 && materiaMin.indexOf('acuicultura') > -1) {
      puntaje += 800;
      motivos.push('Materia: Pesca y Acuicultura');
      conteo.pesca++;
    }

    var dias = diasParaReanudar(c);
    if (dias !== null && dias <= 0) {
      puntaje += 600;
      motivos.push('Suspensión vencida (días para reanudar ≤ 0)');
      conteo.colP++;
    }

    for (var i = 0; i < prioridad.length; i++) {
      var nivel = prioridad[i];
      if (nivel.libro && nivel.libro !== '(Ninguno)' && String(c.libro).trim() === String(nivel.libro).trim()) {
        puntaje += Number(nivel.puntaje) || 0;
        motivos.push('Libro prioridad ' + nivel.nivel + ' (' + nivel.libro + ')');
        break;
      }
    }

    var fechaOrden = 99999999999;
    var fr = parsearFecha(c.fechaRelacion);
    if (fr) fechaOrden = fr.getTime();

    return {
      idFila: c.idFila, rol: c.rol, anio: c.anio, libro: c.libro, materia: c.materia,
      fechaRelacion: c.fechaRelacion, complejidad: c.complejidad, tipo: c.tipo,
      puntaje: puntaje, motivos: motivos, fechaOrden: fechaOrden
    };
  });

  calculadas.sort(function (a, b) {
    if (b.puntaje !== a.puntaje) return b.puntaje - a.puntaje;
    return a.fechaOrden - b.fechaOrden;
  });

  var top = calculadas.slice(0, 40);

  return {
    ok: true,
    ranking: top,
    disponibles: causas.length,
    enRanking: top.length,
    conRIJA: conteo.rija,
    conPesca: conteo.pesca,
    conSuspensionVencida: conteo.colP,
    actualizado: new Date().toISOString()
  };
}

function guardarConfig(prioridad) {
  var cfg = leerConfig();
  if (cfg.filaEncabezado === -1) throw new Error('No se encontró la fila de encabezado (NIVEL/LIBRO/PUNTAJE) en "Config".');
  var h = hojaConfig();
  var filaInicio = cfg.filaEncabezado + 2; // 1-based, fila siguiente al encabezado
  var filas = prioridad.map(function (p, i) { return [i + 1, p.libro || '(Ninguno)', p.puntaje || PUNTAJES_LIBRO_DEFECTO[i] || 0]; });
  if (filas.length) h.getRange(filaInicio, 1, filas.length, 3).setValues(filas);
  return { ok: true };
}

function confirmarTabla(idsFila) {
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    var h = hojaBase();
    var ultimaCol = h.getLastColumn();
    var headers = h.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaCol).getValues()[0];
    var col = mapaColumnasBase(headers);
    var actualizadas = 0;
    idsFila.forEach(function (idFila) {
      var fila = Number(idFila);
      if (!fila || fila <= FILA_ENCABEZADO_BASE) return;
      h.getRange(fila, col.estado + 1).setValue('En Tabla');
      actualizadas++;
    });
    return { ok: true, actualizadas: actualizadas };
  } finally {
    candado.releaseLock();
  }
}

var CAMPO_A_CLAVE_BUSQUEDA = {
  rol: 'rol', anio: 'ano', libro: 'libro', fechaIngreso: 'fecha ingreso',
  observacion: 'observacion', fechaRelacion: 'fecha relacion', materia: 'materia',
  complejidad: 'complejidad', tipo: 'tipo', estado: 'estado de la causa', rija: 'rija',
  fechaSuspension: 'fecha de suspension', diasSuspension: 'cantidad de dias', comentarios: 'comentarios'
};

function actualizarCausa(idFila, cambios) {
  var fila = Number(idFila);
  if (!fila || fila <= FILA_ENCABEZADO_BASE) return { ok: false, error: 'idFila inválido.' };
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    var h = hojaBase();
    var ultimaCol = h.getLastColumn();
    var headers = h.getRange(FILA_ENCABEZADO_BASE, 1, 1, ultimaCol).getValues()[0];
    var col = mapaColumnasBase(headers);

    Object.keys(cambios).forEach(function (campo) {
      var claveBusqueda = CAMPO_A_CLAVE_BUSQUEDA[campo];
      if (!claveBusqueda) return;
      var indice = col[campo];
      if (indice > -1) h.getRange(fila, indice + 1).setValue(cambios[campo]);
    });

    return { ok: true };
  } finally {
    candado.releaseLock();
  }
}

function eliminarFila(idFila) {
  var fila = Number(idFila);
  if (!fila || fila <= FILA_ENCABEZADO_BASE) return { ok: false, error: 'idFila inválido.' };
  var candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    hojaBase().deleteRow(fila);
    return { ok: true };
  } finally {
    candado.releaseLock();
  }
}

/** Agrupa por ROL+AÑO+LIBRO; cada grupo con más de una causa es un duplicado.
 *  Sugiere conservar la de FECHA INGRESO más antigua. */
function detectarDuplicados(causas) {
  var grupos = {};
  causas.forEach(function (c) {
    var clave = claveCruce(c.rol, c.anio, c.libro);
    if (!grupos[clave]) grupos[clave] = [];
    grupos[clave].push(c);
  });
  var resultado = [];
  Object.keys(grupos).forEach(function (clave) {
    var grupo = grupos[clave];
    if (grupo.length < 2) return;
    grupo.sort(function (a, b) {
      var fa = parsearFecha(a.fechaIngreso); var fb = parsearFecha(b.fechaIngreso);
      fa = fa ? fa.getTime() : Infinity; fb = fb ? fb.getTime() : Infinity;
      return fa - fb;
    });
    resultado.push({ clave: clave, conservar: grupo[0].idFila, causas: grupo });
  });
  return resultado;
}

function reportes() {
  var causas = leerCausas();
  var porEstado = {}, porMateriaPendiente = {}, porTipoPendiente = {};
  var pendientesEstados = { 'Disponible': true, 'En Trámite': true, 'Suspendida Procedimiento': true };

  causas.forEach(function (c) {
    var estado = c.estado || 'Sin estado';
    porEstado[estado] = (porEstado[estado] || 0) + 1;
    if (pendientesEstados[estado]) {
      var materia = c.materia || 'Sin materia';
      var tipo = c.tipo || 'Sin tipo';
      porMateriaPendiente[materia] = (porMateriaPendiente[materia] || 0) + 1;
      porTipoPendiente[tipo] = (porTipoPendiente[tipo] || 0) + 1;
    }
  });

  return {
    ok: true,
    total: causas.length,
    porEstado: porEstado,
    pendientesPorMateria: porMateriaPendiente,
    pendientesPorTipo: porTipoPendiente,
    enTabla: causas.filter(function (c) { return c.estado === 'En Tabla'; }),
    pendientes: causas.filter(function (c) { return pendientesEstados[c.estado]; })
  };
}
