/*  FINCA LA ESPERANZA — Backend v2 (con sincronización entre teléfonos)
 *
 *  QUÉ CAMBIA: ahora el script no solo RECIBE datos, también los ENTREGA.
 *  Eso permite que varios teléfonos vean lo mismo.
 *
 *  CÓMO INSTALARLO (reemplaza el anterior):
 *  1) Abre tu Google Sheet → Extensiones → Apps Script
 *  2) Borra TODO el código viejo y pega este.
 *  3) Revisa que el TOKEN sea el mismo que usas en la app (1108).
 *  4) Guarda (💾)
 *  5) Implementar → Gestionar implementaciones → ✏️ (editar)
 *       → Versión: "Nueva versión" → Implementar
 *     (Así conservas la misma URL. Si creas una implementación nueva,
 *      la URL cambia y hay que actualizarla en Ajustes de la app.)
 */

const TOKEN = '1108';   // debe coincidir con el token de la app
// v2.1 — corrige que los teléfonos con la hora desfasada no compartieran sus registros

// Pestañas que maneja la app
const TABLAS = ['registros','produccion','animales','leche','salud',
                'inventario','caja','galeria'];


// ---------- RELLENAR recibido VACÍOS (corrige registros viejos) ----------
// Ejecutar UNA VEZ manualmente desde el editor de Apps Script
function rellenarRecibido() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ahora = Date.now();
  TABLAS.forEach(function(nombre) {
    const sh = ss.getSheetByName(cap(nombre));
    if (!sh || sh.getLastRow() < 2) return;
    const datos = sh.getDataRange().getValues();
    const cab = datos[0];
    const iRec = cab.indexOf('recibido');
    const iCre = cab.indexOf('creado');
    if (iRec < 0) return;
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][iRec] || datos[i][iRec] === '') {
        // usar creado si existe, si no usar ahora
        const val = (iCre >= 0 && datos[i][iCre]) ? datos[i][iCre] : ahora;
        sh.getRange(i + 1, iRec + 1).setValue(val);
      }
    }
    SpreadsheetApp.flush();
  });
  return 'Listo — recibido rellenado en todas las pestañas';
}

// ---------- ENTREGAR DATOS (lo nuevo) ----------
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (!p.token) return texto('Finca La Esperanza: OK');
    if (p.token !== TOKEN) return json({ ok: false, error: 'token' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salida = {};
    const desde = p.desde ? Number(p.desde) : 0;   // solo lo modificado después de X

    TABLAS.forEach(function (nombre) {
      const sh = ss.getSheetByName(cap(nombre));
      if (!sh || sh.getLastRow() < 2) { salida[nombre] = []; return; }
      const datos = sh.getDataRange().getValues();
      const cab = datos[0];
      const filas = [];
      for (let i = 1; i < datos.length; i++) {
        const obj = {};
        for (let c = 0; c < cab.length; c++) {
          if (!cab[c]) continue;
          obj[cab[c]] = normalizar(datos[i][c]);
        }
        if (!obj.id) continue;
        // Se filtra por 'recibido' (hora del servidor). Usar 'modificado' (hora del
        // teléfono) haría que un teléfono atrasado nunca comparta sus registros.
        if (desde && Number(obj.recibido || 0) <= desde) continue;
        filas.push(obj);
      }
      salida[nombre] = filas;
    });

    return json({ ok: true, servidor: Date.now(), datos: salida });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ---------- RECIBIR DATOS ----------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return json({ ok: false, error: 'token' });

    const r = body.registro || {};
    if (!r.id) return json({ ok: false, error: 'sin id' });

    let hoja = body.tabla ? cap(body.tabla)
             : (r.categoria === 'Plan' ? 'Planificador' : 'Registros');

    // Foto opcional -> Drive
    let fotoUrl = r.fotoUrl || r.comprobanteUrl || '';
    if (body.foto) {
      try { fotoUrl = guardarFoto(body.foto, r); } catch (err) {}
    }
    if (fotoUrl) { r.fotoUrl = fotoUrl; r.comprobanteUrl = fotoUrl; }
    r.sync = 'ok';
    if (!r.modificado) r.modificado = Date.now();
    r.recibido = Date.now();   // hora del SERVIDOR: es la única igual para todos los teléfonos

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(hoja);

    // Crear la pestaña con encabezados según los campos del registro
    if (!sh) {
      sh = ss.insertSheet(hoja);
      sh.appendRow(Object.keys(r));
      sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    if (sh.getLastRow() === 0) sh.appendRow(Object.keys(r));

    let cab = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];

    // Si el registro trae campos nuevos, agregar columnas
    const faltan = Object.keys(r).filter(function (k) { return cab.indexOf(k) === -1; });
    if (faltan.length) {
      sh.getRange(1, cab.length + 1, 1, faltan.length).setValues([faltan]);
      sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
      cab = cab.concat(faltan);
    }

    const fila = cab.map(function (c) { return r[c] != null ? r[c] : ''; });

    // Buscar si ya existe ese id (evita duplicados y aplica ediciones)
    const colId = cab.indexOf('id') + 1;
    let pos = -1;
    if (colId > 0 && sh.getLastRow() > 1) {
      const ids = sh.getRange(2, colId, sh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(r.id)) { pos = i + 2; break; }
      }
    }
    if (pos > 0) sh.getRange(pos, 1, 1, cab.length).setValues([fila]);
    else sh.appendRow(fila);

    return json({ ok: true, fotoUrl: fotoUrl, servidor: Date.now() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ---------- Foto a Drive, devolviendo enlace que se puede mostrar ----------
function guardarFoto(dataUrl, r) {
  const partes = dataUrl.match(/^data:(.+);base64,(.*)$/);
  if (!partes) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(partes[2]), partes[1]);

  const f = String(r.fecha || '').split('-');
  const anio = f[0] || 'sin-anio';
  const mes  = f[1] || 'sin-mes';

  const raiz  = carpeta(DriveApp.getRootFolder(), 'Finca La Esperanza');
  const comp  = carpeta(raiz, 'Fotos');
  const cAnio = carpeta(comp, anio);
  const cMes  = carpeta(cAnio, mes);

  const nombre = [r.fecha, r.nombreDesc || r.nombre || r.producto, r.id]
    .filter(Boolean).join('_').replace(/[^\w\-.]+/g, '_') + '.jpg';
  blob.setName(nombre);

  const archivo = cMes.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Enlace directo: se puede mostrar como imagen en la app
  return 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w1200';
}

// ---------- utilidades ----------
function carpeta(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}
function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }
function normalizar(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (v === 'TRUE'  || v === true)  return true;
  if (v === 'FALSE' || v === false) return false;
  return v;
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function texto(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
