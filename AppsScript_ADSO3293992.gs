/**
 * Apps Script único — Asistencia y Seguimiento de actividades
 * SENA ADSO · Ficha 3293992 (Lunes, presencial)
 *
 * Un solo libro de Google Sheets, un solo script, una sola URL /exec.
 * Según el campo "tipo" que llega en el POST, escribe en una hoja
 * ("Asistencia") o en la otra ("Progreso") dentro del mismo libro.
 *
 * Reemplaza a AppsScript_RegistroAsistencia.gs y a AppsScript_Progreso_S1.gs:
 * ya no se despliegan por separado, se usa este único archivo.
 *
 * CÓMO DESPLEGAR (una sola vez):
 * 1. Crear un Google Sheet nuevo, nombrarlo "ADSO 3293992 - Asistencia y Progreso".
 * 2. En el Sheet: Extensiones → Apps Script.
 * 3. Borrar el contenido de Code.gs y pegar todo este archivo.
 * 4. Guardar (Ctrl+S), darle un nombre al proyecto (ej. "ADSO 3293992").
 * 5. Implementar → Nueva implementación → tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Autorizar los permisos que pida Google (es tu propio script, es seguro).
 * 7. Copiar la URL que termina en /exec — esa MISMA URL va en los 3 archivos:
 *      - REGISTRO_ASISTENCIA/index.html       (const URL_APPS_SCRIPT)
 *      - FASE_REACT_TAILWIND/.../02_GUIA_TAILWIND_INTRO.html
 *      - FASE_REACT_TAILWIND/.../03_EJERCICIOS_PRACTICA.html
 * 8. Si vuelves a editar este .gs: "Implementar → Administrar
 *    implementaciones → editar (lápiz) → Nueva versión" para que el
 *    cambio se refleje en la URL ya usada — no genera una URL nueva.
 *
 * Las dos hojas ("Asistencia" y "Progreso") se crean solas la primera
 * vez que llega un registro de cada tipo; no hace falta crearlas a mano.
 */

const HOJA_ASISTENCIA = 'Asistencia';
const HOJA_PROGRESO   = 'Progreso';

// Ventana de entrada de la asistencia (24h, hora de Colombia)
const HORA_CIERRE = { h: 18, m: 45 }; // 6:45pm — después de esto, "Tarde"

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);

    if (datos.tipo === 'asistencia') return registrarAsistencia(datos);
    if (datos.tipo === 'progreso')   return registrarProgreso(datos);

    return respuestaJSON({ ok: false, mensaje: 'Falta el campo "tipo" (asistencia | progreso).' });
  } catch (error) {
    return respuestaJSON({ ok: false, mensaje: 'Error: ' + error.message });
  }
}

function doGet(e) {
  return respuestaJSON({ ok: true, mensaje: 'El script está activo. Usa POST con { tipo: "asistencia" | "progreso", ... }.' });
}

/* ═══ ASISTENCIA ═══ */
function registrarAsistencia(datos) {
  const nombre = (datos.nombre || '').trim();
  const documento = (datos.documento || '').trim();

  if (!nombre || !documento) {
    return respuestaJSON({ ok: false, mensaje: 'Falta nombre o documento.' });
  }

  const hoja = obtenerOCrearHoja(HOJA_ASISTENCIA, ['Fecha/Hora', 'Nombre', 'Documento', 'Estado', 'Hora registrada']);
  const ahora = new Date();
  const fechaHoy = Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM-dd');

  // Revisar si ya existe un registro de este aprendiz hoy
  const datosHoja = hoja.getDataRange().getValues();
  for (let i = 1; i < datosHoja.length; i++) {
    const filaFecha = Utilities.formatDate(new Date(datosHoja[i][0]), 'America/Bogota', 'yyyy-MM-dd');
    const filaNombre = String(datosHoja[i][1]).trim();
    if (filaFecha === fechaHoy && filaNombre.toLowerCase() === nombre.toLowerCase()) {
      const horaPrevia = Utilities.formatDate(new Date(datosHoja[i][0]), 'America/Bogota', 'hh:mm a');
      return respuestaJSON({
        ok: false,
        yaRegistrado: true,
        mensaje: 'Ya registraste tu entrada hoy a las ' + horaPrevia + '.'
      });
    }
  }

  const estado = calcularEstadoAsistencia(ahora);
  const horaTexto = Utilities.formatDate(ahora, 'America/Bogota', 'hh:mm:ss a');

  hoja.appendRow([ahora, nombre, documento, estado, horaTexto]);

  return respuestaJSON({
    ok: true,
    estado: estado,
    hora: horaTexto,
    mensaje: estado === 'A tiempo'
      ? 'Entrada registrada a las ' + horaTexto + ' — ¡a tiempo!'
      : 'Entrada registrada a las ' + horaTexto + ' — llegaste tarde.'
  });
}

function calcularEstadoAsistencia(fecha) {
  const minutosActual = fecha.getHours() * 60 + fecha.getMinutes();
  const minutosCierre = HORA_CIERRE.h * 60 + HORA_CIERRE.m;
  return minutosActual <= minutosCierre ? 'A tiempo' : 'Tarde';
}

/* ═══ PROGRESO / SEGUIMIENTO DE ACTIVIDADES ═══ */
function registrarProgreso(datos) {
  const hoja = obtenerOCrearHoja(HOJA_PROGRESO, ['Fecha/Hora', 'Aprendiz', 'Documento', 'Archivo', 'Pasos completados', 'Total pasos', 'Detalle pasos']);

  hoja.appendRow([
    new Date(),
    datos.aprendiz || '(sin nombre)',
    datos.documento || '',
    datos.archivo || '',
    datos.completados || 0,
    datos.total || 0,
    (datos.pasos || []).join(', ')
  ]);

  return respuestaJSON({ ok: true, mensaje: 'Progreso registrado' });
}

/* ═══ Utilidades compartidas ═══ */
function obtenerOCrearHoja(nombreHoja, encabezados) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = libro.insertSheet(nombreHoja);
    hoja.appendRow(encabezados);
    hoja.setFrozenRows(1);
  }

  return hoja;
}

function respuestaJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
