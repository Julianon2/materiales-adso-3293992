/**
 * Apps Script — Registro de Asistencia
 * SENA ADSO · Ficha 3293992 (Lunes, presencial)
 *
 * Ventana de entrada: 6:30pm a 6:45pm (15 min de tolerancia).
 * Después de las 6:45pm el registro se guarda igual, pero marcado como "Tarde".
 * Un mismo aprendiz (nombre) no puede registrar dos veces el mismo día.
 *
 * Valida nombre + número de documento, igual que en otras fichas. Por
 * eso esta carpeta completa está en .gitignore: index.html, este .gs y
 * Aprendices.md viven solo en este computador y nunca se suben al repo
 * público (GitHub Pages), porque contienen cédulas reales.
 *
 * CÓMO DESPLEGAR (una sola vez):
 * 1. Crear un Google Sheet nuevo, nombrarlo "Asistencia - Ficha 3293992".
 * 2. En el Sheet: Extensiones → Apps Script.
 * 3. Borrar el contenido de Code.gs y pegar todo este archivo.
 * 4. Guardar (Ctrl+S), darle un nombre al proyecto (ej. "Registro Asistencia").
 * 5. Implementar → Nueva implementación → tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Autorizar los permisos que pida Google.
 * 7. Copiar la URL que termina en /exec — esa es la URL_APPS_SCRIPT
 *    que va en index.html de esta misma carpeta.
 * 8. Si vuelves a editar este .gs: "Implementar → Administrar
 *    implementaciones → editar (lápiz) → Nueva versión".
 */

const NOMBRE_HOJA = 'Asistencia';

// Hora de inicio y de cierre de la ventana "a tiempo" (24h, hora de Colombia)
const HORA_INICIO = { h: 18, m: 30 }; // 6:30pm
const HORA_CIERRE = { h: 18, m: 45 }; // 6:45pm

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const nombre = (datos.nombre || '').trim();
    const documento = (datos.documento || '').trim();

    if (!nombre || !documento) {
      return respuestaJSON({ ok: false, mensaje: 'Falta nombre o documento.' });
    }

    const hoja = obtenerOCrearHoja();
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

    const estado = calcularEstado(ahora);
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

  } catch (error) {
    return respuestaJSON({ ok: false, mensaje: 'Error: ' + error.message });
  }
}

function doGet(e) {
  const accion = e && e.parameter ? e.parameter.accion : null;

  if (accion === 'hoy') {
    return respuestaJSON({ ok: true, registros: obtenerRegistrosDeHoy() });
  }

  return respuestaJSON({ ok: true, mensaje: 'El script de asistencia está activo. Usa POST para registrar.' });
}

function obtenerRegistrosDeHoy() {
  const hoja = obtenerOCrearHoja();
  const ahora = new Date();
  const fechaHoy = Utilities.formatDate(ahora, 'America/Bogota', 'yyyy-MM-dd');
  const datosHoja = hoja.getDataRange().getValues();
  const registros = [];

  for (let i = 1; i < datosHoja.length; i++) {
    const marca = new Date(datosHoja[i][0]);
    const filaFecha = Utilities.formatDate(marca, 'America/Bogota', 'yyyy-MM-dd');
    if (filaFecha === fechaHoy) {
      registros.push({
        nombre: String(datosHoja[i][1]).trim(),
        documento: String(datosHoja[i][2]).trim(),
        estado: String(datosHoja[i][3]).trim(),
        hora: String(datosHoja[i][4]).trim()
      });
    }
  }

  return registros;
}

function calcularEstado(fecha) {
  const minutosActual = fecha.getHours() * 60 + fecha.getMinutes();
  const minutosCierre = HORA_CIERRE.h * 60 + HORA_CIERRE.m;
  return minutosActual <= minutosCierre ? 'A tiempo' : 'Tarde';
}

function obtenerOCrearHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(['Fecha/Hora', 'Nombre', 'Documento', 'Estado', 'Hora registrada']);
    hoja.setFrozenRows(1);
  }

  return hoja;
}

function respuestaJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
