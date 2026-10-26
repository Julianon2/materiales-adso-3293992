/**
 * Apps Script — Registro de progreso S1 (Tailwind CSS)
 * SENA ADSO · Ficha 3293992
 *
 * CÓMO DESPLEGAR (una sola vez):
 * 1. Crear un Google Sheet nuevo, nombrarlo "Progreso S1 - Tailwind (Ficha 3293992)".
 * 2. En el Sheet: Extensiones → Apps Script.
 * 3. Borrar el contenido de Code.gs y pegar todo este archivo.
 * 4. Guardar (Ctrl+S), darle un nombre al proyecto (ej. "Progreso S1").
 * 5. Implementar → Nueva implementación → tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Autorizar los permisos que pida Google (es tu propio script, es seguro).
 * 7. Copiar la URL que termina en /exec — esa es la URL_APPS_SCRIPT
 *    que va en 02_GUIA_TAILWIND_INTRO.html y 03_EJERCICIOS_PRACTICA.html.
 * 8. Si vuelves a editar este .gs, hay que hacer "Implementar → Administrar
 *    implementaciones → editar (lápiz) → Nueva versión" para que el cambio
 *    se refleje en la URL ya usada.
 */

const NOMBRE_HOJA = 'Progreso';

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const hoja = obtenerOCrearHoja();

    const filaNueva = [
      new Date(),                          // Fecha y hora de envío
      datos.aprendiz || '(sin nombre)',     // Nombre del aprendiz
      datos.documento || '',                // Documento (opcional, si se agrega luego)
      datos.archivo || '',                  // "02_GUIA_TAILWIND_INTRO" o "03_EJERCICIOS_PRACTICA"
      datos.completados || 0,               // Cuántos pasos marcó
      datos.total || 0,                     // Cuántos pasos tiene el archivo
      (datos.pasos || []).join(', ')        // Lista de IDs de pasos completados
    ];

    hoja.appendRow(filaNueva);

    return respuestaJSON({ ok: true, mensaje: 'Progreso registrado' });
  } catch (error) {
    return respuestaJSON({ ok: false, mensaje: 'Error: ' + error.message });
  }
}

function doGet(e) {
  return respuestaJSON({ ok: true, mensaje: 'El script está activo. Usa POST para enviar datos.' });
}

function obtenerOCrearHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(['Fecha/Hora', 'Aprendiz', 'Documento', 'Archivo', 'Pasos completados', 'Total pasos', 'Detalle pasos']);
    hoja.setFrozenRows(1);
  }

  return hoja;
}

function respuestaJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
