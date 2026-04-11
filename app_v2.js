import { SITE_ID, LIBRARY_ID, FOLDER_PATH } from "./modulos_nuevo.js";
import { obtenerToken, iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";
import { obtenerURLTemporal, moverArchivo } from "./graph_actual.js";
// ======================================================================
// BUSCAR EL JSON ASOCIADO AL EXCEL EN LA MISMA CARPETA
// ======================================================================
async function obtenerJsonFotos(item) {

  const token = await obtenerToken();

  // 1. Listamos TODOS los archivos de la carpeta MCI_Generados
  const urlListar = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${encodeURIComponent(FOLDER_PATH)}:/children`;

  const resp = await fetch(urlListar, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const data = await resp.json();
  if (!data.value) return null;

  // 2. Nombre base del excel
  const base = item.nombre;   // ejemplo: MCI_XXXX.xlsx

  // 3. Buscar un archivo cuyo nombre empiece igual y termine en ".json"
  const jsonFile = data.value.find(f =>
    f.name.startsWith(base) &&
    f.name.endsWith(".json")
  );

  if (!jsonFile) {
    console.warn("No se encontró JSON para fotos.");
    return null;
  }

  // 4. Descargar el JSON
  const urlContenido = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${jsonFile.id}/content`;

  const respJson = await fetch(urlContenido, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const jsonTexto = await respJson.text();

  try {
    return JSON.parse(jsonTexto);
  } catch (e) {
    console.error("Error parsing JSON:", e);
    return null;
  }
}
/* ======================================================================
   9) VISOR — PREVIEW XLSX (Excel) + CARGA DEL JSON DE FOTOS
====================================================================== */
async function verArchivo(item) {

  // Mostrar modal
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  window.__archivoActual = item;
  window.__mciIdActual = item.mciId ?? null;

  const token = await obtenerToken();

  // ============================================================
  // ✅ 1. DESCARGAR EL EXCEL DESDE SHAREPOINT
  // ============================================================
  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();
  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // ============================================================
  // ✅ 2. PREVIEW DEL EXCEL USANDO SHEETJS
  // ============================================================
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const visor = document.getElementById("visorIframe");
  visor.innerHTML = ""; // limpiar

  let html = `<h3 style="margin:0 0 10px 0;">Vista previa del informe (Excel)</h3>`;

  html += `<table style="
      width:100%;
      border-collapse:collapse;
      font-family:Segoe UI, sans-serif;
      font-size:14px;
      margin-bottom:25px;
  ">`;

  rows.forEach(row => {
    html += `<tr>`;
    row.forEach(cell => {
      html += `<td style="
          border:1px solid #d0d7e7;
          padding:6px 10px;
          background:white;
      ">${cell ?? ""}</td>`;
    });
    html += `</tr>`;
  });

  html += `</table>`;

  // Ponemos temporalmente el HTML del Excel (fotos vienen después)
  visor.innerHTML = html;

  // ============================================================
  // ✅ 3. AHORA BUSCAR EL JSON ASOCIADO (PARTE 1)
  // ============================================================
  const jsonFotos = await obtenerJsonFotos(item);

  // Guardamos fotos, aunque sean null
  item.fotosPreview = jsonFotos;

  // La Parte 3 (próxima que te envío) dibuja las fotos.
  if (jsonFotos) {
    visor.innerHTML += `<h3 style="margin-top:20px;">Fotos registradas</h3>`;
  } else {
    visor.innerHTML += `<p style="color:#888;">(Este informe no tiene fotos asociadas)</p>`;
  }
}
/* ======================================================================
   9-B) RENDERIZAR FOTOS DEL JSON (Galería dentro del visor)
====================================================================== */
async function renderizarFotos(item) {

  const visor = document.getElementById("visorIframe");

  // Si no hay JSON de fotos, no se hace nada
  if (!item.fotosPreview) {
    visor.innerHTML += `<p style="color:#888;">(Este informe no tiene fotos asociadas)</p>`;
    return;
  }

  const fotos = item.fotosPreview; // Objeto JSON con fotos base64

  // Área de galería
  let htmlFotos = `
    <h3 style="margin-top:20px;">Fotos del informe</h3>
    <div style="
      display:flex;
      flex-wrap:wrap;
      gap:20px;
      margin-top:10px;
    ">
  `;

  // Recorre cada propiedad del JSON: gps, apInt, apExt1, ...
  for (const clave in fotos) {
    const base64 = fotos[clave];    // valor tipo "data:image/jpeg;base64,..."

    if (!base64) continue;

    htmlFotos += `
      <div style="
        width:260px;
        border:1px solid #d0d7e7;
        border-radius:10px;
        padding:10px;
        background:white;
      ">
        <img src="${base64}" style="width:100%; border-radius:8px;" />
        <div style="text-align:center; margin-top:8px; font-size:13px;">
          ${clave}
        </div>
      </div>
    `;
  }

  htmlFotos += `</div>`;

  // Agregar fotos al visor bajo el Excel
  visor.innerHTML += htmlFotos;
}
/* ======================================================================
   10) BOTÓN "VOLVER" — CERRAR VISOR Y REGRESAR A LA TABLA
====================================================================== */
document.getElementById("visorVolver").addEventListener("click", () => {

  // Ocultar modal
  document.getElementById("modalVisor").style.display = "none";

  // Mostrar tabla otra vez
  document.getElementById("contenedor-modulo").style.display = "block";

  // Volver a dibujar tabla
  renderTabla();
});


/* ======================================================================
   11) APROBAR — ACTUALIZA KV Y CIERRA VISOR
====================================================================== */
document.getElementById("visorAprobar").addEventListener("click", async () => {

  const mciIdReal = window.__mciIdActual;

  if (!mciIdReal) {
    alert("❌ No se encontró el mciId para este informe.");
    return;
  }

  // Llamar API KV para aprobar
  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciIdReal}`,
    { method: "PUT" }
  );

  // Cerrar visor
  document.getElementById("visorVolver").click();

  // Actualizar tabla
  renderTabla();
});


/* ======================================================================
   12) BOTÓN "RECHAZAR" — (OPCIONAL)
====================================================================== */
document.getElementById("visorRechazar").addEventListener("click", async () => {

  const mciIdReal = window.__mciIdActual;

  if (!mciIdReal) {
    alert("❌ No se encontró el mciId para este informe.");
    return;
  }

  // Rechazar en KV (si existiera endpoint, ejemplo)
  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/rechazar/${mciIdReal}`,
    { method: "PUT" }
  );

  // Cerrar visor
  document.getElementById("visorVolver").click();

  // Actualizar tabla
  renderTabla();
});
