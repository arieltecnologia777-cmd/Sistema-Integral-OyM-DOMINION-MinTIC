/* ======================================================================
   0) CONFIGURACIÓN — FLUJOS ONEDRIVE
====================================================================== */

// ✅ Flow SOLO para descargar / previsualizar Excel
const FLOW_GET_ONEDRIVE_FILE =
  "https://defaulte4e1bc33e2834312bb3789010224b7.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/bd9e2227be594ecdb47c0da4a898d474/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=j3SlbYcxilxwhnHJfL95lpTA-Y2RzAtiNrmug_D01eQ";

// ✅ Flow SOLO para obtener JSON de fotos (base64)
const FLOW_GET_FOTOS_PREVIEW =
  "https://defaulte4e1bc33e2834312bb3789010224b7.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/dc99f30c70a64d57b309dce1c13d1290/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NMxJMh4pAr98EPpIwDJGzHb5_glsVkAv-A1TVjR9zsA";

/* ======================================================================
   IMPORTS
====================================================================== */
import { obtenerModulo } from "./modulos_v2.js";
import { iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";

/* ======================================================================
   VARIABLES GLOBALES
====================================================================== */
window.moduloActivo = null;
window.datosActuales = [];
window.__archivoActual = null;

/* ======================================================================
   INICIO
====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {
  if (!usuarioActual()) iniciarSesion().catch(() => {});
  prepararSidebar();
  seleccionarModulo("inicio");
});

/* ======================================================================
   SIDEBAR
====================================================================== */
function prepararSidebar() {
  document.querySelectorAll(".sb-item").forEach(btn => {
    btn.addEventListener("click", () => {
      if (btn.classList.contains("logout")) {
        cerrarSesion();
        return;
      }
      document.querySelectorAll(".sb-item").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      seleccionarModulo(btn.dataset.mod);
    });
  });
}

/* ======================================================================
   SELECCIONAR MÓDULO
====================================================================== */
async function seleccionarModulo(mod) {
  const cont = document.getElementById("contenedor-modulo");
  cont.innerHTML = "";

  if (mod === "inicio") {
    cont.innerHTML = `<div style="padding:20px">Bienvenido al <b>Panel Auditor</b></div>`;
    return;
  }

  window.moduloActivo = obtenerModulo(mod);
  cont.innerHTML = generarTablaHTML(window.moduloActivo);
  await cargarDatosModulo();
}

/* ======================================================================
   TABLA
====================================================================== */
function generarTablaHTML(modulo) {
  const ths = modulo.columnas.map(c => `<th>${c.label}</th>`).join("");
  return `
    <table class="tabla">
      <thead><tr>${ths}<th>Acciones</th></tr></thead>
      <tbody id="tbodyDatos"></tbody>
    </table>`;
}

/* ======================================================================
   CARGAR DATOS (KV)
====================================================================== */
async function cargarDatosModulo() {
  const resp = await fetch(`https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar`);
  const listaKV = await resp.json();

  // ✅ Mapeo correcto (evita undefined)
  window.datosActuales = listaKV.map(reg => {
    const fechaObj = reg.fechaGenerado ? new Date(reg.fechaGenerado) : null;
    return {
      nombre: reg.fileName,
      fecha: fechaObj ? fechaObj.toLocaleString("es-CO") : "",
      tamano: "", // KV no trae tamaño; se deja vacío
      mciId: reg.mciId,
      estadoKV: reg.estado,
      fileIdentifierExcel: reg.fileIdentifierExcel,
      jsonFileId: reg.jsonFileId,
      fechaReal: fechaObj
    };
  });

  renderTabla();
}

/* ======================================================================
   RENDER TABLA
====================================================================== */
function renderTabla() {
  const tbody = document.getElementById("tbodyDatos");
  tbody.innerHTML = "";

  window.datosActuales.forEach((item, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.nombre}</td>
      <td>${item.fecha}</td>
      <td>${item.tamano}</td>
      <td><button class="btn-revisar" data-idx="${idx}">Revisar</button></td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll(".btn-revisar").forEach(btn => {
    btn.onclick = () => verArchivo(window.datosActuales[btn.dataset.idx]);
  });
}
/* ======================================================================
   OBTENER JSON DE FOTOS
====================================================================== */
async function obtenerJsonFotos(item) {
  const resp = await fetch(FLOW_GET_FOTOS_PREVIEW, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fileId: item.jsonFileId })
  });
  if (!resp.ok) return null;
  const { fileB64 } = await resp.json();
  return JSON.parse(atob(fileB64));
}

/* ======================================================================
   VER ARCHIVO (EXCEL)
====================================================================== */
async function verArchivo(item) {
  window.__archivoActual = item;

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  // ✅ Excel SIEMPRE desde el flow de Excel
  const resp = await fetch(FLOW_GET_ONEDRIVE_FILE, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fileId: item.fileIdentifierExcel })
  });

  if (!resp.ok) throw new Error("No se pudo obtener el Excel");

  const blob = await resp.blob();
  const wb = XLSX.read(await blob.arrayBuffer());

  // ✅ Leer la hoja donde escribe el script (última)
  const sheet = wb.Sheets[wb.SheetNames[wb.SheetNames.length - 1]];

  let htmlInfoGeneral = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B9:P18" });
  let htmlDescripcion = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B69:P69" });
  let htmlDeclaracion = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B71:M77" });

  const visor = document.getElementById("visorIframe");
  visor.innerHTML = `
    <div style="background:white;padding:25px;border-radius:14px;border:1px solid #dce3f5;box-shadow:0 8px 24px rgba(0,0,0,.12);">
      <div style="background:#eef1f6;padding:14px 18px;border-radius:10px;font-weight:800;margin-bottom:14px;">
        Información del Beneficiario y la Institución
      </div>
      <div class="auditor-block">${htmlInfoGeneral}</div>

      <div style="background:#eef1f6;padding:14px 18px;border-radius:10px;font-weight:800;margin:28px 0 14px;">
        Descripción del Caso
      </div>
      <div class="auditor-block">${htmlDescripcion}</div>

      <div style="background:#eef1f6;padding:14px 18px;border-radius:10px;font-weight:800;margin:28px 0 14px;">
        Declaración
      </div>
      <div class="auditor-block">${htmlDeclaracion}</div>

      <h2 style="margin:30px 0 10px;">Fotos del informe (vista previa)</h2>
      <div id="visorFotos"></div>
    </div>
  `;

  const fotos = await obtenerJsonFotos(item);
  if (fotos) renderizarFotos(fotos);
  else document.getElementById("visorFotos").innerHTML =
    "<p style='color:#777;'>Este informe no tiene fotos adjuntas.</p>";
}

/* ======================================================================
   RENDER FOTOS
====================================================================== */
function renderizarFotos(fotos) {
  const cont = document.getElementById("visorFotos");
  cont.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:22px;width:100%;">
      ${Object.keys(fotos).map(k => `
        <div style="background:#fff;border:1px solid #dde5f8;border-radius:12px;box-shadow:0 6px 15px rgba(0,0,0,.08);overflow:hidden;">
          <div style="padding:10px 12px;font-weight:700;font-size:14px;">${k}</div>
          <img src="${fotos[k]}" style="width:100%;height:180px;object-fit:cover;display:block;">
        </div>
      `).join("")}
    </div>
  `;
}

/* ======================================================================
   VOLVER
====================================================================== */
document.getElementById("visorVolver").addEventListener("click", () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  renderTabla();
});

/* ======================================================================
   APROBAR / RECHAZAR
====================================================================== */
document.getElementById("visorAprobar").addEventListener("click", async () => {
  const mciId = window.__archivoActual?.mciId;
  if (!mciId) return;
  await fetch(`https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciId}`, { method: "PUT" });
  await cargarDatosModulo();
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
});

document.getElementById("visorRechazar").addEventListener("click", async () => {
  const mciId = window.__archivoActual?.mciId;
  if (!mciId) return;
  await fetch(`https://cloudflare-index.modulo-de-exclusiones.workers.dev/rechazar/${mciId}`, { method: "PUT" });
});
