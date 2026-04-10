/* ======================================================================
   APP.JS — Panel Auditor (Versión corregida completa)
====================================================================== */

/* ============================
   VARIABLES GLOBALES REALES
============================ */
window.moduloActivo = null;
window.datosActuales = [];
window.estadoInformes = {};


/* ============================
   IMPORTS
============================ */
import { listarArchivosMCI, obtenerModulo } from "./modulos_v2.js";
import { obtenerToken, iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";
import { moverArchivo, obtenerURLTemporal } from "./graph_v2.js";


/* ============================
   ESTADO LOCAL
============================ */
function guardarEstados() {
  localStorage.setItem("estadoInformesAuditor", JSON.stringify(window.estadoInformes));
}

function cargarEstados() {
  const raw = localStorage.getItem("estadoInformesAuditor");
  if (raw) {
    try { window.estadoInformes = JSON.parse(raw); }
    catch { window.estadoInformes = {}; }
  }
}


/* ======================================================================
   1) INICIALIZACIÓN
====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {
  if (!usuarioActual()) await iniciarSesion();
  prepararSidebar();
  cargarEstados();
  seleccionarModulo("inicio");
});


/* ======================================================================
   2) SIDEBAR
====================================================================== */
function prepararSidebar() {
  const botones = document.querySelectorAll(".sb-item");

  botones.forEach(btn => {
    btn.addEventListener("click", () => {

      if (btn.classList.contains("logout")) {
        cerrarSesion();
        return;
      }

      botones.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");

      seleccionarModulo(btn.dataset.mod);
    });
  });
}


/* ======================================================================
   3) SELECTOR DE MÓDULO
====================================================================== */
async function seleccionarModulo(mod) {

  const cont = document.getElementById("contenedor-modulo");
  cont.innerHTML = "";

  if (mod === "inicio") {
    window.moduloActivo = null;
    cont.innerHTML = `
      <div style="padding:20px; font-size:16px;">
        Bienvenido al <strong>Panel Auditor</strong>.<br>
        Selecciona un módulo en la barra lateral para comenzar.
      </div>`;
    return;
  }

  window.moduloActivo = obtenerModulo(mod);

  if (!window.moduloActivo) {
    cont.innerHTML = "<p>Error: módulo no encontrado.</p>";
    return;
  }

  cont.innerHTML = generarTablaHTML(window.moduloActivo);

  prepararEventosTabla();
  await cargarDatosModulo();
}
/* ======================================================================
   4) CARGAR DATOS DEL MÓDULO — SharePoint + KV
====================================================================== */
async function cargarDatosModulo() {

  if (!window.moduloActivo.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
    <tr><td colspan="99" style="padding:20px; text-align:center;">
      No hay ruta configurada para este módulo.
    </td></tr>`;
    return;
  }

  const token = await obtenerToken();

  // ✅ Lista real desde SharePoint
  const listaOD = await listarArchivosMCI(token);
  window.debugListaOD = listaOD;

  // ✅ Lista desde KV
  const tecnico = "usuario";
  const respKV = await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
  );
  const listaKV = await respKV.json();

  // Mezcla SP + KV
  for (const a of listaOD) {
    const reg = listaKV.find(k => k.fileName === a.nombre);
    a.mciId = reg ? reg.mciId : null;
    a.estadoKV = reg ? reg.estado : "pendiente";
  }

  // ✅ GLOBAL REAL
  window.datosActuales = listaOD;

  renderTabla();
  setTimeout(() => activarOrdenamientoFecha(), 0);
}


/* ======================================================================
   5) GENERAR TABLA
====================================================================== */
function generarTablaHTML(modulo) {

  const ths = modulo.columnas.map(col => {
    if (col.id === "fecha") {
      return `
      <th style="cursor:pointer;">
        <span class="sortable" data-col="fecha" data-order="desc">
          ${col.label} <span class="flecha">🔽</span>
        </span>
      </th>`;
    }
    return `<th>${col.label}</th>`;
  }).join("");

  return `
  <div class="tabla-box">
    <table class="tabla">
      <thead><tr>${ths}<th>Acciones</th></tr></thead>
      <tbody id="tbodyDatos">
        <tr><td colspan="${modulo.columnas.length + 1}" style="text-align:center; padding:20px;">
          Cargando…
        </td></tr>
      </tbody>
    </table>
  </div>`;
}


/* ======================================================================
   6) RENDER TABLA
====================================================================== */
function renderTabla() {

  const tbody = document.getElementById("tbodyDatos");

  if (!window.datosActuales || window.datosActuales.length === 0) {
    tbody.innerHTML = `
    <tr><td colspan="99" style="padding:20px; text-align:center;">
      No hay informes pendientes.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = "";

  const filtrados = window.datosActuales.filter(item =>
    item.nombre.endsWith(".xlsx") &&
    !item.nombre.includes("PreviewFotos")
  );

  filtrados.forEach(item => {
    const idxReal = window.datosActuales.indexOf(item);

    const tds = window.moduloActivo.columnas
      .map(col => `<td>${item[col.id]}</td>`).join("");

    const estado = item.estadoKV ?? "pendiente";

    const estadoBoton =
      estado === "pendiente"      ? `<button class="btn-estado btn-gris btn-revisar" data-idx="${idxReal}">Revisar</button>` :
      estado === "en_revision"    ? `<button class="btn-estado btn-azul btn-revisar" data-idx="${idxReal}">✏️ Continuar revisión</button>` :
      estado === "aprobado"       ? `<button class="btn-estado btn-verde" disabled>✅ Aprobado</button>` :
      `<button class="btn-estado btn-rojo" disabled>⚠️ Pendiente por técnico</button>`;

    const tr = document.createElement("tr");
    tr.innerHTML = `${tds}<td style="text-align:center;">${estadoBoton}</td>`;
    tbody.appendChild(tr);
  });

  activarOrdenamientoFecha();
  prepararEventosTabla();
}


/* ======================================================================
   7) ORDENAMIENTO POR FECHA
====================================================================== */
function activarOrdenamientoFecha() {

  const th = document.querySelector("span.sortable[data-col='fecha']");
  if (!th) return;

  th.onclick = () => {

    const estado = th.dataset.order ?? "desc";

    window.datosActuales.sort((a, b) => {
      const fA = new Date(a.fechaReal);
      const fB = new Date(b.fechaReal);
      return estado === "desc" ? fA - fB : fB - fA;
    });

    th.dataset.order = (estado === "desc" ? "asc" : "desc");

    th.querySelector(".flecha").textContent =
      estado === "desc" ? "🔽" : "🔼";

    renderTabla();
  };
}
/* ======================================================================
   8) EVENTOS DE TABLA
====================================================================== */
function prepararEventosTabla() {
  document.querySelectorAll(".btn-revisar").forEach(btn => {
    btn.addEventListener("click", async () => {
      const idx = btn.dataset.idx;
      const item = window.datosActuales[idx];
      await verArchivo(item);
    });
  });
}


/* ======================================================================
   9) VISOR — PREVIEW SheetJS (COMPLETO)
====================================================================== */
async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  window.__archivoActual = item;
  window.__mciIdActual = item.mciId ?? null;

  const token = await obtenerToken();

  // Descargar Excel
  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();

  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const visor = document.getElementById("visorIframe");
  visor.innerHTML = "";

  let html = `<table style="
      width:100%;
      border-collapse:collapse;
      font-family:Segoe UI, sans-serif;
      font-size:14px;
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

  visor.innerHTML = html;
}
/* ======================================================================
   10) APROBAR
====================================================================== */
document.getElementById("visorAprobar").addEventListener("click", async () => {

  const mciIdReal = window.__mciIdActual;

  if (!mciIdReal) {
    alert("❌ No se encontró el mciId para este informe.");
    return;
  }

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciIdReal}`,
    { method: "PUT" }
  );

  document.getElementById("visorVolver").click();
  renderTabla();
});
