/* ======================================================================
   APP.JS — Panel Auditor
   Controlador principal del sistema
   - Cambio entre módulos
   - Carga de datos
   - Render de tabla
   - Acciones: ver, aprobar
   - Uso de modulos.js, auth.js y graph.js

   Ariel-friendly: limpio, comentado y escalable
   ====================================================================== */

import {
  listarArchivosMCI,
  descargarArchivo,
  formatearFecha,
  formatearTamano,
  obtenerModulo,
  MODULOS
} from './modulos_v2.js';

import { cargarDesdeCarpeta, obtenerURLTemporal, moverArchivo } from "./graph_v2.js";

import { iniciarSesion, usuarioActual, cerrarSesion, obtenerToken } from "./auth.js";

/* ======================================================================
   ESTADO GLOBAL
   ====================================================================== */
let moduloActivo = null;
let datosActuales = [];

/* ======================================================================
   1) INICIALIZACIÓN GENERAL
   ====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {

  if (!usuarioActual()) {
    await iniciarSesion();
  }

  prepararSidebar();
  seleccionarModulo("inicio");
});

/* ======================================================================
   2) CONFIGURAR SIDEBAR
   ====================================================================== */
function prepararSidebar() {
  const botones = document.querySelectorAll(".sb-item");

  botones.forEach(btn => {
    btn.addEventListener("click", async () => {

      if (btn.classList.contains("logout")) {
        cerrarSesion();
        return;
      }

      botones.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");

      const mod = btn.dataset.mod;
      seleccionarModulo(mod);
    });
  });
}

/* ======================================================================
   3) CAMBIAR DE MÓDULO
   ====================================================================== */
async function seleccionarModulo(mod) {

  const contenedor = document.getElementById("contenedor-modulo");
  contenedor.innerHTML = "";

  if (mod === "inicio") {
    moduloActivo = null;

    contenedor.innerHTML = `
      <div style="padding:20px; font-size:16px;">
        Bienvenido al <strong>Panel Auditor</strong>.<br>
        Selecciona un módulo en la barra lateral para comenzar.
      </div>
    `;
    return;
  }

  moduloActivo = obtenerModulo(mod);

  if (!moduloActivo) {
    contenedor.innerHTML = "<p>Error: módulo desconocido.</p>";
    return;
  }

  contenedor.innerHTML = generarTablaHTML(moduloActivo);

  await cargarDatosModulo();
}

/* ======================================================================
   4) CREAR TABLA
   ====================================================================== */
function generarTablaHTML(modulo) {

  const ths = modulo.columnas
    .map(col => `<th>${col.label}</th>`)
    .join("");

  return `
    <div class="tabla-box">
      <table class="tabla">
        <thead>
          <tr>${ths}<th>Acciones</th></tr>
        </thead>
        <tbody id="tbodyDatos">
          <tr><td colspan="${modulo.columnas.length + 1}" 
              style="text-align:center; padding:20px;">Cargando…</td></tr>
        </tbody>
      </table>
    </div>
  `;
}

/* ======================================================================
   5) CARGAR DATOS DESDE ONEDRIVE
   ====================================================================== */
async function cargarDatosModulo() {

  if (!moduloActivo.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
      <tr><td colspan="99" style="padding:20px; text-align:center;">
        No hay ruta configurada para este módulo.<br>
        (Ariel deberá especificarla cuando toque)
      </td></tr>
    `;
    return;
  }

  datosActuales = await cargarDesdeCarpeta(moduloActivo, false);

  renderTabla();
}

/* ======================================================================
   6) RENDER DE TABLA
   ====================================================================== */
function renderTabla() {

  const tbody = document.getElementById("tbodyDatos");

  if (!datosActuales || datosActuales.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="99" style="text-align:center; padding:20px;">
          No hay informes pendientes.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = "";

  datosActuales.forEach((item, idx) => {

    const tds = moduloActivo.columnas
      .map(col => `<td>${item[col.id]}</td>`)
      .join("");

    const tr = document.createElement("tr");

    tr.innerHTML = `
      ${tds}
      <td>
        <button class="btn-ver" data-idx="${idx}" style="margin-right:6px;">Ver</button>
        <button class="btn-aprobar" data-idx="${idx}">Aprobar</button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  prepararEventosTabla();
}

/* ======================================================================
   7) EVENTOS DE TABLA
   ====================================================================== */
function prepararEventosTabla() {

  document.querySelectorAll(".btn-ver").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await verArchivo(item);
    });
  });

  document.querySelectorAll(".btn-aprobar").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await aprobarArchivo(item);
    });
  });
}

/* ======================================================================
   8) VER ARCHIVO — VISOR MODAL SHAREPOINT
   ====================================================================== */
async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  window.__archivoActual = item;

  const token = await obtenerToken();
  if (!token) {
    alert("No se pudo obtener token.");
    return;
  }

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0${item.archivo.ruta}`,
    { headers: { "Authorization": `Bearer ${token}` } }
  );

  const data = await resp.json();
  if (!data?.webUrl) {
    alert("No se pudo obtener la URL del archivo.");
    return;
  }

  const embedUrl = `${data.webUrl}?web=1&action=embedview`;

  document.getElementById("visorIframe").innerHTML = `
    <iframe 
        src="${embedUrl}"
        width="100%"
        height="100%"
        frameborder="0"
        allowfullscreen
        style="border:0; background:white;"
    ></iframe>
  `;
}

/* ======================================================================
   9) APROBAR (MOVER ARCHIVO)
   ====================================================================== */
async function aprobarArchivo(item) {

  if (!moduloActivo.aprobados) {
    alert("No hay carpeta de aprobados configurada.");
    return;
  }

  const r1 = item.archivo.ruta;
  const r2 = `${moduloActivo.aprobados}/${item.archivo.nombre}`;

  const ok = await moverArchivo(r1, r2);

  if (!ok) {
    alert("Error moviendo archivo.");
    return;
  }

  alert(`✅ Informe aprobado: ${item.archivo.nombre}`);

  await cargarDatosModulo();
}

/* ======================================================================
   10) EVENTOS DEL MODAL (Cerrar / Descargar / Aprobar / Rechazar)
   ====================================================================== */

// ✅ Cerrar visor
document.getElementById("visorVolver").addEventListener("click", () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  document.getElementById("visorIframe").innerHTML = "";
});

// ✅ Descargar archivo desde modal
document.getElementById("visorDescargar").addEventListener("click", async () => {

  const item = window.__archivoActual;
  if (!item) return;

  const token = await obtenerToken();
  const url = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;

  const resp = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const blob = await resp.blob();
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = item.archivo.nombre;
  link.click();
});

// ✅ Aprobar desde modal
document.getElementById("visorAprobar").addEventListener("click", async () => {
  const item = window.__archivoActual;
  if (!item) return;

  await aprobarArchivo(item);
  document.getElementById("visorVolver").click();
});

// ✅ Rechazar (placeholder)
document.getElementById("visorRechazar").addEventListener("click", () => {
  alert("Función de rechazo pendiente.");
});
