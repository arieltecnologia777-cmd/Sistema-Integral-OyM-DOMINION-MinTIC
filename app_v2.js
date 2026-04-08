/* ======================================================================
   APP.JS — Panel Auditor
   Controlador principal del sistema
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


// ✅ Convierte "2/4/2026, 9:52:10 a. m." → Date real
function parseFechaCol(fechaStr) {
  if (!fechaStr) return new Date(0);
  return new Date(
    fechaStr
      .replace(" a. m.", " AM")
      .replace(" p. m.", " PM")
      .replace(/\./g, "")
  );
}


/* ======================================================================
   ESTADO GLOBAL
   ====================================================================== */
let moduloActivo = null;
let datosActuales = [];

// ======================================================================
// ESTADO DE INFORMES (auditoría)
// ======================================================================
// Estados posibles por archivo:
// "pendiente"          → botón Revisar (gris)
// "en_revision"        → botón Continuar revisión (azul)
// "aprobado"           → botón Aprobado (verde)
// "rechazado"          → botón Pendiente por técnico (rojo)
let estadoInformes = {};

// Guardar estados en localStorage
function guardarEstados() {
  localStorage.setItem("estadoInformesAuditor", JSON.stringify(estadoInformes));
}

// Cargar estados al inicio
function cargarEstados() {
  const raw = localStorage.getItem("estadoInformesAuditor");
  if (raw) {
    try {
      estadoInformes = JSON.parse(raw);
    } catch (e) {
      estadoInformes = {};
    }
  }
}

/* ======================================================================
   1) INICIALIZACIÓN GENERAL
   ====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {

  if (!usuarioActual()) {
    await iniciarSesion();
  }

  prepararSidebar();

  // ✅ Cargar estados guardados del localStorage
  cargarEstados();      

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

// ✅ Re-activar eventos del encabezado (incluyendo sort)
prepararEventosTabla();

await cargarDatosModulo();
}

/* ======================================================================
   4) CREAR TABLA
   ====================================================================== */
function generarTablaHTML(modulo) {

  const ths = modulo.columnas
    .map(col => {
      if (col.id === "fecha") {
  return `
    <th style="cursor:pointer;">
      <span class="sortable" data-col="fecha" data-order="desc">
        ${col.label} <span class="flecha">🔽</span>
      </span>
    </th>`;
}
return `<th>${col.label}</th>`;
    })
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

  // ✅ Cargar archivos usando listarArchivosMCI (esto sí trae la fecha real del archivo)
const token = await obtenerToken();
datosActuales = await listarArchivosMCI(token);

// ✅ Debug para verificar fechas reales
console.log("=== FECHAS REALES RECIBIDAS ===");
datosActuales.forEach(x => {
  console.log(x.nombre, " → fechaReal:", x.fechaReal, " | fecha:", x.fecha);
});
console.log("================================");

// ✅ Ordenar por FECHA REAL — más reciente primero
datosActuales.sort((a, b) => {
  const fechaA = new Date(a.fechaReal);
  const fechaB = new Date(b.fechaReal);
  return fechaB - fechaA;
});

// ✅ Renderizar la tabla
renderTabla();

// ✅ Activar el ordenamiento después de que la tabla exista en el DOM
setTimeout(() => activarOrdenamientoFecha(), 0);
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

// ✅ FILTRADOS PERO MANTENIENDO EL ORDEN YA ORDENADO
const filtrados = datosActuales
  .filter(item =>
    item.archivo.nombre.endsWith(".xlsx") &&
    !item.archivo.nombre.includes("PreviewFotos")
  );

// ✅ IMPORTANTE: ya NO creamos una nueva lista, usamos filtrados EN EL ORDEN ACTUAL

filtrados.forEach((item) => {

  const idxReal = datosActuales.indexOf(item);

  const tds = moduloActivo.columnas
    .map(col => `<td>${item[col.id]}</td>`)
    .join("");

  const tr = document.createElement("tr");

 // Estado actual del informe (por ID de archivo)
const estado = estadoInformes[item.id] || "pendiente";

let botonHTML = "";

/* ---------------------------------------------------
   1) ESTADO: PENDIENTE  → Botón gris “Revisar”
--------------------------------------------------- */
if (estado === "pendiente") {
  botonHTML = `
    <button class="btn-estado btn-gris btn-revisar" data-idx="${idxReal}" data-id="${item.id}">
      <svg viewBox="0 0 24 24" width="16" height="16" fill="none"
           stroke="#324a78" stroke-width="2">
        <path d="M2 12s3.8-6 10-6 10 6 10 6-3.8 6-10 6-10-6-10-6Z"/>
        <circle cx="12" cy="12" r="3.2"/>
      </svg>
      Revisar
    </button>`;
}

/* ---------------------------------------------------
   2) ESTADO: EN REVISIÓN → Botón azul “Continuar revisión”
--------------------------------------------------- */
else if (estado === "en_revision") {
  botonHTML = `
    <button class="btn-estado btn-azul btn-revisar" data-idx="${idxReal}" data-id="${item.id}">
      ✏️ Continuar revisión
    </button>`;
}

/* ---------------------------------------------------
   3) ESTADO: APROBADO → Botón verde con check
--------------------------------------------------- */
else if (estado === "aprobado") {
  botonHTML = `
    <button class="btn-estado btn-verde" disabled>
      ✅ Aprobado
    </button>`;
}

/* ---------------------------------------------------
   4) ESTADO: RECHAZADO → Botón rojo “Pendiente por técnico”
--------------------------------------------------- */
else if (estado === "rechazado") {
  botonHTML = `
    <button class="btn-estado btn-rojo" disabled>
      ⚠️ Pendiente por técnico
    </button>`;
}

tr.innerHTML = `
  ${tds}
  <td style="text-align:center;">${botonHTML}</td>
`;

 tbody.appendChild(tr);
});

// ✅ activar ordenamiento después de pintar filas y con thead ya generado
activarOrdenamientoFecha();

// ✅ activar botones ver/aprobar
prepararEventosTabla();

} // ← esta es la llave final que cierra renderTabla()

// ✅ Activar ordenamiento por fecha al hacer clic en el encabezado
function activarOrdenamientoFecha() {
  const th = document.querySelector("span.sortable[data-col='fecha']");
  if (!th) return;

  th.onclick = () => {
    const estado = th.dataset.order || "desc";

    // Ordenar por fecha real
    datosActuales.sort((a, b) => {
      const fA = new Date(a.fechaReal);
      const fB = new Date(b.fechaReal);
      return estado === "desc" ? fA - fB : fB - fA;
    });

    // Alternar estado
    const nuevoEstado = estado === "desc" ? "asc" : "desc";
    th.dataset.order = nuevoEstado;

    // Actualizar la flecha en pantalla
    const flecha = th.querySelector(".flecha");
    flecha.textContent = nuevoEstado === "desc" ? "🔽" : "🔼";

    // Repintar tabla
    renderTabla();
  };
}
/* ======================================================================
   7) EVENTOS DE TABLA
   ====================================================================== */
function prepararEventosTabla() {

  // === EVENTO REVISAR ===
  document.querySelectorAll(".btn-revisar").forEach(btn => {
    btn.addEventListener("click", async () => {
      const idx = btn.dataset.idx;
      const item = datosActuales[idx];

      // ✅ marcar como en revisión
      estadoInformes[item.id] = "en_revision";
      guardarEstados();

      await verArchivo(item);
      renderTabla();
    });
  });

} // ✅ CIERRE CORRECTO DE prepararEventosTabla()

/* ======================================================================
   8) VER ARCHIVO — Excel REAL embebido desde SharePoint
   (Se elimina totalmente el preview hack)
   ====================================================================== */
async function verArchivo(item) {

  // Ocultar tabla y mostrar modal
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";
  window.__archivoActual = item;

  const visor = document.getElementById("visorIframe");
  visor.innerHTML = `
    <div style="padding:20px; text-align:center;">
      Cargando Excel Online…
    </div>
  `;

  try {
    const token = await obtenerToken();

    // Obtener metadata del archivo desde Graph
    const metaResp = await fetch(
      `https://graph.microsoft.com/v1.0${item.archivo.ruta}`,
      { headers: { "Authorization": `Bearer ${token}` } }
    );

    if (!metaResp.ok) {
      visor.innerHTML = `
        <div style="padding:20px; color:red; text-align:center;">
          Error cargando el archivo desde SharePoint.
        </div>
      `;
      return;
    }

    const meta = await metaResp.json();

    // Construir URL de embebido Excel Online
    const embedUrl =
  meta.webUrl.split("/_layouts/15/Doc.aspx")[0] +
  `/_layouts/15/Doc.aspx?sourcedoc=${meta.id}&action=view`;


    // Inyectar iframe
    visor.innerHTML = `
      <iframe
        src="${embedUrl}"
        style="width:100%; height:70vh; border:none;"
        allowfullscreen
      ></iframe>
    `;

  } catch (error) {
    console.error("Error en verArchivo:", error);
    visor.innerHTML = `
      <div style="padding:20px; color:red; text-align:center;">
        Error inesperado al abrir el archivo.
      </div>
    `;
  }
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
   10) EVENTOS DEL MODAL
   ====================================================================== */

document.getElementById("visorVolver").addEventListener("click", () => {
  const item = window.__archivoActual;

  // ✅ Si estaba revisando y no aprobó/rechazó → queda "en_revision"
  if (item && estadoInformes[item.id] === "en_revision") {
    // se mantiene el estado, no se toca
  }

  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  document.getElementById("visorIframe").innerHTML = "";

  // actualizar tabla después de cerrar
  renderTabla();
});

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

// ✅ Aprobar desde el visor (SIN mover archivo)
document.getElementById("visorAprobar").addEventListener("click", async () => {

  const item = window.__archivoActual;
  if (!item) return;

  // ✅ 1. Cambiar estado local
  estadoInformes[item.id] = "aprobado";
  guardarEstados();

  // ✅ 2. Registrar aprobación en Cloudflare KV usando fileIdReal
  await fetch("https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/" + item.archivo.fileIdReal, {
      method: "PUT"
  });

  // ✅ 3. Cerrar visor
  document.getElementById("visorVolver").click();

  // ✅ 4. Refrescar tabla (para mostrar ✅ Aprobado)
  renderTabla();
});
