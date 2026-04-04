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
let estadoInformes = {};

function guardarEstados() {
  localStorage.setItem("estadoInformesAuditor", JSON.stringify(estadoInformes));
}

function cargarEstados() {
  const raw = localStorage.getItem("estadoInformesAuditor");
  if (raw) {
    try { estadoInformes = JSON.parse(raw); }
    catch { estadoInformes = {}; }
  }
}

/* ======================================================================
   INICIALIZACIÓN
   ====================================================================== */

window.addEventListener("DOMContentLoaded", async () => {
  if (!usuarioActual()) await iniciarSesion();
  prepararSidebar();
  cargarEstados();
  seleccionarModulo("inicio");
});

/* ======================================================================
   SIDEBAR
   ====================================================================== */

function prepararSidebar() {
  const botones = document.querySelectorAll(".sb-item");
  botones.forEach(btn => {
    btn.addEventListener("click", async () => {
      if (btn.classList.contains("logout")) {
        cerrarSesion(); return;
      }
      botones.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      seleccionarModulo(btn.dataset.mod);
    });
  });
}

/* ======================================================================
   CAMBIO DE MÓDULO
   ====================================================================== */

async function seleccionarModulo(mod) {
  const cont = document.getElementById("contenedor-modulo");
  cont.innerHTML = "";

  if (mod === "inicio") {
    moduloActivo = null;
    cont.innerHTML = `<div style="padding:20px;font-size:16px;">
      Bienvenido al <strong>Panel Auditor</strong>.</div>`;
    return;
  }

  moduloActivo = obtenerModulo(mod);
  if (!moduloActivo) {
    cont.innerHTML = "<p>Error: módulo desconocido.</p>";
    return;
  }

  cont.innerHTML = generarTablaHTML(moduloActivo);
  prepararEventosTabla();
  await cargarDatosModulo();
}

/* ======================================================================
   TABLA
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
        <tr><td colspan="99" style="padding:20px;text-align:center;">Cargando…</td></tr>
      </tbody>
    </table>
  </div>`;
}

/* ======================================================================
   CARGAR DATOS REALES DESDE ONEDRIVE
   ====================================================================== */

async function cargarDatosModulo() {
  if (!moduloActivo.pendientes) {
    document.getElementById("tbodyDatos").innerHTML =
      `<tr><td colspan="99" style="padding:20px;text-align:center;">
      No hay ruta configurada.</td></tr>`;
    return;
  }

  const token = await obtenerToken();
  datosActuales = await listarArchivosMCI(token);

  datosActuales.sort((a, b) => new Date(b.fechaReal) - new Date(a.fechaReal));

  renderTabla();
}

/* ======================================================================
   RENDER TABLA
   ====================================================================== */

function renderTabla() {
  const tbody = document.getElementById("tbodyDatos");
  tbody.innerHTML = "";

  if (!datosActuales.length) {
    tbody.innerHTML = `<tr><td colspan="99" style="padding:20px;text-align:center;">
      No hay informes.</td></tr>`;
    return;
  }

  datosActuales.forEach((item, idx) => {
    const estado = estadoInformes[item.id] || "pendiente";

    let boton = "";
    if (estado === "pendiente") {
      boton = `<button class="btn-revisar" data-idx="${idx}">Revisar</button>`;
    } else if (estado === "en_revision") {
      boton = `<button class="btn-revisar" data-idx="${idx}">Continuar revisión</button>`;
    } else if (estado === "aprobado") {
      boton = `<button disabled class="btn-verde">✅ Aprobado</button>`;
    } else {
      boton = `<button disabled class="btn-rojo">Rechazado</button>`;
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.nombre}</td>
      <td>${item.fecha}</td>
      <td>${item.tamano}</td>
      <td>${boton}</td>
    `;
    tbody.appendChild(tr);
  });
}

/* ======================================================================
   EVENTOS TABLA
   ====================================================================== */

function prepararEventosTabla() {
  document.querySelectorAll(".btn-revisar").forEach(btn => {
    btn.onclick = async () => {
      const idx = btn.dataset.idx;
      const item = datosActuales[idx];
      estadoInformes[item.id] = "en_revision";
      guardarEstados();
      await verArchivo(item);
      renderTabla();
    };
  });
}

/* ======================================================================
   VER ARCHIVO
   ====================================================================== */

async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  window.__archivoActual = item;

  const token = await obtenerToken();
  const urlDesc = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;

  const resp = await fetch(urlDesc, { headers: { "Authorization": `Bearer ${token}` }});
  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();

  // ... tu código de preview ...

}

/* ======================================================================
   APROBAR (MOVER ARCHIVO)
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
   BOTÓN APROBAR
   ====================================================================== */

document.getElementById("visorAprobar").onclick = async () => {
  const item = window.__archivoActual;
  if (!item) return;

  await aprobarArchivo(item);

  estadoInformes[item.id] = "aprobado";
  guardarEstados();

  document.getElementById("visorVolver").click();
};

/* ======================================================================
   BOTÓN VOLVER
   ====================================================================== */

document.getElementById("visorVolver").onclick = () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  renderTabla();
};
