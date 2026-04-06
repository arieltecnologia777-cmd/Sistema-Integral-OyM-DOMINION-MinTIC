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
    catch (e) { estadoInformes = {}; }
  }
}

/* ======================================================================
   1) INICIALIZACIÓN GENERAL
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
   3) SELECCIONAR MÓDULO
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
      </div>`;
    return;
  }

  moduloActivo = obtenerModulo(mod);

  if (!moduloActivo) {
    contenedor.innerHTML = "<p>Error: módulo desconocido.</p>";
    return;
  }

  contenedor.innerHTML = generarTablaHTML(moduloActivo);
  prepararEventosTabla();
  await cargarDatosModulo();
}

/* ======================================================================
   4) CARGAR DATOS DEL MÓDULO
   ====================================================================== */

async function cargarDatosModulo() {

  if (!moduloActivo.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
    <tr><td colspan="99" style="padding:20px; text-align:center;">
    No hay ruta configurada para este módulo.<br>
    (Ariel deberá especificarla cuando toque)
    </td></tr>`;
    return;
  }

  const token = await obtenerToken();

  // ✅ 1. Cargar archivos desde OneDrive
  const listaOD = await listarArchivosMCI(token);
  window.debugListaOD = listaOD;

  // ✅ 2. KV (pendientes, aprobados, etc.)
  const tecnico = "usuario";
  const respKV = await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
  );
  const listaKV = await respKV.json();

  // ✅ 3. Mezclar OneDrive + KV usando fileName
  for (const a of listaOD) {

    // ✅ COINCIDENCIA PERFECTA POR fileName
    const registro = listaKV.find(k => k.fileName === a.archivo.nombre);

    if (registro) {
      // Guardamos mciId en raíz (NO se pierde)
      a.mciId = registro.mciId;
      a.estadoKV = registro.estado;
    } else {
      a.mciId = null;
      a.estadoKV = "pendiente";
    }
  }

  datosActuales = listaOD;
  renderTabla();
  setTimeout(() => activarOrdenamientoFecha(), 0);
}

/* ======================================================================
   5) GENERAR TABLA
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
  <thead><tr>${ths}<th>Acciones</th></tr></thead>
  <tbody id="tbodyDatos">
  <tr><td colspan="${modulo.columnas.length + 1}" style="text-align:center; padding:20px;">Cargando…</td></tr>
  </tbody>
  </table>
  </div>`;
}

/* ======================================================================
   6) RENDER TABLA
   ====================================================================== */

function renderTabla() {
  const tbody = document.getElementById("tbodyDatos");

  if (!datosActuales || datosActuales.length === 0) {
    tbody.innerHTML = `
    <tr><td colspan="99" style="padding:20px; text-align:center;">
    No hay informes pendientes.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = "";

  const filtrados = datosActuales.filter(item =>
    item.archivo.nombre.endsWith(".xlsx") &&
    !item.archivo.nombre.includes("PreviewFotos")
  );

  filtrados.forEach(item => {
    const idxReal = datosActuales.indexOf(item);

    const tds = moduloActivo.columnas
      .map(col => `<td>${item[col.id]}</td>`)
      .join("");

    const tr = document.createElement("tr");

    const estado = item.estadoKV || "pendiente";
    let botonHTML = "";

    if (estado === "pendiente") {
      botonHTML = `
      <button class="btn-estado btn-gris btn-revisar" data-idx="${idxReal}">
      Revisar
      </button>`;
    }
    else if (estado === "en_revision") {
      botonHTML = `
      <button class="btn-estado btn-azul btn-revisar" data-idx="${idxReal}">
      ✏️ Continuar revisión
      </button>`;
    }
    else if (estado === "aprobado") {
      botonHTML = `<button class="btn-estado btn-verde" disabled>✅ Aprobado</button>`;
    }
    else if (estado === "rechazado") {
      botonHTML = `<button class="btn-estado btn-rojo" disabled>⚠️ Pendiente por técnico</button>`;
    }

    tr.innerHTML = `${tds}<td style="text-align:center;">${botonHTML}</td>`;
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
    const estado = th.dataset.order || "desc";

    datosActuales.sort((a, b) => {
      const fA = new Date(a.fechaReal);
      const fB = new Date(b.fechaReal);
      return estado === "desc" ? fA - fB : fB - fA;
    });

    th.dataset.order = estado === "desc" ? "asc" : "desc";
    th.querySelector(".flecha").textContent = estado === "desc" ? "🔽" : "🔼";

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
      const item = datosActuales[idx];
      await verArchivo(item);
      renderTabla();
    });
  });
}

/* ======================================================================
   9) VER ARCHIVO
   ====================================================================== */

async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  // ✅ GUARDAR EL ITEM ANTES DE QUE ONEDRIVE LO REEMPLACE
  window.__archivoActual = item;
  window.__mciIdActual = item.mciId ?? null;
   estadoInformes[item.id] = "en_revision";
guardarEstados();
   

  const token = await obtenerToken();

  // === DESCARGA EXCEL ===
  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, { headers: { "Authorization": `Bearer ${token}` } });
  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();

  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // (Preview omitido para ahorrar espacio… lo tuyo queda igual)

  // … tu preview …
}

/* ======================================================================
   10) APROBAR — VERSIÓN CORRECTA FINAL (USANDO mciId de KV)
   ====================================================================== */

document.getElementById("visorAprobar").addEventListener("click", async () => {

  const mciIdReal = window.__mciIdActual;

  if (!mciIdReal) {
    alert("❌ Error: No se encontró el mciId para este informe.");
    return;
  }

  await fetch(
    "https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/" + mciIdReal,
    { method: "PUT" }
  );

  document.getElementById("visorVolver").click();
  renderTabla();
});
