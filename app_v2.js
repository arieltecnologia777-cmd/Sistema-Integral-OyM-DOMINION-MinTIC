/* ======================================================================
   APP.JS — Panel Auditor 
   Controlador principal del sistema 
   Versión Final — Ariel
   ====================================================================== */

import {
  listarArchivosMCI,
  descargarArchivo,
  formatearFecha,
  formatearTamano,
  obtenerModulo,
  MODULOS
} from './modulos_v2.js';

import { obtenerToken } from "./auth.js";

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
   INICIO
   ====================================================================== */

window.addEventListener("DOMContentLoaded", async () => {
  cargarEstados();
  prepararSidebar();
  seleccionarModulo("inicio");
});

/* ======================================================================
   SIDEBAR
   ====================================================================== */

function prepararSidebar() {
  document.querySelectorAll(".sb-item").forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll(".sb-item")
        .forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      seleccionarModulo(btn.dataset.mod);
    };
  });
}

/* ======================================================================
   SELECCIONAR MÓDULO
   ====================================================================== */

async function seleccionarModulo(mod) {
  const cont = document.getElementById("contenedor-modulo");
  cont.innerHTML = "";

  if (mod === "inicio") {
    moduloActivo = null;
    cont.innerHTML = `
      <div style="padding:20px">
        Bienvenido al <b>Panel Auditor</b><br>
      </div>`;
    return;
  }

  moduloActivo = obtenerModulo(mod);
  if (!moduloActivo) {
    cont.innerHTML = "<p>Error: módulo no configurado.</p>";
    return;
  }

  cont.innerHTML = generarTablaHTML(moduloActivo);
  await cargarDatosModulo();
}

/* ======================================================================
   CARGAR DATOS
   ====================================================================== */

async function cargarDatosModulo() {

  const token = await obtenerToken();

  const listaOD = await listarArchivosMCI(token); // Archivos OneDrive
  window.debugOD = listaOD;

  const tecnico = "usuario";
  const respKV = await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
  );
  const listaKV = await respKV.json();
  window.debugKV = listaKV;

  // ✅ Mezclar por fileName (la llave verdadera DEL SISTEMA)
  for (const a of listaOD) {

    const registro = listaKV.find(k => k.fileName === a.archivo.nombre);

    if (registro) {
      a.mciId = registro.mciId;
      a.estadoKV = registro.estado;
    } else {
      a.mciId = null;
      a.estadoKV = "pendiente";
    }
  }

  datosActuales = listaOD;
  renderTabla();
}

/* ======================================================================
   TABLA
   ====================================================================== */

function generarTablaHTML(modulo) {
  const ths = modulo.columnas.map(col => `<th>${col.label}</th>`).join("");

  return `
  <table class="tabla">
    <thead>
      <tr>${ths}<th>Acciones</th></tr>
    </thead>
    <tbody id="tbodyDatos">
      <tr><td colspan="99" style="text-align:center;padding:20px">Cargando…</td></tr>
    </tbody>
  </table>
  `;
}

function renderTabla() {
  const tbody = document.getElementById("tbodyDatos");
  tbody.innerHTML = "";

  const filtrados = datosActuales.filter(i =>
    i.archivo.nombre.endsWith(".xlsx") &&
    !i.archivo.nombre.includes("PreviewFotos")
  );

  filtrados.forEach(item => {
    const idx = datosActuales.indexOf(item);
    const estado = item.estadoKV || "pendiente";

    let boton = "";
    if (estado === "pendiente") {
      boton = `<button class="btn-estado btn-gris btn-revisar" data-idx="${idx}">Revisar</button>`;
    }
    else if (estado === "en_revision") {
      boton = `<button class="btn-estado btn-azul btn-revisar" data-idx="${idx}">✏️ Continuar revisión</button>`;
    }
    else if (estado === "aprobado") {
      boton = `<button class="btn-estado btn-verde" disabled>✅ Aprobado</button>`;
    }

    const tds = moduloActivo.columnas.map(col => `<td>${item[col.id]}</td>`).join("");

    tbody.innerHTML += `
      <tr>
        ${tds}
        <td>${boton}</td>
      </tr>
    `;
  });

  prepararEventosTabla();
}

/* ======================================================================
   EVENTOS DE TABLA
   ====================================================================== */

function prepararEventosTabla() {

  document.querySelectorAll(".btn-revisar").forEach(btn => {

    btn.onclick = async () => {
      const idx = btn.dataset.idx;
      const item = datosActuales[idx];
      await verArchivo(item);

      // ✅ Si entra a revisar → EN REVISIÓN
      estadoInformes[item.id] = "en_revision";
      guardarEstados();

      renderTabla();
    };
  });
}

/* ======================================================================
   VISOR PREVIEW (COMPLETO Y FUNCIONAL)
   ====================================================================== */

async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  window.__archivoActual = item;
  window.__mciIdActual = item.mciId;

  const token = await obtenerToken();

  // === Descargar Excel ===
  const url = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } });
  const blob = await resp.blob();
  const ab = await blob.arrayBuffer();

  const wb = XLSX.read(ab);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // === Generar Previews ===
  const rango1 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B9:P18" });
  const rango2 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B69:P69" });
  const rango3 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B71:M77" });

  let htmlPreview = `
    <h3>Información General</h3>
    ${rango1}
    <h3>Descripción</h3>
    ${rango2}
    <h3>Declaración</h3>
    ${rango3}
  `;

  // Fallback si Excel viene vacío
  if (!htmlPreview || htmlPreview.trim() === "") {
    htmlPreview = "<p style='color:#555;'>No se pudo generar vista previa del Excel.</p>";
  }

  // === Pintar el visor ===
  document.getElementById("visorIframe").innerHTML = `
    <div style="padding:20px">
      <h3>Vista previa del archivo</h3>
      <div style="background:white;padding:20px;border-radius:8px">
        ${htmlPreview}
      </div>
    </div>
  `;
}

/* ======================================================================
   APROBAR (USANDO mciId)
   ====================================================================== */

document.getElementById("visorAprobar").onclick = async () => {

  const mciId = window.__mciIdActual;
  const item = window.__archivoActual;

  if (!mciId) {
    alert("❌ No se encontró el mciId.");
    return;
  }

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciId}`,
    { method: "PUT" }
  );

  // ✅ Estado local
  estadoInformes[item.id] = "aprobado";
  guardarEstados();

  document.getElementById("visorVolver").click();
  renderTabla();
};

/* ======================================================================
   VOLVER DEL VISOR
   ====================================================================== */

document.getElementById("visorVolver").onclick = () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  document.getElementById("visorIframe").innerHTML = "";
};
