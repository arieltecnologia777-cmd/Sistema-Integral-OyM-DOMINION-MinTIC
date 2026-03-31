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

import { MODULOS, obtenerModulo } from "./modulos.js";
import { iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";
import { cargarDesdeCarpeta, obtenerURLTemporal, moverArchivo } from "./graph.js";

/* ======================================================================
   ESTADO GLOBAL
   ====================================================================== */
let moduloActivo = null;
let datosActuales = []; // contenido actual de la tabla


/* ======================================================================
   1) INICIALIZACIÓN GENERAL
   ====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {

  // ✅ Verificamos sesión
  if (!usuarioActual()) {
    await iniciarSesion();
  }

  // ✅ Configuramos navegación del sidebar
  prepararSidebar();

  // ✅ Iniciar en módulo Inicio (placeholder)
  seleccionarModulo("inicio");
});

/* ======================================================================
   2) CONFIGURAR SIDEBAR
   ====================================================================== */
function prepararSidebar() {
  const botones = document.querySelectorAll(".sb-item");

  botones.forEach(btn => {
    btn.addEventListener("click", async () => {

      // Cerrar sesión
      if (btn.classList.contains("logout")) {
        cerrarSesion();
        return;
      }

      // Apagar estados previos
      botones.forEach(b => b.classList.remove("active"));

      // Activar botón
      btn.classList.add("active");

      // Detectar módulo
      const mod = btn.dataset.mod;
      seleccionarModulo(mod);
    });
  });
}

/* ======================================================================
   3) CAMBIAR DE MÓDULO (Inicio / MCI / MPR)
   ====================================================================== */
async function seleccionarModulo(mod) {

  const contenedor = document.getElementById("contenedor-modulo");
  contenedor.innerHTML = ""; // limpiar pantalla

  /* -----------------------------------------
     MODULO INICIO (Pantalla simple)
     ----------------------------------------- */
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

  /* -----------------------------------------
     MODULOS MCI / MPR
     ----------------------------------------- */
  moduloActivo = obtenerModulo(mod);

  if (!moduloActivo) {
    contenedor.innerHTML = "<p>Error: módulo desconocido.</p>";
    return;
  }

  // Dibujar tabla vacía
  contenedor.innerHTML = generarTablaHTML(moduloActivo);

  // Cargar datos del módulo
  await cargarDatosModulo();
}

/* ======================================================================
   4) CREAR TABLA (HTML dinámico por módulo)
   ====================================================================== */
function generarTablaHTML(modulo) {

  // Columnas de modulos.js
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
          <tr><td colspan="${modulo.columnas.length + 1}" style="text-align:center; padding:20px;">Cargando…</td></tr>
        </tbody>
      </table>
    </div>
  `;
}

/* ======================================================================
   5) CARGAR DATOS DESDE ONE DRIVE
   ====================================================================== */
async function cargarDatosModulo() {

  if (!moduloActivo.pendientes) {
    console.warn("⚠️ Aún no se ha configurado la carpeta de pendientes.");
    document.getElementById("tbodyDatos").innerHTML = `
      <tr><td colspan="99" style="padding:20px; text-align:center;">
        No hay ruta configurada para este módulo.<br>
        (Ariel deberá especificarla cuando toque)
      </td></tr>
    `;
    return;
  }

  // Llamamos graph.js → carga normalizada
  datosActuales = await cargarDesdeCarpeta(moduloActivo, false);

  renderTabla();
}

/* ======================================================================
   6) RENDER DE TABLA (llenar <tbody>)
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

  // Botón VER
  document.querySelectorAll(".btn-ver").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await verArchivo(item);
    });
  });

  // Botón APROBAR
  document.querySelectorAll(".btn-aprobar").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await aprobarArchivo(item);
    });
  });
}

/* ======================================================================
   8) VER ARCHIVO (Preview mediante URL temporal)
   ====================================================================== */
async function verArchivo(item) {

  if (!item?.archivo?.ruta) {
    alert("No se pudo localizar el archivo.");
    return;
  }

  const urlTemp = await obtenerURLTemporal(item.archivo.ruta);

  if (!urlTemp) {
    alert("No se pudo obtener vista previa.");
    return;
  }

  // Abrir en nueva pestaña
  window.open(urlTemp, "_blank");
}

/* ======================================================================
   9) APROBAR (mover archivo OneDrive)
   ====================================================================== */
async function aprobarArchivo(item) {

  if (!moduloActivo.aprobados) {
    alert("No hay carpeta de aprobados configurada.");
    return;
  }

  const r1 = item.archivo.ruta; // ruta actual
  const r2 = `${moduloActivo.aprobados}/${item.archivo.nombre}`;

  const ok = await moverArchivo(r1, r2);

  if (!ok) {
    alert("Error moviendo archivo.");
    return;
  }

  alert(`✅ Informe aprobado: ${item.archivo.nombre}`);

  // Recargar datos
  await cargarDatosModulo();
}
