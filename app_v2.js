/* ======================================================================
   0) CONFIGURACIÓN — FLUJO ÚNICO ONE DRIVE (EXCEL + JSON)
====================================================================== */

// ✅ ✅ URL del TRIGGER HTTP del flujo "Fotos Preview"
// (copiada directamente desde Power Automate, con & normales)
const FLOW_FOTOS =
  "https://defaulte4e1bc33e2834312bb3789010224b7.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e5f65d8cc4aa4001b6966552ed454170/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ybNuejYFtJf4p_P2vNPf_TY_Zzm2uvkSVYkqPu0GyQg";


async function tienePermisoOneDrive() {
  try {
    const token = await obtenerToken();
    if (!token) return false;

    // ✅ Intento mínimo: listar root (o una carpeta específica si prefieres)
    const resp = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=1",
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    // ✅ 200 = tiene acceso
    if (resp.status === 200) return true;

    // ❌ 403 / 401 = NO tiene permisos
    return false;

  } catch (err) {
    console.warn("Sin permisos OneDrive:", err);
    return false;
  }
}
/* ======================================================================
   1) IMPORTS — NECESARIOS
====================================================================== */
import {obtenerModulo } from "./modulos_v2.js";
import { obtenerToken, iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";

/* ======================================================================
   2) VARIABLES GLOBALES
====================================================================== */
window.moduloActivo = null;
window.datosActuales = [];
window.estadoInformes = {};
window.__archivoActual = null;
window.__mciIdActual = null;
window.__excelAbierto = false;

// ✅ Convierte correo en nombre legible
function nombreBonitoDesdeEmail, "");function nombreBonitoDesdeEmail(email) {

  // ✅ Construir nombre "Juanito Perez"
  return base
    .split(".")
    .map(function (p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    })
    .join(" ");
}
  if (!email || !email.includes("@")) {
    return "Desconocido";
  }

  // Parte antes del @
  let base = email.split("@")[0]; // ej: juanito.perez-ext

  // ✅ Eliminar sufijos tipo -ext, -etx, -external


/* =========================================================
   UTILIDAD — USUARIO LOGUEADO (AUDITOR)
========================================================= */
function obtenerUsuarioAuditor() {
  try {
    const user = usuarioActual();
    if (!user || !user.username) return null;

    // Extrae solo la parte antes del @
    return user.username.split("@")[0];
  } catch (e) {
    return null;
  }
}
/* ======================================================================
   3) GUARDAR / CARGAR ESTADO LOCAL
====================================================================== */
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
   4) INICIO DEL MÓDULO (MSAL SIN BLOQUEAR)
====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {
  const usuario = usuarioActual();

  // 🔒 SI NO HAY SESIÓN → FORZAR LOGIN Y NO CARGAR NADA
  if (!usuario) {
    await iniciarSesion().catch(() => {
      alert("Debes iniciar sesión con tu cuenta corporativa.");
    });

    if (!usuarioActual()) {
      // ❌ SIN SESIÓN → NO AVANZAMOS
      return;
    }
  }

  // ✅ SOLO CON SESIÓN VÁLIDA
  prepararSidebar();
  cargarEstados();
  seleccionarModulo("inicio");
});

/* ======================================================================
   5) SIDEBAR
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
   6) SELECCIONAR MÓDULO
====================================================================== */
async function seleccionarModulo(mod) {

  const cont = document.getElementById("contenedor-modulo");
  cont.innerHTML = "";

  if (mod === "inicio") {
    window.moduloActivo = null;
    cont.innerHTML = `
      <div style="padding:20px;">
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
   7) GENERAR TABLA HTML
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
        <tr><td colspan="${modulo.columnas.length + 1}" style="padding:20px; text-align:center;">
          Cargando…
        </td></tr>
      </tbody>
    </table>
  </div>`;
}

/* ======================================================================
   8) CARGAR DATOS DEL MÓDULO (SharePoint + KV)
====================================================================== */
async function cargarDatosModulo() {

  // 🔒 0) Si no hay sesión activa, no cargar nada
  if (!usuarioActual()) {
    return;
  }

  // 🔒 1) Verificar permisos reales en OneDrive
  try {
    const token = await obtenerToken();
    if (!token) throw new Error("Sin token");

    const check = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=1",
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    // ❌ Sin permisos → tabla vacía con leyenda estándar
    if (check.status !== 200) {
      document.getElementById("tbodyDatos").innerHTML = `
        <tr>
          <td colspan="99" style="padding:20px; text-align:center;">
            No hay informes pendientes.
          </td>
        </tr>
      `;
      return;
    }

  } catch (e) {
    // ❌ Error / sin permisos → misma salida limpia
    document.getElementById("tbodyDatos").innerHTML = `
      <tr>
        <td colspan="99" style="padding:20px; text-align:center;">
          No hay informes pendientes.
        </td>
      </tr>
    `;
    return;
  }

  // 🔒 2) Si el módulo no tiene pendientes, mostrar leyenda
  if (!window.moduloActivo?.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
      <tr>
        <td colspan="99" style="padding:20px; text-align:center;">
          No hay informes pendientes.
        </td>
      </tr>`;
    return;
  }

  // ✅ 3) Cargar datos desde KV (usuario AUTORIZADO)
  const tecnico = "usuario"; // o auditor logueado
  const respKV = await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
  );

  const listaKV = await respKV.json();

  // ✅ Mapear datos correctamente desde el KV
  window.datosActuales = listaKV.map(reg => {
    const fechaTexto = reg.fechaGenerado || "";

    return {
      // ✅ Columnas visibles
      nombre: reg.fileName,
      fecha: fechaTexto
        ? new Date(fechaTexto).toLocaleString("es-CO", {
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            hour: "2-digit",
            minute: "2-digit"
          })
        : "",
      tamano: reg.sizeBytes
        ? (reg.sizeBytes / 1024 / 1024).toFixed(2) + " MB"
        : "",

      // ✅ Datos internos
      fechaReal: fechaTexto,
      mciId: reg.mciId,
      estadoKV: reg.estado,
      fileIdentifierExcel: reg.fileIdentifierExcel,
      jsonFileId: reg.jsonFileId
    };
  });

  // ✅ 4) Ordenar por fecha descendente
  window.datosActuales.sort((a, b) => {
    const fa = Date.parse(b.fechaReal || "") || 0;
    const fb = Date.parse(a.fechaReal || "") || 0;
    return fa - fb;
  });

  renderTabla();
  setTimeout(() => activarOrdenamientoFecha(), 0);
}
/* ======================================================================
   9) RENDER TABLA
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
    item.nombre && item.nombre.endsWith(".xlsx") &&
    !item.nombre.includes("PreviewFotos")
  );

  filtrados.forEach(item => {
    const idx = window.datosActuales.indexOf(item);

    // ✅ Evitar undefined en columnas como fecha y tamaño
    const tds = window.moduloActivo.columnas
      .map(col => {
        const valor = item[col.id];
        return `<td>${valor ?? ""}</td>`;
      })
      .join("");

    const estado = item.estadoKV ?? "pendiente";

    const btn =
  estado === "pendiente"
    ? `<button class="btn-estado btn-gris btn-revisar" data-idx="${idx}">Revisar</button>`
  : estado === "en_revision"
    ? `<button class="btn-estado btn-azul btn-revisar" data-idx="${idx}">✏️ Continuar</button>`
  : estado === "aprobado"
  ? `<button class="btn-estado btn-verde btn-ver" data-idx="${idx}">✅ Aprobado</button>`
: estado === "rechazado"
  ? `<button class="btn-estado btn-rechazado btn-ver" data-idx="${idx}">⛔ Rechazado</button>`
  : `<button class="btn-estado btn-rojo" disabled>⚠️ Pendiente por técnico</button>`;

    const tr = document.createElement("tr");
    tr.innerHTML = `${tds}<td style="text-align:center;">${btn}</td>`;
    tbody.appendChild(tr);
  });

  prepararEventosTabla();
}

/* ======================================================================
   10) ORDENAR POR FECHA
====================================================================== */
function activarOrdenamientoFecha() {
  const th = document.querySelector("span.sortable[data-col='fecha']");
  if (!th) return;

  th.onclick = () => {
    const orden = th.dataset.order ?? "desc";

    window.datosActuales.sort((a, b) => {
      const FA = new Date(a.fechaReal);
      const FB = new Date(b.fechaReal);
      return orden === "desc" ? FA - FB : FB - FA;
    });

    th.dataset.order = orden === "desc" ? "asc" : "desc";
    th.querySelector(".flecha").textContent =
      orden === "desc" ? "🔽" : "🔼";

    renderTabla();
  };
}

/* ======================================================================
   11) EVENTOS DE TABLA
====================================================================== */
function prepararEventosTabla() {
  document.querySelectorAll(".btn-revisar, .btn-ver").forEach(btn => {
  btn.addEventListener("click", async () => {
    const idx = btn.dataset.idx;
    const item = window.datosActuales[idx];
    await verArchivo(item);
  });
});
}
/* ======================================================================
   12) BUSCAR JSON DE FOTOS EN ONEDRIVE
====================================================================== */
async function obtenerJsonFotos(item) {
  const resp = await fetch(FLOW_FOTOS, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileId: item.jsonFileId
    })
  });

  const data = await resp.json();

  // 🔑 imgsJson llega como STRING → lo convertimos a objeto
  if (typeof data.imgsJson === "string") {
    return JSON.parse(data.imgsJson);
  }

  return data.imgsJson;
}

/* =========================================================
   LECTOR SEGURO DE CELDAS EXCEL (SIN RENDER)
========================================================= */
function leerCeldaExcel(workbook, ref) {
  try {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const celda = sheet?.[ref];
    return celda ? String(celda.v).trim() : "—";
  } catch (e) {
    return "—";
  }
}

/* =========================================================
   LECTOR DE CELDA CON HOJA ESPECÍFICA
========================================================= */
function leerCeldaExcelHoja(workbook, sheetName, ref) {
  try {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return "—";

    const celda = sheet[ref];
    return celda ? String(celda.v).trim() : "—";
  } catch (e) {
    return "—";
  }
}

/* ======================================================================
   13) VER ARCHIVO — Vista previa del Excel + Fotos (IGUAL A VERSIÓN VIEJA)
====================================================================== */
async function verArchivo(item) {
  window.__archivoActual = item;
   // Evita reconstruir el visor más de una vez por apertura
let visorConstruido = false;

   // ==============================
// PASO 1 — Fuente única de verdad del informe
// (NO toca el DOM, NO renderiza)
// ==============================
const infoInforme = {
  tecnico: "Cargando datos…",
  celular: "Cargando datos…",
  depto: "Cargando datos…",
  beneficiario: "Cargando datos…",
  ot: "Cargando datos…",
  fecha: item.fecha ?? "—",
  lat: "Cargando datos…",
  lng: "Cargando datos…"
};
window.__infoInforme = infoInforme;
   // ==============================
// PASO 2 — Estado inicial (cargando datos)
// ==============================
infoInforme.tecnico = "Cargando datos…";
infoInforme.celular = "Cargando datos…";
infoInforme.depto = "Cargando datos…";
infoInforme.beneficiario = "Cargando datos…";
infoInforme.ot = "Cargando datos…";
infoInforme.fecha = item.fecha ?? "—"; // la fecha sí puede mostrarse


   // ==============================
// PASO 4 — Render único de la info del informe
// ==============================
function renderInfoInforme(info) {
  document.getElementById("infoTecnico").innerText = info.tecnico;
  document.getElementById("infoCelular").innerText = info.celular;
  document.getElementById("infoDepto").innerText = info.depto;
  document.getElementById("infoBeneficiario").innerText = info.beneficiario;
  document.getElementById("infoOT").innerText = info.ot;
  document.getElementById("infoFecha").innerText = info.fecha;

  // Coordenadas
  const coords =
    info.lat !== "No informado" && info.lng !== "No informado"
      ? `${info.lat}, ${info.lng}`
      : "No informado";

  document.getElementById("infoCoords").innerText = coords;
}

   // ✅ Renderizar información del informe UNA sola vez


   
   // ✅ Estado actual del informe
const estado = item.estadoKV || "pendiente";
   // 🔒 Resetear estado de apertura de Excel
window.__excelAbierto = false;
   


  // Ocultar tabla y mostrar modal
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

   // 🔄 Reset visual de aprobador / rechazador al abrir el modal
const spanAprobadoPor  = document.getElementById("infoAprobadoPor");
const spanRechazadoPor = document.getElementById("infoRechazadoPor");

if (spanAprobadoPor)  spanAprobadoPor.innerText  = "—";
if (spanRechazadoPor) spanRechazadoPor.innerText = "—";

   // 🔄 Estado inicial del botón Abrir Excel (mientras llega la URL)
const btnAbrirExcelUI = document.getElementById("visorAbrirExcel");

if (btnAbrirExcelUI) {
  btnAbrirExcelUI.disabled = true;
  btnAbrirExcelUI.innerText = "⏳ Cargando datos…";
  btnAbrirExcelUI.style.opacity = "0.6";
  btnAbrirExcelUI.style.cursor = "wait";
}

   // 🔒 Estado inicial de botones del modal (VISIBLE PERO DESHABILITADO)
const btnAprobarUI  = document.getElementById("visorAprobar");
const btnRechazarUI = document.getElementById("visorRechazar");

if (btnAprobarUI) {
  btnAprobarUI.disabled = true;
  btnAprobarUI.style.opacity = "0.4";
  btnAprobarUI.style.cursor = "not-allowed";
  btnAprobarUI.style.pointerEvents = "auto";
}

if (btnRechazarUI) {
  btnRechazarUI.disabled = true;
  btnRechazarUI.style.opacity = "0.4";
  btnRechazarUI.style.cursor = "not-allowed";
  btnRechazarUI.style.pointerEvents = "auto";
}

   // ✅ Crear contenedor de fotos (OBLIGATORIO)
const visor = document.getElementById("visorIframe");
visor.innerHTML = `
  <div style="
    border:1px solid #e5e7eb;
    border-radius:10px;
    padding:12px;
    background:#f9fafb;
    margin-bottom:16px;
  ">
    <div style="font-weight:700; margin-bottom:8px;">
      Datos del informe
    </div>

    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px; font-size:14px;">
      <div><strong>Técnico:</strong> <span id="infoTecnico">—</span></div>
      <div><strong>Celular:</strong> <span id="infoCelular">—</span></div>
      <div><strong>Departamento:</strong> <span id="infoDepto">—</span></div>
      <div><strong>ID Beneficiario:</strong> <span id="infoBeneficiario">—</span></div>
      <div><strong>IM / OT:</strong> <span id="infoOT">—</span></div>
      <div><strong>Fecha reporte:</strong> <span id="infoFecha">—</span></div>
      <div><strong>Coordenadas:</strong> <span id="infoCoords">—</span></div>
      <div><strong>Aprobado por:</strong> <span id="infoAprobadoPor">—</span></div>
<div><strong>Rechazado por:</strong> <span id="infoRechazadoPor">—</span></div>
    </div>
  </div>

  <h3 style="font-weight:800; margin-bottom:10px;">
    Fotos del informe
  </h3>
  <div id="visorFotos"></div>
`;

   renderInfoInforme(infoInforme);
   
// ==============================
// Fallback de datos base (SÍ llegan siempre)
// ==============================
   
   // ==============================
// PASO 1 — Modal dinámico por estado (CON RESET)
// ==============================
const btnDescargarUI = document.getElementById("visorDescargar");

// ✅ RESET SIEMPRE (MUY IMPORTANTE)
if (btnAprobarUI)   btnAprobarUI.style.display = "inline-block";
if (btnRechazarUI)  btnRechazarUI.style.display = "inline-block";
if (btnDescargarUI) btnDescargarUI.style.display = "inline-block";

// ✅ APLICAR REGLAS SEGÚN ESTADO
if (estado === "aprobado") {
  btnAprobarUI.style.display = "none";
  btnRechazarUI.style.display = "none";
  btnDescargarUI.style.display = "none";
}

if (estado === "rechazado") {
  btnAprobarUI.style.display = "none";
  btnRechazarUI.style.display = "none";
  btnDescargarUI.style.display = "none";
}

// ✅ Para pendientes / en revisión → se mantienen visibles
   
   // 🔒 Forzar Aprobar DESACTIVADO una vez el modal ya está visib

   console.log("ENVIANDO AL FLOW (Excel):", {
  fileIdentifierExcel: item.fileIdentifierExcel
});
  // === OBTENER EXCEL DESDE ONEDRIVE (FLOW DESCARGADOR) ===
  const resp = await fetch(FLOW_FOTOS, {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({
    fileIdentifierExcel: item.fileIdentifierExcel
  })
});


  if (!resp.ok) {
    throw new Error("No se pudo obtener el Excel desde OneDrive");
  }

  // === RESPUESTA DEL FLOW ===
const data = await resp.json();
console.log("RESPUESTA FLOW EXCEL:", data);
console.log("excelWebUrl recibido:", data.excelWebUrl);

  // ==============================
// PASO 3 — Sobrescribir infoInforme desde Excel (celdas reales)
// ==============================
if (data.excelBase64) {
  try {
    const wb = XLSX.read(data.excelBase64, { type: "base64" });

    // Técnico → M72
    const tecnicoExcel = leerCeldaExcel(wb, "M72");

    // N° de caso (IM / OT) → C9:F9
    const otExcel = [
      leerCeldaExcel(wb, "C9"),
      leerCeldaExcel(wb, "D9"),
      leerCeldaExcel(wb, "E9"),
      leerCeldaExcel(wb, "F9")
    ].filter(v => v !== "—").join(" ");

    // ID Beneficiario → B16:E16
    const beneficiarioExcel = [
      leerCeldaExcel(wb, "B16"),
      leerCeldaExcel(wb, "C16"),
      leerCeldaExcel(wb, "D16"),
      leerCeldaExcel(wb, "E16")
    ].filter(v => v !== "—").join(" ");

    // Departamento → B14:E14
    const deptoExcel = [
      leerCeldaExcel(wb, "B14"),
      leerCeldaExcel(wb, "C14"),
      leerCeldaExcel(wb, "D14"),
      leerCeldaExcel(wb, "E14")
    ].filter(v => v !== "—").join(" ");

    // Celular → M75:P75
    const celularExcel = [
      leerCeldaExcel(wb, "M75"),
      leerCeldaExcel(wb, "N75"),
      leerCeldaExcel(wb, "O75"),
      leerCeldaExcel(wb, "P75")
    ].filter(v => v !== "—").join(" ");

    if (tecnicoExcel)       infoInforme.tecnico = tecnicoExcel;
    if (otExcel)            infoInforme.ot = otExcel;
    if (beneficiarioExcel)  infoInforme.beneficiario = beneficiarioExcel;
    if (deptoExcel)         infoInforme.depto = deptoExcel;
    if (celularExcel)       infoInforme.celular = celularExcel;

     // ==============================
    // PASO 2 — Coordenadas geográficas
    // (hoja REG FOTOG PRUEBAS NECESARIAS)
    // ==============================
    const latExcel = leerCeldaExcelHoja(
      wb,
      "REG FOTOG PRUEBAS NECESARIAS",
      "F12"
    );

    const lngExcel = leerCeldaExcelHoja(
      wb,
      "REG FOTOG PRUEBAS NECESARIAS",
      "E12"
    );

    if (latExcel !== "—") infoInforme.lat = latExcel;
    if (lngExcel !== "—") infoInforme.lng = lngExcel;

  } catch (e) {
    console.warn("Error leyendo Excel:", e);
  }
}

// ✅ Guardar URL del Excel para abrir en línea
window.__archivoActual.excelWebUrl = data.excelWebUrl;

// ✅ Flow respondió → habilitar Abrir Excel

if (btnAbrirExcelUI && window.__archivoActual?.excelWebUrl) {
  btnAbrirExcelUI.disabled = false;
  btnAbrirExcelUI.innerText = "📊 Abrir Excel en línea";
  btnAbrirExcelUI.style.opacity = "1";
  btnAbrirExcelUI.style.cursor = "pointer";
}

   // ==============================
// NORMALIZAR CAMPOS NO DILIGENCIADOS
// ==============================
function normalizarCampo(valor) {
  if (!valor || valor === "Cargando datos…") {
    return "No informado";
  }
  return valor;
}

infoInforme.tecnico = normalizarCampo(infoInforme.tecnico);
infoInforme.celular = normalizarCampo(infoInforme.celular);
infoInforme.depto = normalizarCampo(infoInforme.depto);
infoInforme.beneficiario = normalizarCampo(infoInforme.beneficiario);
infoInforme.ot = normalizarCampo(infoInforme.ot);
   
// ==============================
// PASO 4 — Render único (YA con datos del Excel)
// ==============================
renderInfoInforme(infoInforme);

// ==============================
// CARGA DE FOTOS (NO TOCAR)
// ==============================
const jsonFotos = await obtenerJsonFotos(item);
item.fotosPreview = jsonFotos;

if (jsonFotos) {
  await renderizarFotos(item);
} else {
  document.getElementById("visorFotos").innerHTML =
    "<p style='color:#777;'>Este informe no tiene fotos adjuntas.</p>";
}

   }
/* ======================================================================
   14) RENDER FOTOS — ESTILO DOMINION
====================================================================== */
function renderizarFotos(item) {
  const fotos = item.fotosPreview;
  const galeria = document.getElementById("visorFotos");
  galeria.innerHTML = "";

  if (!fotos) {
    galeria.innerHTML = "<p style='color:#666;'>Sin fotos en preview.</p>";
    return;
  }

  // ✅ ORDEN ÚNICO DE FOTOS
  const orden = [
    { key: "gps", titulo: "GPS" },
    { key: "apInt", titulo: "AP Interior" },
    { key: "apExt1", titulo: "AP Exterior 1" },
    { key: "apExt2", titulo: "AP Exterior 2" },
    { key: "pcInt", titulo: "PC Interior" },
    { key: "movilExt", titulo: "Móvil Exterior" },
    { key: "senal", titulo: "Señalética" },
    { key: "med1", titulo: "Medición Eléctrica 1" }
  ];

  // ✅ UN SOLO GRID
  const grid = document.createElement("div");
  grid.className = "fotos-grid";

  orden.forEach(f => {
    const base64 = fotos[f.key];
    if (!base64) return;

    const card = document.createElement("div");
    card.className = "foto-card";
    card.innerHTML = `
      <img src="${base64}" alt="${f.titulo}">
    `;

    card.onclick = () => window.open(base64, "_blank");

    grid.appendChild(card);
  });

  galeria.appendChild(grid);
}
/* ======================================================================
   15) VOLVER
====================================================================== */
document.getElementById("visorVolver").addEventListener("click", () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  renderTabla();
});

/* ======================================================================
   16) APROBAR
====================================================================== */
document.getElementById("visorAprobar").addEventListener("click", async () => {

  // 🔒 Validación: Excel debe estar abierto
  if (!window.__excelAbierto) {
    alert("Debes abrir el Excel en línea antes de aprobar el informe.");
    return;
  }

  const item = window.__archivoActual;
  const mciId = item?.mciId || null;
  if (!mciId) return;

  // ✅ Obtener usuario y construir nombre legible
  const usuario = usuarioActual();
  const emailUsuario = usuario?.username || usuario?.email || "";
  const nombreUsuario = nombreBonitoDesdeEmail(emailUsuario);

  // ✅ Mostrar "Aprobado por" en el modal
  const spanAprobadoPor = document.getElementById("infoAprobadoPor");
  if (spanAprobadoPor) {
    spanAprobadoPor.innerText = nombreUsuario;
  }

  // ✅ Guardar metadata necesaria (EXISTENTE)
  const payloadMetadata = {
    departamento: window.__infoInforme.depto,
    ot: window.__infoInforme.ot,
    idBeneficiario: window.__infoInforme.beneficiario,
    lat: window.__infoInforme.lat,
    lng: window.__infoInforme.lng
  };

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/guardar-metadata/${mciId}`,
    {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payloadMetadata)
    }
  );

  // ✅ Aprobar informe
  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciId}`,
    { method: "PUT" }
  );

  // ✅ Cerrar modal y refrescar tabla
  await cargarDatosModulo();
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
});

/* ======================================================================
   17) RECHAZAR (OPCIONAL)
====================================================================== */
document.getElementById("visorRechazar").addEventListener("click", async () => {

  // 🔒 MISMA VALIDACIÓN QUE APROBAR
  if (!window.__excelAbierto) {
    alert("Debes abrir el Excel en línea antes de rechazar el informe.");
    return;
  }

  const item = window.__archivoActual;
  const mciId = item?.mciId ?? null;

  if (!mciId) {
    alert("No se pudo identificar el informe a rechazar.");
    return;
  }

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/rechazar/${mciId}`,
    { method: "PUT" }
  );

  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  await cargarDatosModulo();
});
/* =========================================================
   ABRIR EXCEL EN LÍNEA — HABILITA APROBAR Y RECHAZAR
========================================================= */
const btnAbrirExcel = document.getElementById("visorAbrirExcel");

if (btnAbrirExcel) {
  btnAbrirExcel.addEventListener("click", () => {
    const url = window.__archivoActual?.excelWebUrl;

    if (!url) {
      // Seguridad: no debería pasar
      alert("El Excel aún no está listo.");
      return;
    }

    // Abrir Excel
    window.open(url, "_blank");
    window.__excelAbierto = true;

    // ✅ Habilitar Aprobar
    const btnAprobar = document.getElementById("visorAprobar");
    if (btnAprobar) {
      btnAprobar.disabled = false;
      btnAprobar.style.opacity = "1";
      btnAprobar.style.cursor = "pointer";
      btnAprobar.style.pointerEvents = "auto";
    }

    // ✅ Habilitar Rechazar
    const btnRechazar = document.getElementById("visorRechazar");
    if (btnRechazar) {
      btnRechazar.disabled = false;
      btnRechazar.style.opacity = "1";
      btnRechazar.style.cursor = "pointer";
      btnRechazar.style.pointerEvents = "auto";
    }
  });
}

/* =========================================================
   ZOOM DINÁMICO REAL (FUNCIONA)
========================================================= */

document.addEventListener("mouseover", (e) => {
  const card = e.target.closest(".foto-card");
  if (!card) return;

  const img = card.querySelector("img");
  if (!img) return;

  card.addEventListener("mousemove", moverZoom);
  card.addEventListener("mouseleave", resetZoom);
});

function moverZoom(e) {
  const card = e.currentTarget;
  const img = card.querySelector("img");
  if (!img) return;

  const rect = card.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const y = e.clientY - rect.top;

  const px = (x / rect.width) * 100;
  const py = (y / rect.height) * 100;

  img.style.transformOrigin = `${px}% ${py}%`;
  img.style.transform = "scale(1.5)";
}

function resetZoom(e) {
  const card = e.currentTarget;
  const img = card.querySelector("img");
  if (!img) return;

  img.style.transform = "scale(1)";
  card.removeEventListener("mousemove", moverZoom);
  card.removeEventListener("mouseleave", resetZoom);
}
