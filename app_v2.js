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
   btn.onclick = async () => {
     if (btn.classList.contains("logout")) return cerrarSesion();
     document.querySelectorAll(".sb-item").forEach(b => b.classList.remove("active"));
     btn.classList.add("active");
     seleccionarModulo(btn.dataset.mod);
   };
 });
}

/* ======================================================================
   3) SELECCIONAR MÓDULO
   ====================================================================== */

async function seleccionarModulo(mod) {
 const cont = document.getElementById("contenedor-modulo");
 cont.innerHTML = "";

 if (mod === "inicio") {
   moduloActivo = null;
   cont.innerHTML = `
     <div style="padding:20px; font-size:16px;">
       Bienvenido al <strong>Panel Auditor</strong>.
     </div>`;
   return;
 }

 moduloActivo = obtenerModulo(mod);
 if (!moduloActivo) {
   cont.innerHTML = "<p>Error: módulo desconocido.</p>";
   return;
 }

 cont.innerHTML = generarTablaHTML(moduloActivo);
 await cargarDatosModulo();
}

/* ======================================================================
   4) GENERAR TABLA
   ====================================================================== */

function generarTablaHTML(modulo) {
 const ths = modulo.columnas
   .map(col => `<th>${col.label}</th>`)
   .join("");

 return `
 <table class="tabla">
   <thead>
     <tr>${ths}<th>Acciones</th></tr>
   </thead>
   <tbody id="tbodyDatos">
     <tr><td colspan="99" style="padding:20px;text-align:center;">Cargando…</td></tr>
   </tbody>
 </table>`;
}

/* ======================================================================
   5) CARGAR DATOS
   ====================================================================== */

async function cargarDatosModulo() {
 const token = await obtenerToken();
 datosActuales = await listarArchivosMCI(token);

 // Mezcla con KV
 const tecnico = "usuario";
 const respKV = await fetch(
   `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
 );
 const listaKV = await respKV.json();

 for (const a of datosActuales) {
   const registro = listaKV.find(k => k.fileName === a.archivo.nombre);
   if (registro) {
     a.mciId = registro.mciId;
     a.estadoKV = registro.estado;
   } else {
     a.mciId = null;
     a.estadoKV = "pendiente";
   }
 }

 // Ordenar por fecha
 datosActuales.sort((a, b) => new Date(b.fechaReal) - new Date(a.fechaReal));

 renderTabla();
}

/* ======================================================================
   6) RENDER TABLA
   ====================================================================== */

function renderTabla() {
 const tbody = document.getElementById("tbodyDatos");
 tbody.innerHTML = "";

 const filtrados = datosActuales.filter(i =>
   i.archivo.nombre.endsWith(".xlsx") &&
   !i.archivo.nombre.includes("PreviewFotos")
 );

 filtrados.forEach(item => {
   const idx = datosActuales.indexOf(item);
   const estado = estadoInformes[item.id] || "pendiente";

   let boton = "";
   if (estado === "pendiente") {
     boton = `<button class="btn-estado btn-gris btn-revisar" data-idx="${idx}">Revisar</button>`;
   } else if (estado === "en_revision") {
     boton = `<button class="btn-estado btn-azul btn-revisar" data-idx="${idx}">✏️ Continuar revisión</button>`;
   } else if (estado === "aprobado") {
     boton = `<button class="btn-estado btn-verde" disabled>✅ Aprobado</button>`;
   }

   const tds = moduloActivo.columnas
     .map(col => `<td>${item[col
/* ======================================================================
   8) VER ARCHIVO — Vista previa del Excel + Fotos
   ====================================================================== */
async function verArchivo(item) {

  // ✅ MOSTRAR VISOR
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  // ✅ Guardar referencia del item y del mciId ANTES de que Graph destruya info
  window.__archivoActual = item;
  window.__mciIdActual  = item.mciId ?? null;

  const token = await obtenerToken();

  /* ============================================================
     1. DESCARGAR EXCEL REAL DESDE ONEDRIVE
     ============================================================ */
  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();

  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  /* ============================================================
     2. ELIMINAR PARTES PARA HACER PREVIEW
     ============================================================ */

  const eliminarFilas = (sheet, desde, hasta) => {
    for (let r = desde; r <= hasta; r++) {
      for (let c = 65; c <= 90; c++) {
        const celda = String.fromCharCode(c) + r;
        delete sheet[celda];
      }
    }
  };

  eliminarFilas(sheet, 19, 67);   // Quitar SAP / Equipos / Seriales
  for (let c = 66; c <= 80; c++) delete sheet[String.fromCharCode(c) + 10]; // Quitar título repetido

  /* ============================================================
     3. GENERAR LOS FRAGMENTOS HTML DEL EXCEL
     ============================================================ */

  const rango1 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B9:P18" });
  const rango2 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B69:P69" });
  const rango3 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B71:M77" });

  let htmlPreview = `
    <h3 style="font-weight:800;margin-bottom:10px;">Información General</h3>
    ${rango1}

    <h3 style="font-weight:800;margin-top:20px;margin-bottom:10px;">Descripción de la Falla / Hallazgos</h3>
    ${rango2}

    <h3 style="font-weight:800;margin-top:20px;margin-bottom:10px;">Declaración</h3>
    ${rango3}
  `;

  // ✅ Fallback si Excel viene vacío
  if (!htmlPreview || htmlPreview.trim() === "") {
    htmlPreview = `<p style="padding:20px;color:#444;">No se pudo generar vista previa del Excel.</p>`;
  }

  /* ============================================================
     4. OBTENER METADATOS (Excel Online Link)
     ============================================================ */

  const metaResp = await fetch(`https://graph.microsoft.com/v1.0${item.archivo.ruta}`, {
    headers: { "Authorization": `Bearer ${token}` }
  });
  const meta   = await metaResp.json();
  const webUrl = meta.webUrl;

  /* ============================================================
     5. PINTAR EL VISOR COMPLETO
     ============================================================ */

  const visor = document.getElementById("visorIframe");

  const cssEncabezados = `
    <style>
      h3 { background: transparent !important; }
      td { padding:4px 6px; }
      .encabezado-interno { background:#e6e6e6 !important; font-weight:700; padding:4px 6px; }
    </style>
  `;

  visor.innerHTML = `
    ${cssEncabezados}

    <div style="padding:20px;overflow:auto;">
      
      <div style="text-align:center;margin-bottom:20px;">
        <button id="btnExcelOnline" style="
          background:#0d6efd;color:white;border:none;
          padding:10px 20px;border-radius:8px;
          font-size:16px;cursor:pointer;font-weight:700;">
          🔵 Abrir versión completa en Excel Online
        </button>
      </div>

      <h3 style="font-weight:800;margin-bottom:10px;">Vista previa del archivo</h3>
      <div style="
        border:1px solid #dce3f5;
        background:white;border-radius:8px;
        padding:20px;margin-bottom:30px;">
        ${htmlPreview}
      </div>

      <h3 style="font-weight:800;margin-top:20px;">Fotos del informe (vista previa)</h3>
      <div id="galeriaPreview" style="
        margin-top:15px;
        display:grid;
        grid-template-columns:repeat(auto-fill,minmax(220px,1fr));
        gap:14px;">
      </div>

    </div>
  `;

  /* ============================================================
     6. APLICAR COLOREO DINÁMICO A ENCABEZADOS
     ============================================================ */

  setTimeout(() => {
    const patrones = [
      "N° DE CASO","Nº DE CASO","FECHA","CONTRATO","CONTRATISTA",
      "DEPARTAMENTO","MUNICIPIO","CENTRO POBLADO","SEDE INSTITUCIÓN EDUCATIVA",
      "CASO ESPECIAL","ID BENEFICIARIO","NOMBRE DEL RESPONSABLE",
      "NÚMERO DE CEDULA","NÚMERO DE CONTACTO","DESCRIPCIÓN DE LA FALLA",
      "DECLARACIÓN","DATOS DE QUIÉN ACOMPAÑA","DATOS DE QUIÉN REPARA",
      "NOMBRES Y APELLIDOS","CARGO","TELÉFONO","CELULAR",
      "CORREO ELECTRÓNICO","CORREO ELECTRONICO","FIRMA"
    ];

    visor.querySelectorAll("td").forEach(td => {
      const txt = td.innerText.toUpperCase().trim();
      if (patrones.some(p => txt.includes(p.toUpperCase()))) {
        td.style.backgroundColor = "#e6e6e6";
        td.style.fontWeight = "700";
      }
    });
  }, 80);

  /* ============================================================
     7. ABRIR EXCEL ONLINE
     ============================================================ */
  document.getElementById("btnExcelOnline").onclick = () => {
    window.open(webUrl, "_blank");
  };

  /* ============================================================
     8. FOTOS
     ============================================================ */

  const fotos   = item.fotosPreview;
  const galeria = document.getElementById("galeriaPreview");

  if (fotos) {
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

    orden.forEach(f => {
      const img = fotos[f.key];
      if (!img) return;

      const cont = document.createElement("div");
      cont.style.cssText = `
        border:1px solid #dce3f5;
        border-radius:10px;
        overflow:hidden;
        background:white;
        box-shadow:0 4px 12px rgba(0,0,0,0.1);
        cursor:pointer;
        display:flex;
        flex-direction:column;
      `;
      cont.innerHTML = `
        <div style="padding:6px 10px;font-weight:700;font-size:14px;border-bottom:1px solid #eee;">
          ${f.titulo}
        </div>
        <img src="${img}" style="width:100%;height:180px;object-fit:cover;">
      `;
      cont.onclick = () => window.open(img, "_blank");
      galeria.appendChild(cont);
    });

  } else {
    galeria.innerHTML = `<p style="color:#666;">Sin fotos en preview.</p>`;
  }
}

/* ======================================================================
   9) APROBAR (SIN MOVER ARCHIVO) — USANDO mciId CORRECTO
   ====================================================================== */
document.getElementById("visorAprobar").onclick = async () => {
  const item = window.__archivoActual;
  const mciId = window.__mciIdActual;

  if (!mciId) {
    alert("❌ No se encontró el mciId para este informe.");
    return;
  }

  // ✅ Marcar como aprobado localmente
  estadoInformes[item.id] = "aprobado";
  guardarEstados();

  // ✅ Registrar en KV
  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciId}`,
    { method: "PUT" }
  );

  // ✅ Cerrar visor y refrescar tabla
  document.getElementById("visorVolver").click();
  renderTabla();
};

/* ======================================================================
   10) BOTÓN VOLVER
   ====================================================================== */
document.getElementById("visorVolver").onclick = () => {
  document.getElementById("modalVisor").style.display  = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  document.getElementById("visorIframe").innerHTML = "";
};
