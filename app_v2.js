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
        return `<th class="sortable" data-sort="fecha" style="cursor:pointer;">
                  ${col.label} 🔽
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

  datosActuales = await cargarDesdeCarpeta(moduloActivo, false);
  console.log("=== FECHAS REALES RECIBIDAS ===");
datosActuales.forEach(x => {
  console.log(x.nombre, " → fechaReal:", x.fechaReal, " | fecha:", x.fecha);
  console.log("Objeto completo x:", x);   // ✅ AÑÁDE ESTO
});
console.log("================================");

// ✅ ORDENAR POR FECHA REAL — MÁS RECIENTE PRIMERO
datosActuales.sort((a, b) => {
  const fechaA = new Date(a.fechaReal);
  const fechaB = new Date(b.fechaReal);
  return fechaB - fechaA;
});

// ✅ Renderizar tabla
renderTabla();

// ✅ Activar ordenamiento una vez el DOM esté construido
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

  tr.innerHTML = `
    ${tds}
    <td>
      <button class="btn-ver" data-idx="${idxReal}" style="margin-right:6px;">Ver</button>
      <button class="btn-aprobar" data-idx="${idxReal}">Aprobar</button>
    </td>
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
  const th = document.querySelector("th.sortable");
  if (!th) return;

  th.onclick = () => {
    const dir = th.dataset.order || "desc";


    th.dataset.order = dir === "desc" ? "asc" : "desc";

    renderTabla();
  };
}

/* ======================================================================
   7) EVENTOS DE TABLA
   ====================================================================== */
function prepararEventosTabla() {

  // === EVENTO VER ===
  document.querySelectorAll(".btn-ver").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await verArchivo(item);
    });
  });

  // === EVENTO APROBAR ===
  document.querySelectorAll(".btn-aprobar").forEach(btn => {
    btn.addEventListener("click", async () => {
      const item = datosActuales[btn.dataset.idx];
      await aprobarArchivo(item);
    });
  });

  // ✅ ORDENAR POR FECHA AL CLIC
  document.querySelectorAll("th.sortable").forEach(th => {
    th.addEventListener("click", () => {

      const direccionActual = th.dataset.order || "desc";

      datosActuales.sort((a, b) => {
  const fA = parseFechaCol(a.fecha);
  const fB = parseFechaCol(b.fecha);
  return direccionActual === "desc"
    ? fA - fB   // ascendente
    : fB - fA;  // descendente
});

      // Alternar dirección
      th.dataset.order = direccionActual === "desc" ? "asc" : "desc";

      // Volver a pintar la tabla
      renderTabla();
    });
  });

}

/* ======================================================================
   8) VER ARCHIVO — Vista previa del Excel + Fotos
   ====================================================================== */
async function verArchivo(item) {

  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";
  window.__archivoActual = item;

  const token = await obtenerToken();

  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, {
    headers: { "Authorization": `Bearer ${token}` }
  });
  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();
  const wb = XLSX.read(arrayBuffer);

  const sheet = wb.Sheets[wb.SheetNames[0]];

  // === ELIMINAR SAP, EQUIPOS, SERIALES ===
  const eliminarFilas = (sheet, desde, hasta) => {
    for (let r = desde; r <= hasta; r++) {
      for (let c = 65; c <= 90; c++) {
        const celda = String.fromCharCode(c) + r;
        delete sheet[celda];
      }
    }
  };

  eliminarFilas(sheet, 19, 67);

  // === OCULTAR TÍTULO DUPLICADO DE SECCIÓN 1 (FILA 10) ===
  for (let c = 66; c <= 80; c++) {
    delete sheet[String.fromCharCode(c) + 10];
  }

  const rango1 = XLSX.utils.sheet_to_html({
    ...sheet,
    '!ref': "B9:P18"
  });

  const rango2 = XLSX.utils.sheet_to_html({
    ...sheet,
    '!ref': "B69:P69"
  });

  const rango3 = XLSX.utils.sheet_to_html({
    ...sheet,
    '!ref': "B71:M77"
  });

  const htmlPreview = `
  <h3 style="font-weight:800; margin-bottom:8px;">Información General</h3>
  ${rango1}

  <h3 style="font-weight:800; margin-top:20px; margin-bottom:8px;">Descripción de la falla / hallazgos</h3>
  ${rango2}

  <h3 style="font-weight:800; margin-top:20px; margin-bottom:8px;">Declaración</h3>
  ${rango3}
`;
   
  const metaResp = await fetch(
    `https://graph.microsoft.com/v1.0${item.archivo.ruta}`,
    { headers: { "Authorization": `Bearer ${token}` } }
  );
  const meta = await metaResp.json();
  const webUrl = meta.webUrl;

  const visor = document.getElementById("visorIframe");

const cssEncabezados = `
  <style>

    /* NO aplicar gris a los títulos principales */
    h3 { background: transparent !important; }

    /* Encabezados internos: texto totalmente en MAYÚSCULAS */
    td {
      padding: 4px 6px;
    }

    td > * {
      display: inline-block;
    }

    /* Fondo gris SOLO a títulos detectados por la clase */
    .encabezado-interno {
      background-color: #e6e6e6 !important;
      font-weight: 700 !important;
      padding: 4px 6px;
      display: inline-block;
    }

  </style>
`;

  visor.innerHTML = `
    ${cssEncabezados}
    <div style="padding:20px; overflow:auto;">

      <div style="text-align:center; margin-bottom:20px;">
        <button id="btnExcelOnline" style="
          background:#0d6efd;
          color:white;
          border:none;
          padding:10px 20px;
          border-radius:8px;
          font-size:16px;
          cursor:pointer;
          font-weight:700;">
          🔵 Abrir versión completa en Excel Online
        </button>
      </div>

      <h3 style="font-weight:800; margin-bottom:10px;">Vista previa del archivo</h3>

<div style="
  border:1px solid #dce3f5;
  background:white;
  border-radius:8px;
  padding:20px;
  margin-bottom:30px;">
  ${htmlPreview}
</div>


<h3 style="font-weight:800; margin-top:20px;">Fotos del informe (vista previa)</h3>

      <div id="galeriaPreview" style="
        margin-top:15px;
        display:grid;
        grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
        gap:14px;">
      </div>

    </div>
  `;
   
// ✅ Pintar encabezados internos específicos en gris (versión robusta)
setTimeout(() => {

  const patrones = [
    "N° DE CASO",
    "Nº DE CASO",
    "FECHA",
    "CONTRATO",
    "CONTRATISTA",
    "DEPARTAMENTO",
    "MUNICIPIO",
    "CENTRO POBLADO",
    "SEDE INSTITUCIÓN EDUCATIVA",
    "CASO ESPECIAL",
    "ID BENEFICIARIO",
    "NOMBRE DEL RESPONSABLE",
    "NÚMERO DE CEDULA",
    "NÚMERO DE CONTACTO",
    "DESCRIPCIÓN DE LA FALLA",
    "DECLARACIÓN",
    "DATOS DE QUIÉN ACOMPAÑA",
    "DATOS DE QUIÉN REPARA",
    "NOMBRES Y APELLIDOS",
    "CARGO",
    "TELÉFONO",
    "CELULAR",
    "CORREO ELECTRÓNICO",
    "CORREO ELECTRONICO",   // ✅ ESTA ES LA NUEVA
    "FIRMA"
  ];


  const celdas = visor.querySelectorAll("td");

  celdas.forEach(td => {
    const texto = td.innerText.toUpperCase().trim();

    const coincide = patrones.some(p => texto.includes(p.toUpperCase()));

    if (coincide) {
      td.style.backgroundColor = "#e6e6e6";
      td.style.fontWeight = "700";
    }
  });

}, 80);

  document.getElementById("btnExcelOnline").onclick = () => {
    window.open(webUrl, "_blank");
  };

  // === FOTOS ===
  const fotos = item.fotosPreview;
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
      const base64 = fotos[f.key];
      if (!base64) return;

      const cont = document.createElement("div");
      cont.style.border = "1px solid #dce3f5";
      cont.style.borderRadius = "10px";
      cont.style.overflow = "hidden";
      cont.style.background = "#fff";
      cont.style.boxShadow = "0 4px 12px rgba(0,0,0,0.1)";
      cont.style.cursor = "pointer";
      cont.style.display = "flex";
      cont.style.flexDirection = "column";

      cont.innerHTML = `
        <div style="padding:6px 10px; font-weight:700; font-size:14px; border-bottom:1px solid #eee;">
          ${f.titulo}
        </div>
        <img src="${base64}" style="width:100%; height:180px; object-fit:cover;">
      `;

      cont.onclick = () => window.open(base64, "_blank");

      galeria.appendChild(cont);
    });

  } else {
    galeria.innerHTML = "<p style='color:#666;'>Sin fotos en preview.</p>";
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
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  document.getElementById("visorIframe").innerHTML = "";
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

document.getElementById("visorAprobar").addEventListener("click", async () => {
  const item = window.__archivoActual;
  if (!item) return;

  await aprobarArchivo(item);
  document.getElementById("visorVolver").click();
});

document.getElementById("visorRechazar").addEventListener("click", () => {
  alert("Función de rechazo pendiente.");
});
