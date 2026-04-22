/* ======================================================================
   0) CONFIGURACIÓN — FLUJO ÚNICO ONE DRIVE (EXCEL + JSON)
====================================================================== */

// ✅ ✅ URL del TRIGGER HTTP del flujo "Fotos Preview"
// (copiada directamente desde Power Automate, con & normales)
const FLOW_FOTOS =
  "https://defaulte4e1bc33e2834312bb3789010224b7.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e5f65d8cc4aa4001b6966552ed454170/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ybNuejYFtJf4p_P2vNPf_TY_Zzm2uvkSVYkqPu0GyQg";

/* ======================================================================
   0) IMPORTS — NECESARIOS
====================================================================== */
import {obtenerModulo } from "./modulos_v2.js";
import { obtenerToken, iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";

/* ======================================================================
   1) VARIABLES GLOBALES
====================================================================== */
window.moduloActivo = null;
window.datosActuales = [];
window.estadoInformes = {};
window.__archivoActual = null;
window.__mciIdActual = null;

/* ======================================================================
   2) GUARDAR / CARGAR ESTADO LOCAL
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
   3) INICIO DEL MÓDULO (MSAL SIN BLOQUEAR)
====================================================================== */
window.addEventListener("DOMContentLoaded", async () => {

  // NO BLOQUEAR NUNCA EL MÓDULO
  if (!usuarioActual()) iniciarSesion().catch(() => console.warn("MSAL pendiente…"));

  prepararSidebar();
  cargarEstados();
  seleccionarModulo("inicio");
});

/* ======================================================================
   4) SIDEBAR
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
   5) SELECCIONAR MÓDULO
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
   6) GENERAR TABLA HTML
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
   7) CARGAR DATOS DEL MÓDULO (SharePoint + KV)
====================================================================== */
async function cargarDatosModulo() {

  if (!window.moduloActivo?.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
      <tr><td colspan="99" style="padding:20px; text-align:center;">
        No hay informes pendientes.
      </td></tr>`;
    return;
  }

  const tecnico = "usuario"; // o el auditor logueado
  const respKV = await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`
  );

  const listaKV = await respKV.json();

  // ✅ Mapear datos correctamente desde el KV
  window.datosActuales = listaKV.map(reg => {

    const fechaTexto = reg.fechaGenerado || ""; // viene ISO desde el worker

    return {
      // ✅ Columnas visibles en la tabla
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
      fechaReal: fechaTexto, // para ordenamiento
      mciId: reg.mciId,
      estadoKV: reg.estado,
      fileIdentifierExcel: reg.fileIdentifierExcel,
      jsonFileId: reg.jsonFileId
    };
  });

  // ✅ ORDENAR POR DEFECTO: MÁS RECIENTE ARRIBA
  window.datosActuales.sort((a, b) => {
    const fa = Date.parse(b.fechaReal || "") || 0;
    const fb = Date.parse(a.fechaReal || "") || 0;
    return fa - fb;
  });

  renderTabla();
  setTimeout(() => activarOrdenamientoFecha(), 0);
}

/* ======================================================================
   8) RENDER TABLA
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
      estado === "pendiente" ? `<button class="btn-estado btn-gris btn-revisar" data-idx="${idx}">Revisar</button>` :
      estado === "en_revision" ? `<button class="btn-estado btn-azul btn-revisar" data-idx="${idx}">✏️ Continuar</button>` :
      estado === "aprobado" ? `<button class="btn-estado btn-verde" disabled>✅ Aprobado</button>` :
      `<button class="btn-estado btn-rojo" disabled>⚠️ Pendiente por técnico</button>`;

    const tr = document.createElement("tr");
    tr.innerHTML = `${tds}<td style="text-align:center;">${btn}</td>`;
    tbody.appendChild(tr);
  });

  prepararEventosTabla();
}

/* ======================================================================
   9) ORDENAR POR FECHA
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
   10) EVENTOS DE TABLA
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
   11) BUSCAR JSON DE FOTOS EN ONEDRIVE
====================================================================== */
async function obtenerJsonFotos(item) {
  console.log("DEBUG: entrar a obtenerJsonFotos", item);

  const resp = await fetch(FLOW_FOTOS, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileId: item.jsonFileId
    })
  });

  const fotos = await resp.json(); // ✅ JSON puro que viene del flujo
  return fotos;
}
/* ======================================================================
   12) VER ARCHIVO — Vista previa del Excel + Fotos (IGUAL A VERSIÓN VIEJA)
====================================================================== */
async function verArchivo(item) {
  window.__archivoActual = item;

  // Ocultar tabla y mostrar modal
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  // === OBTENER EXCEL DESDE ONEDRIVE (FLOW DESCARGADOR) ===
  const resp = await fetch(FLOW_FOTOS, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      tipo: "excel",
      fileId: item.fileIdentifierExcel
    })
  });

  if (!resp.ok) {
    throw new Error("No se pudo obtener el Excel desde OneDrive");
  }

  const blob = await resp.blob();
  const arrayBuffer = await blob.arrayBuffer();

  // === LEER EXCEL ===
  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // === ELIMINAR SAP, EQUIPOS, SERIALES (IGUAL QUE ANTES) ===
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

  // === CONVERTIR RANGOS A HTML ===
  const rango1 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B9:P18" });
  const rango2 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B69:P69" });
  const rango3 = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B71:M77" });

  const htmlPreview = `
    <h3 style="font-weight:800; margin-bottom:8px;">Información General</h3>
    ${rango1}

    <h3 style="font-weight:800; margin-top:20px; margin-bottom:8px;">
      Descripción de la falla / hallazgos
    </h3>
    ${rango2}

    <h3 style="font-weight:800; margin-top:20px; margin-bottom:8px;">
      Declaración
    </h3>
    ${rango3}
  `;

  const visor = document.getElementById("visorIframe");

  visor.innerHTML = `
    <div style="padding:20px; overflow:auto;">

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
      <div id="visorFotos" style="
  margin-top:15px;
  display:grid;
  grid-template-columns: repeat(3, 1fr);
  gap:14px;">
</div>

    </div>
  `;

  // ✅ Pintar encabezados internos en gris (versión vieja)
  setTimeout(() => {
    const patrones = [
      "N° DE CASO","Nº DE CASO","FECHA","CONTRATO","CONTRATISTA",
      "DEPARTAMENTO","MUNICIPIO","CENTRO POBLADO",
      "SEDE INSTITUCIÓN EDUCATIVA","CASO ESPECIAL","ID BENEFICIARIO",
      "NOMBRE DEL RESPONSABLE","NÚMERO DE CEDULA","NÚMERO DE CONTACTO",
      "DESCRIPCIÓN DE LA FALLA","DECLARACIÓN",
      "DATOS DE QUIÉN ACOMPAÑA","DATOS DE QUIÉN REPARA",
      "NOMBRES Y APELLIDOS","CARGO","TELÉFONO","CELULAR",
      "CORREO ELECTRÓNICO","CORREO ELECTRONICO","FIRMA"
    ];

    const celdas = visor.querySelectorAll("td");
    celdas.forEach(td => {
      const texto = td.innerText.toUpperCase().trim();
      if (patrones.some(p => texto.includes(p))) {
        td.style.backgroundColor = "#e6e6e6";
        td.style.fontWeight = "700";
      }
    });
  }, 80);

  // === CARGA DE FOTOS (NO TOCADO) ===
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
   13) RENDER FOTOS — ESTILO DOMINION
====================================================================== */
function renderizarFotos(item) {
  const fotos = item.fotosPreview;
  const galeria = document.getElementById("visorFotos");

  galeria.innerHTML = "";

  if (!fotos) {
    galeria.innerHTML = "<p style='color:#666;'>Sin fotos en preview.</p>";
    return;
  }

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
}
/* ======================================================================
   14) VOLVER
====================================================================== */
document.getElementById("visorVolver").addEventListener("click", () => {
  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
  renderTabla();
});

/* ======================================================================
   15) APROBAR
====================================================================== */
document.getElementById("visorAprobar").addEventListener("click", async () => {

  const item = window.__archivoActual;
  const mciId = item?.mciId || null;
  if (!mciId) return;

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/aprobar/${mciId}`,
    { method: "PUT" }
  );

  await cargarDatosModulo();

  document.getElementById("modalVisor").style.display = "none";
  document.getElementById("contenedor-modulo").style.display = "block";
});

/* ======================================================================
   16) RECHAZAR (OPCIONAL)
====================================================================== */
document.getElementById("visorRechazar").addEventListener("click", async () => {

  const item = window.__archivoActual;
  const mciId = item?.mciId || null;

  if (!mciId) return;

  await fetch(
    `https://cloudflare-index.modulo-de-exclusiones.workers.dev/rechazar/${mciId}`,
    { method: "PUT" }
  );
});
/* ======================================================================
 17 ABRIR EXCEL EN LÍNEA — PASO SEGURO (NO TOCA PREVIEW)
====================================================================== */
const btnAbrirExcel = document.getElementById("visorAbrirExcel");

if (btnAbrirExcel) {
  btnAbrirExcel.addEventListener("click", () => {
    const item = window.__archivoActual;

    if (!item) {
      alert("No hay archivo cargado.");
      return;
    }

    // ✅ Solo abrir si el valor YA es una URL válida
    if (
      item.fileIdentifierExcel &&
      typeof item.fileIdentifierExcel === "string" &&
      item.fileIdentifierExcel.startsWith("http")
    ) {
      window.open(item.fileIdentifierExcel, "_blank");
    } else {
      alert(
        "Este informe no tiene un enlace directo al Excel en línea.\n\n" +
        "La vista previa sigue disponible y funcionando."
      );
    }
  });
}
/* ======================================================================
 18 ABRIR EXCEL EN LÍNEA — OneDrive Corporativo (Graph)
====================================================================== */
const btnAbrirExcel = document.getElementById("visorAbrirExcel");

if (btnAbrirExcel) {
  btnAbrirExcel.addEventListener("click", async () => {
    const item = window.__archivoActual;

    if (!item?.fileIdentifierExcel) {
      alert("No se encontró el identificador del archivo.");
      return;
    }

    try {
      // ✅ Token MSAL (ya usado en tu app)
      const token = await obtenerToken();

      // ✅ Obtener el webUrl real desde OneDrive
      const resp = await fetch(
        `https://graph.microsoft.com/v1.0/me/drive/items/${item.fileIdentifierExcel}`,
        {
          headers: {
            Authorization: `Bearer ${token}`
          }
        }
      );

      if (!resp.ok) {
        throw new Error("Graph no devolvió el archivo");
      }

      const data = await resp.json();

      if (!data.webUrl) {
        throw new Error("El archivo no tiene webUrl");
      }

      // ✅ Abrir Excel real en línea (como antes)
      window.open(data.webUrl, "_blank");

    } catch (e) {
      console.error(e);
      alert(
        "No fue posible abrir el Excel en línea.\n" +
        "La vista previa sigue funcionando correctamente."
      );
    }
  });
}
