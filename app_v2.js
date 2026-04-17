/* ======================================================================
   0) IMPORTS — NECESARIOS
====================================================================== */
import { SITE_ID, LIBRARY_ID, FOLDER_PATH, listarArchivosMCI, obtenerModulo } from "./modulos_v2.js";
import { obtenerToken, iniciarSesion, usuarioActual, cerrarSesion } from "./auth.js";
import { obtenerURLTemporal } from "./graph_v2.js";

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

  if (!window.moduloActivo.pendientes) {
    document.getElementById("tbodyDatos").innerHTML = `
      <tr><td colspan="99" style="padding:20px; text-align:center;">
        No hay ruta configurada para este módulo.
      </td></tr>`;
    return;
  }

  const token = await obtenerToken();

  // ✅ ARCHIVOS DESDE SHAREPOINT
  const listaOD = await listarArchivosMCI(token);
  window.debugListaOD = listaOD;

  // ✅ KV
  const tecnico = "usuario";
  const respKV = await fetch(`https://cloudflare-index.modulo-de-exclusiones.workers.dev/consultar/${tecnico}`);
  const listaKV = await respKV.json();
  console.log("KV recibido:", listaKV);

  // ✅ Mezcla SP + KV
  listaOD.forEach(a => {
  const reg = listaKV.find(k => {
    const id = k.mciId || k.mcid;
    return id && a.nombre.includes(id);
  });

  a.mciId    = reg ? (reg.mciId || reg.mcid) : null;
  a.estadoKV = reg ? reg.estado : "pendiente";

  // ✅ fileId REAL viene de Graph (a.id)
  a.fileId = a.id || null;
});

  window.datosActuales = listaOD;

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
    item.nombre.endsWith(".xlsx") &&
    !item.nombre.includes("PreviewFotos")
  );

  filtrados.forEach(item => {
    const idx = window.datosActuales.indexOf(item);

    const tds = window.moduloActivo.columnas
      .map(col => `<td>${item[col.id]}</td>`)
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
   11) BUSCAR JSON DE FOTOS EN SHAREPOINT
====================================================================== */
async function obtenerJsonFotos(item) {

  const token = await obtenerToken();

  const urlListar = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${encodeURIComponent(FOLDER_PATH)}:/children`;

  const resp = await fetch(urlListar, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const data = await resp.json();
  if (!data.value) return null;

  const base = item.nombre;

  const jsonFile = data.value.find(f =>
    f.name.startsWith(base) &&
    f.name.endsWith(".json")
  );

  if (!jsonFile) return null;

  const urlContenido = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${jsonFile.id}/content`;

  const respJson = await fetch(urlContenido, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const jsonTexto = await respJson.text();
  try { return JSON.parse(jsonTexto); }
  catch { return null; }
}

/* ======================================================================
   12) VER ARCHIVO — PREVIEW EXCEL + JSON (ESTILO ORIGINAL)
====================================================================== */
async function verArchivo(item) {
   window.__archivoActual = item;

  console.log("🟨 item completo:", item);
  console.log("🟦 fileId en auditor:", item.fileId);


  // ✅ Ocultar tabla y mostrar modal
  document.getElementById("contenedor-modulo").style.display = "none";
  document.getElementById("modalVisor").style.display = "block";

  // ✅ Obtener token para Graph
  const token = await obtenerToken();

  // ✅ Descargar EXCEL desde SharePoint
  const urlDescarga = `https://graph.microsoft.com/v1.0${item.archivo.ruta}/content`;
  const resp = await fetch(urlDescarga, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  const arrayBuffer = await (await resp.blob()).arrayBuffer();

  // ✅ Leer Excel
  const wb = XLSX.read(arrayBuffer);
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // ✅ RANGOS ORIGINALES EXACTOS (IMPORTANTE: todas en let)
  let htmlInfoGeneral  = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B9:P18" });
  let htmlDescripcion  = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B69:P69" });
  let htmlDeclaracion  = XLSX.utils.sheet_to_html({ ...sheet, "!ref": "B71:M77" });

  // ✅ FUNCIÓN PARA PINTAR SOLO LOS CAMPOS REALES DEL EXCEL
  function pintarGris(html) {

    const campos = [
      "N° DE CASO",
      "FECHA",
      "CONTRATO No",
      "CONTRATISTA",
      "DEPARTAMENTO",
      "MUNICIPIO",
      "CENTRO POBLADO",
      "SEDE INSTITUCIÓN EDUCATIVA O CASO ESPECIAL",
      "ID BENEFICIARIO",
      "NOMBRE DEL RESPONSABLE (RESPONSABLE DE LA INSTITUCIÓN EDUCATIVA / AUTORIDAD COMPETENTE)",
      "NÚMERO DE CEDULA",
      "NÚMERO DE CONTACTO",
      "CORREO ELECTRÓNICO",
      "3. DESCRIPCIÓN DE LA FALLA / HALLAZGOS",
      "4. DECLARACIÓN",
      "DATOS DE QUIEN ACOMPAÑA EN EL CENTRO DIGITAL",
      "NOMBRES Y APELLIDOS",
      "CARGO",
      "NÚMERO DE CEDULA",
      "NÚMERO DE TELÉFONO O CELULAR 1",
      "NÚMERO DE TELÉFONO O CELULAR 2",
      "DATOS DE QUIEN REPARA EL SERVICIO EN EL CENTRO DIGITAL",
      "NÚMERO DE TELÉFONO O  CELULAR",
      "FIRMA"
    ];

    campos.forEach(t => {
      const rg = new RegExp(`(<td[^>]*>\\s*${t}[^<]*</td>)`, "gi");
      html = html.replace(rg, celda =>
        celda.replace(
          "<td",
          `<td style="background:#eef1f6; font-weight:700; border:1px solid #d6dce8;"`
        )
      );
    });

    return html;
  }

  // ✅ APLICAR A LOS TRES RANGOS
  htmlInfoGeneral  = pintarGris(htmlInfoGeneral);
  htmlDescripcion  = pintarGris(htmlDescripcion);
  htmlDeclaracion  = pintarGris(htmlDeclaracion);

  // ✅ Renderizar visor
  const visor = document.getElementById("visorIframe");

  visor.innerHTML = `
    <div style="
      background:white;
      padding:25px;
      border-radius:14px;
      border:1px solid #dce3f5;
      box-shadow:0 8px 24px rgba(0,0,0,.12);
    ">

      <!-- Encabezado 1 -->
      <div style="
        background:#eef1f6;
        padding:14px 18px;
        border-radius:10px;
        font-weight:800;
        font-size:15px;
        color:#203054;
        border:1px solid #d6dce8;
        margin-bottom:14px;
      ">
        Información del Beneficiario y la Institución
      </div>

      <div class="auditor-block">${htmlInfoGeneral}</div>

      <!-- Encabezado 2 -->
      <div style="
        background:#eef1f6;
        padding:14px 18px;
        border-radius:10px;
        font-weight:800;
        font-size:15px;
        color:#203054;
        border:1px solid #d6dce8;
        margin:28px 0 14px 0;
      ">
        Descripción del Caso
      </div>

      <div class="auditor-block">${htmlDescripcion}</div>

      <!-- Encabezado 3 -->
      <div style="
        background:#eef1f6;
        padding:14px 18px;
        border-radius:10px;
        font-weight:800;
        font-size:15px;
        color:#203054;
        border:1px solid #d6dce8;
        margin:28px 0 14px 0;
      ">
        Declaración
      </div>

      <div class="auditor-block">${htmlDeclaracion}</div>

      <h2 style="margin:30px 0 10px 0;">Fotos del informe (vista previa)</h2>
      <div id="visorFotos"></div>

    </div>
  `;

  // ✅ Cargar fotos (JSON)
  const jsonFotos = await obtenerJsonFotos(item);
  item.fotosPreview = jsonFotos;

  if (jsonFotos) await renderizarFotos(item);
  else document.getElementById("visorFotos").innerHTML =
    "<p style='color:#777;'>Este informe no tiene fotos adjuntas.</p>";
}
/* ======================================================================
   13) RENDER FOTOS — ESTILO DOMINION
====================================================================== */
async function renderizarFotos(item) {

  const cont = document.getElementById("visorFotos");
  const fotos = item.fotosPreview;
  if (!fotos) return;

  // === CONTENEDOR GRID RESPONSIVE (SIN ESPACIOS BLANCOS) ===
  let html = `
    <div style="
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
      gap: 22px;
      width: 100%;
    ">
  `;

  for (const clave in fotos) {

    const base64 = fotos[clave];
    if (!base64) continue;

    html += `
      <div style="
        background: #fff;
        border: 1px solid #dde5f8;
        border-radius: 12px;
        box-shadow: 0 6px 15px rgba(0,0,0,.08);
        overflow: hidden;
        display: flex;
        flex-direction: column;
      ">

        <!-- Título superior -->
        <div style="
          padding: 10px 12px;
          font-weight: 700;
          font-size: 14px;
        ">
          ${clave}
        </div>

        <!-- Imagen recortada, centrada y del mismo tamaño -->
        <img src="${base64}" style="
          width: 100%;
          height: 180px;
          object-fit: cover;
          object-position: center;
          display: block;
          border-radius: 0 0 12px 12px;
        ">
      </div>
    `;
  }

  html += `</div>`;
  cont.innerHTML = html;
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
