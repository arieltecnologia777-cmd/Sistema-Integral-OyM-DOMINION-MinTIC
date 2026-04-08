// ======================================================
// CONFIG — IDs confirmados por Graph
// ======================================================

export const DRIVE_ID = "b!qDLeuVb8dE-_ocg255wGZSbL4Q0zxaNDvZnBorHVVnQq_CH66fH5Q6vXRgtmy0ua";

export const FOLDERS = {
  pendientes: "01IWRV3SZ7VKZ6DTAIUNDZ4GDTQ7RDSN34",   // ✅ EL VERDADERO
  aprobados: "01IWRV3S7JHBELGMR54FAYX3Z3HRZFVODA"
};

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// ======================================================
// DEFINICIÓN DE MÓDULOS (lo que usa app.js)
// ======================================================

export const MODULOS = {
    mci: {
        columnas: [
            { id: "nombre", label: "Archivo" },
            { id: "fecha", label: "Fecha" },
            { id: "tamano", label: "Tamaño" }
        ],

        pendientes: FOLDERS.pendientes,
        aprobados: FOLDERS.aprobados
    }
};

// ======================================================
// FUNCIÓN QUE NECESITA app.js
// ======================================================

export function obtenerModulo(nombre) {
    return MODULOS[nombre] || null;
}

// ======================================================
// LISTAR ARCHIVOS (pendientes) — SHAREPOINT
// ======================================================
export async function listarArchivosMCI(token) {

  // ============================
  // 🔧 CONFIG SHAREPOINT (ANCLAS)
  // ============================
  // ⚠️ DEBES AJUSTAR ESTOS 2 VALORES
  const SITE_ID = "TU_SITE_ID_SHAREPOINT";
  // Ejemplo real:
  // dominio.sharepoint.com,aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee,ffffffff-1111-2222-3333-444444444444

  const LIBRARY_ID = "TU_DOCUMENT_LIBRARY_ID";
  // Normalmente es la biblioteca "Documentos"

  const FOLDER_NAME = "MCI_Generados";
  // ============================

  const url = `${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${FOLDER_NAME}:/children`;

  const res = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  if (!res.ok) {
    console.error("❌ Error listando archivos desde SharePoint:", res.status);
    return [];
  }

  const data = await res.json();

  if (!data.value || !Array.isArray(data.value)) {
    console.warn("⚠️ SharePoint no devolvió archivos válidos");
    return [];
  }

  // ============================
  // ✅ Filtrar solo Excel
  // ============================
  const excels = data.value.filter(f =>
    f.name && f.name.endsWith(".xlsx")
  );

  const lista = [];

  for (const x of excels) {

    const d = new Date(x.lastModifiedDateTime);
    const pad = n => String(n).padStart(2, "0");

    lista.push({
      id: x.id,
      nombre: x.name,

      // Fecha real (UTC)
      fechaReal: x.lastModifiedDateTime,

      // Fecha formateada (hora local)
      fecha: `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`,

      tamano: formatearTamano(x.size),

      archivo: {
        // 👉 ESTA RUTA ES LA QUE USA verArchivo() PARA EMBEBER
        ruta: `/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${x.id}`,
        nombre: x.name,
        fileIdReal: x.id,
        carpeta: FOLDER_NAME
      },

      // 🚫 Preview hack eliminado
      fotosPreview: null
    });
  }

  return lista;
}


// ======================================================
// DESCARGAR ARCHIVO
// ======================================================

export async function descargarArchivo(token, fileId) {
    const url = `${GRAPH_BASE}/drives/${DRIVE_ID}/items/${fileId}/content`;

    const res = await fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });

    return res;
}

// ======================================================
// FORMATOS
// ======================================================

export function formatearFecha(f) {
  return new Date(f).toLocaleString("es-CO", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  });
}

export function formatearTamano(b) {
    if (b < 1024) return b + " B";
    if (b < 1024 * 1024) return (b / 1024).toFixed(1) + " KB";
    return (b / 1024 / 1024).toFixed(1) + " MB";
}
