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
  // CONFIG SHAREPOINT
  // ============================
  const SITE_ID = "dominionglobal.sharepoint.com,10887380-6934-45ab-adf2-97b2aad78311,433b4a3a-96f7-4bf3-929a-eb5f824c466d";

  const LIBRARY_ID = "b!gHOIEDRpq0Wt8peyqteDETpKO0P3lvNLkprrX4JMRm3TjVug-HIEToXXjMDkE9rM";

  const FOLDER_NAME = "Base MCI - Proyecto automatización/MCI_Generados";

  // ============================
  // URL CON RUTA CORRECTA
  // ============================
  const url = `${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${encodeURIComponent(FOLDER_NAME)}:/children`;

  const res = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  if (!res.ok) {
    console.error("❌ Error listando archivos desde SharePoint:", res.status);
    return [];
  }

  const data = await res.json();

  if (!Array.isArray(data.value)) {
    console.warn("⚠️ SharePoint no devolvió una lista válida");
    return [];
  }

  const lista = [];

  for (const x of data.value) {
    if (!x.name || !x.name.endsWith(".xlsx")) continue;

    const d = new Date(x.lastModifiedDateTime);
    const pad = n => String(n).padStart(2, "0");

    lista.push({
      id: x.id,
      nombre: x.name,
      fechaReal: x.lastModifiedDateTime,
      fecha: `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`,
      tamano: formatearTamano(x.size),
      archivo: {
        ruta: `/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${x.id}`,
        nombre: x.name,
        fileIdReal: x.id,
        carpeta: FOLDER_NAME
      },
      fotosPreview: null
    });
  }

  return lista;
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
