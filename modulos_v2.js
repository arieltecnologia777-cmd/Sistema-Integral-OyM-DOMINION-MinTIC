// ======================================================================
// CONFIG SHAREPOINT ONLINE — ID REALES Y VALORES CORRECTOS
// ======================================================================

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// ✅ DATOS REALES EXTRAÍDOS DE TU ARCHIVO ANTERIOR (EL QUE SÍ FUNCIONABA)
export const SITE_ID =
 "dominionglobal.sharepoint.com,10887380-6934-45ab-adf2-97b2aad78311,433b4a3a-96f7-4bf3-929a-eb5f824c466d";

export const LIBRARY_ID =
 "b!gHOIEDRpq0Wt8peyqteDETpKO0P3lvNLkprrX4JMRm3TjVug-HIEToXXjMDkE9rM";

// ✅ Carpeta REAL que usabas en tu versión estable
export const FOLDER_PATH =
 "Base MCI - Proyecto automatización/MCI_Generados";


// ======================================================================
// DEFINICIÓN DE MÓDULOS (el Auditor usa este objeto)
// ======================================================================

export const MODULOS = {
    mci: {
        columnas: [
            { id: "nombre", label: "Archivo" },
            { id: "fecha",  label: "Fecha" },
            { id: "tamano", label: "Tamaño" }
        ],

        // ✅ carpeta de pendientes
        pendientes: FOLDER_PATH,

        // ✅ si luego quieres agregamos aprobados
        aprobados: null
    }
};


// ======================================================================
// OBTENER CONFIGURACIÓN DEL MÓDULO
// ======================================================================

export function obtenerModulo(nombre) {
    return MODULOS[nombre] || null;
}
// ======================================================================
// LISTAR ARCHIVOS DESDE SHAREPOINT
// ======================================================================

export async function listarArchivosMCI(token) {

    const url =
`${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${encodeURIComponent(FOLDER_PATH)}:/children`;

    const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` }
    });

    if (!res.ok) {
        console.error("❌ Error listando en SharePoint:", await res.text());
        return [];
    }

    const data = await res.json();
    if (!Array.isArray(data.value)) return [];

    const lista = [];

    // ✅ Solo traer excels reales
    const excels = data.value.filter(f => f.name.endsWith(".xlsx"));

    for (const x of excels) {

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
                carpeta: FOLDER_PATH
            },

            fotosPreview: null
        });
    }

    return lista;
}


// ======================================================================
// DESCARGAR DESDE SHAREPOINT
// ======================================================================

export async function descargarArchivo(token, fileId) {
    const url =
`${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${fileId}/content`;

    const res = await fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });

    return res;
}


// ======================================================================
// FORMATOS — FECHA Y TAMAÑO
// ======================================================================

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
