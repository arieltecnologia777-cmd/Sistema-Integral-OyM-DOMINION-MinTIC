// ======================================================================
// CONFIG PARA SHAREPOINT ONLINE (PARTE 1/2)
// ======================================================================

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/*
    ⚠️ REEMPLAZA ESTOS 3 VALORES CON LOS TUYOS:

    1) SITE_ID:
       Ejemplo:
       "dominionglobal.sharepoint.com,XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX,YYYYYYYY-YYYY-YYYY-YYYY-YYYYYYYYYYYY"

    2) LIBRARY_ID:
       Ejemplo:
       "b!zzZZZzzzZZzzZZzZZ_zZzZzzZZZZzzZzzZ"

    3) FOLDER_PATH:
       Ejemplo:
       "Base MCI - Proyecto automatización/MCI_Generados"
*/

export const SITE_ID      = "";   // ← TU SITE ID REAL AQUÍ
export const LIBRARY_ID   = "";   // ← TU LIBRARY ID REAL AQUÍ
export const FOLDER_PATH  = "";   // ← TU RUTA REAL AQUÍ


// ======================================================================
// DEFINICIÓN DE MÓDULOS
// ======================================================================

export const MODULOS = {
    mci: {
        columnas: [
            { id: "nombre", label: "Archivo" },
            { id: "fecha",  label: "Fecha" },
            { id: "tamano", label: "Tamaño" }
        ],

        // SE USA SOLO FOLDER_PATH
        pendientes: FOLDER_PATH,

        // Si luego tienes carpeta de aprobados:
        aprobados: null
    }
};


// ======================================================================
// OBTENER CONFIGURACIÓN DE MÓDULO
// ======================================================================

export function obtenerModulo(nombre) {
    return MODULOS[nombre] || null;
}
// ======================================================================
// LISTAR ARCHIVOS DESDE SHAREPOINT (PARTE 2/2)
// ======================================================================

export async function listarArchivosMCI(token) {

    if (!SITE_ID || !LIBRARY_ID || !FOLDER_PATH) {
        console.error("❌ ERROR: modulos_v2.js no tiene SITE_ID / LIBRARY_ID / FOLDER_PATH configurados.");
        return [];
    }

    const url =
`${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/root:/${encodeURIComponent(FOLDER_PATH)}:/children`;

    const res = await fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });

    if (!res.ok) {
        console.error("❌ Error listando archivos MCI:", await res.text());
        return [];
    }

    const data = await res.json();
    if (!Array.isArray(data.value)) return [];

    const lista = [];

    // ✅ Filtrar solo los excels reales
    const excels = data.value.filter(f => f.name.endsWith(".xlsx"));

    for (const x of excels) {

        const d = new Date(x.lastModifiedDateTime);
        const pad = n => String(n).padStart(2, "0");

        const item = {
            id: x.id,
            nombre: x.name,

            // ✅ Fecha UTC de SharePoint
            fechaReal: x.lastModifiedDateTime,

            // ✅ Fecha local humana
            fecha: `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`,

            tamano: formatearTamano(x.size),

            archivo: {
                ruta: `/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${x.id}`,
                nombre: x.name,
                fileIdReal: x.id,
                carpeta: FOLDER_PATH
            },

            fotosPreview: null   // ← tu visor usa esto
        };

        lista.push(item);
    }

    return lista;
}


// ======================================================================
// DESCARGAR ARCHIVO DESDE SHAREPOINT
// ======================================================================

export async function descargarArchivo(token, fileId) {
    const url = `${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}/items/${fileId}/content`;

    const res = await fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });

    return res;
}


// ======================================================================
// FORMATOS
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
