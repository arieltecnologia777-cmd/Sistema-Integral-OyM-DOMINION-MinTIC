// ======================================================
// CONFIG — IDs confirmados por Graph
// ======================================================

export const DRIVE_ID = "b!qDLeuVb8dE-_ocg255wGZSbL4Q0zxaNDvZnBorHVVnQq_CH66fH5Q6vXRgtmy0ua";

export const FOLDERS = {
    pendientes: "01IWRV3SZ7VKZ6DTAIUNDZ4GDTQ7RDSN34",   // MCI_Salidas
    aprobados: "01IWRV3S7JHBELGMR54FAYX3Z3HRZFVODA"     // MCI_Aprobados
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
// LISTAR ARCHIVOS (pendientes)
// ======================================================

export async function listarArchivosMCI(token) {
    const url = `${GRAPH_BASE}/drives/${DRIVE_ID}/items/${FOLDERS.pendientes}/children`;

    const res = await fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });

    const data = await res.json();

    return data.value.map(x => ({
        id: x.id,
        nombre: x.name,
        fecha: formatearFecha(x.lastModifiedDateTime),
        tamano: formatearTamano(x.size),
        archivo: {
            ruta: `/drives/${DRIVE_ID}/items/${x.id}`,
            nombre: x.name
        }
    }));
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
    return new Date(f).toLocaleString("es-CO");
}

export function formatearTamano(b) {
    if (b < 1024) return b + " B";
    if (b < 1024 * 1024) return (b / 1024).toFixed(1) + " KB";
    return (b / 1024 / 1024).toFixed(1) + " MB";
}
