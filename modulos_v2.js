// ======================================================
// CONFIG — IDs confirmados por Graph
// ======================================================

export const DRIVE_ID = "b!qDLeuVb8dE-_ocg255wGZSbL4Q0zxaNDvZnBorHVVnQq_CH66fH5Q6vXRgtmy0ua";

export const FOLDERS = {
 pendientes: "01IWRV3S42TYZGW7YV7VDZSYL6SVNMT5NY",   // ✅ TU CARPETA REAL
 aprobados: "01IWRV3S7JHBELGMR54FAYX3Z3HRZFVODA"    // este se usa si algún día decides mover
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

    // === 1) Separar excels y previewFotos ===
    // === 1) Separar excels y previewFotos ===
const excels = data.value.filter(f => {
  const isExcel = f.name.endsWith(".xlsx");

  if (isExcel) {
    console.log("✅ ARCHIVO GRAPH ORIGINAL:", f);  // <-- ESTE SÍ IMPRIME EL OBJETO REAL
  }

  return isExcel;
});

const previews = data.value.filter(f => f.name.includes("PreviewFotos"));

    const lista = [];

    for (const x of excels) {
    const metaResp = await fetch(
    `${GRAPH_BASE}/drives/${DRIVE_ID}/items/${x.id}?$select=parentReference`,
    { headers: { "Authorization": `Bearer ${token}` } }
);
const meta = await metaResp.json();    
    const item = {
  id: x.id,
  nombre: x.name,

 // ✅ Fecha REAL desde OneDrive (UTC)
fechaReal: x.fileSystemInfo?.lastModifiedDateTime,

// ✅ Fecha REAL humana (convertida a tu hora local Colombia)
fecha: (() => {
  const d = new Date(x.fileSystemInfo?.lastModifiedDateTime);
  const pad = n => String(n).padStart(2, "0");
  
  // ✅ getHours(), getMinutes(), etc → usan tu zona local (UTC-5)
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
})(),

  tamano: formatearTamano(x.size),

  archivo: {
    ruta: `/drives/${DRIVE_ID}/items/${x.id}`,
    nombre: x.name,

    // ✅ fileIdReal usado por KV
    fileIdReal: `${DRIVE_ID}.${x.id}`,

    // ✅ Carpeta real (opcional)
    carpeta: meta?.parentReference?.path ?? null
},

  fotosPreview: null
};

        // Buscar su archivo JSON correspondiente
        const base = x.name.replace(".xlsx", "");
        const jsonMatch = previews.find(p => p.name.startsWith(base));

        if (jsonMatch) {
            try {
                const urlJ = `${GRAPH_BASE}/drives/${DRIVE_ID}/items/${jsonMatch.id}/content`;
                const respJ = await fetch(urlJ, {
                    headers: { "Authorization": `Bearer ${token}` }
                });
                const texto = await respJ.text();
                item.fotosPreview = JSON.parse(texto);
            } catch(e) {
                console.error("Error leyendo fotosPreview:", e);
            }
        }

        lista.push(item);
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
