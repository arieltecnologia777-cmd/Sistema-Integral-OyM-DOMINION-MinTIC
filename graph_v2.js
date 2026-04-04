import { DRIVE_ID } from "./modulos_v2.js";
import { obtenerToken } from "./auth.js";

// ============================================================
// TOKEN AUTOMÁTICO
// ============================================================
async function graphFetch(url, method = "GET", body = null) {
  const token = await obtenerToken();

  if (!token) {
    console.error("❌ Token inválido");
    throw new Error("Token no disponible");
  }

  const options = {
    method,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  };

  if (body) options.body = JSON.stringify(body);

  const resp = await fetch(url, options);

  if (!resp.ok) {
    const t = await resp.text();
    console.error("❌ Error Graph:", resp.status, t);
    throw new Error(`Graph error ${resp.status}: ${t}`);
  }

  return resp.json();
}

// ============================================================
// LISTAR ARCHIVOS
// ============================================================
export async function listarArchivos(folderId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${folderId}/children`;
  const data = await graphFetch(url);

  const archivos = [];

  for (const x of data.value) {
    // ✅ Segundo llamado para obtener downloadUrl real
    const detalle = await graphFetch(
      `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${x.id}`
    );

    archivos.push({
      id: x.id,
      nombre: x.name,
      fecha: x.lastModifiedDateTime,
      tamano: x.size,
      archivo: {
        ruta: `/drives/${DRIVE_ID}/items/${x.id}`,
        nombre: x.name,
        webUrl: x.webUrl,
        downloadUrl: detalle["@microsoft.graph.downloadUrl"] // ✅ AHORA SÍ SIRVE
      }
    });
  }

  return archivos;
}

// ============================================================
// PREVIEW
// ============================================================
export async function obtenerURLTemporal(ruta) {
  const url = `https://graph.microsoft.com/v1.0${ruta}/createLink`;

  const data = await graphFetch(url, "POST", {
    type: "view",
    scope: "anonymous"
  });

  return data?.link?.webUrl ?? null;
}

// ============================================================
// MOVER ARCHIVO (aprobar)
// ============================================================
export async function moverArchivo(rutaOrigen, rutaDestino) {
  const nombre = rutaDestino.split("/").pop();
  const carpetaDestino = rutaDestino.replace(`/${nombre}`, "");

  const body = {
    parentReference: {
      driveId: DRIVE_ID,
      id: carpetaDestino
    },
    name: nombre
  };

  const url = `https://graph.microsoft.com/v1.0${rutaOrigen}`;

  try {
    await graphFetch(url, "PATCH", body);
    return true;
  } catch (err) {
    console.error("❌ Error moviendo archivo:", err);
    return false;
  }
}

// ============================================================
// CARGA CENTRAL
// ============================================================
export async function cargarDesdeCarpeta(modulo) {

  const token = await obtenerToken();
  const archivos = await listarArchivos(modulo.pendientes);

  // Vamos a construir item por item manualmente
  const items = [];

  for (const a of archivos) {

    // 1. Construimos el item base (como antes)
    const item = {
      nombre: a.nombre,
      fecha: new Date(a.fecha).toLocaleString("es-CO"),
      tamano: a.tamano,
      archivo: a.archivo,
      fotosPreview: null  // <-- aquí guardaremos las 8 fotos
    };

    // 2. Nombre esperado del archivo JSON
    const previewJsonName = a.nombre.replace(".xlsx", "_PreviewFotos.json");

    // 3. Buscar archivo JSON entre los archivos listados
    const previewEntry = archivos.find(f => f.nombre === previewJsonName);

    if (previewEntry) {
      try {
        const resp = await fetch(
          `https://graph.microsoft.com/v1.0${previewEntry.archivo.ruta}/content`,
          { headers: { "Authorization": `Bearer ${token}` } }
        );

        const jsonText = await resp.text();
        item.fotosPreview = JSON.parse(jsonText);

      } catch (err) {
        console.error("❌ Error cargando PreviewFotos.json:", err);
      }
    }

    // 4. Guardamos este item completo
    items.push(item);
  }

  return items;
}
