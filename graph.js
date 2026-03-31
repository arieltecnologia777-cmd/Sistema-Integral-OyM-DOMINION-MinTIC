import { obtenerToken } from "./auth.js";

/* ============================================================
   Helper: Fetch autenticado a Microsoft Graph
   ============================================================ */
async function graphFetch(url, method = "GET", body = null) {
  const token = await obtenerToken();
  if (!token) {
    console.error("❌ No se pudo obtener el token de acceso.");
    return null;
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
    console.error("❌ Error en Graph:", resp.status, await resp.text());
    return null;
  }

  // Graph retorna JSON
  return resp.json();
}

/* ============================================================
   1) LISTAR ARCHIVOS — usando driveId  
   (NO SE USA /me !!!)
   ============================================================ */
export async function listarArchivos(rutaCarpeta) {
  if (!rutaCarpeta) {
    console.warn("⚠️ Ruta vacía.");
    return [];
  }

  // ✅ Ruta correcta → Graph con driveId real
  const url = `https://graph.microsoft.com/v1.0${rutaCarpeta}:/children`;

  const data = await graphFetch(url);
  if (!data || !data.value) return [];

  return data.value.filter(item => item.file);
}

/* ============================================================
   2) OBTENER ARCHIVO (Blob)
   ============================================================ */
export async function obtenerArchivo(rutaArchivo) {
  const token = await obtenerToken();
  if (!token) return null;

  const url = `https://graph.microsoft.com/v1.0${rutaArchivo}:/content`;

  const resp = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  if (!resp.ok) {
    console.error("❌ No se pudo obtener el archivo:", resp.status);
    return null;
  }

  return resp.blob();
}

/* ============================================================
   3) MOVER ARCHIVO (Aprobar)
   ============================================================ */
export async function moverArchivo(rutaOrigen, rutaDestino) {

  // Nombre del archivo
  const nombre = rutaDestino.split("/").pop();

  // Carpeta destino → quitar el nombre del archivo
  const carpetaDestino = rutaDestino.replace(`/${nombre}`, "");

  const body = {
    parentReference: {
      path: carpetaDestino
    },
    name: nombre
  };

  const url = `https://graph.microsoft.com/v1.0${rutaOrigen}`;

  const resp = await graphFetch(url, "PATCH", body);

  if (!resp) {
    console.error("❌ Error moviendo archivo.");
    return false;
  }

  console.log("✅ Archivo movido");
  return true;
}

/* ============================================================
   4) CARGAR ARCHIVOS NORMALIZADOS
   ============================================================ */
export async function cargarDesdeCarpeta(modulo, esAprobados = false) {
  const ruta = esAprobados ? modulo.aprobados : modulo.pendientes;

  const archivos = await listarArchivos(ruta);
  if (!archivos || archivos.length === 0) return [];

  return archivos.map(file =>
    modulo.normalizar({
      nombre: file.name,
      ruta: `${ruta}/${file.name}`,
      modificado: file.lastModifiedDateTime ?? "—",
      tamano: file.size ?? "—",
      tipo: file.file?.mimeType ?? "—"
    })
  );
}

/* ============================================================
   5) URL TEMPORAL PARA PREVIEW
   ============================================================ */
export async function obtenerURLTemporal(rutaArchivo) {
  const blob = await obtenerArchivo(rutaArchivo);
  if (!blob) return null;

  return URL.createObjectURL(blob);
}
