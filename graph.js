/* ======================================================================
   GRAPH.JS — Panel Auditor
   Acceso a OneDrive usando Microsoft Graph API.
   Funciones esenciales:
   - listar archivos de una carpeta
   - obtener metadata
   - descargar archivo
   - mover archivo (aprobar)
   Ariel-friendly — versión estable y limpia.
====================================================================== */

import { obtenerToken } from "./auth.js";

/* ======================================================================
   Helper: Peticiones autenticadas a Microsoft Graph
====================================================================== */
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

  return resp.json();
}

/* ======================================================================
   1) LISTAR ARCHIVOS EN UNA CARPETA (OneDrive personal del usuario)
   Ejemplo ruta: "/drive/root:/Documents/Base MCI - Proyecto automatización/MCI_Salidas"
====================================================================== */
export async function listarArchivos(rutaCarpeta) {
  if (!rutaCarpeta) {
    console.warn("⚠️ Ruta vacía. Aún no configurada.");
    return [];
  }

  // ✅ Usamos OneDrive personal del usuario: /me/drive/...
  const url = `https://graph.microsoft.com/v1.0/me${rutaCarpeta}:/children`;

  const data = await graphFetch(url);
  if (!data || !data.value) return [];

  // ✅ Filtrar SOLO archivos
  return data.value.filter(item => item.file);
}

/* ======================================================================
   2) OBTENER CONTENIDO BLOb de un archivo (para abrir en vista previa)
====================================================================== */
export async function obtenerArchivo(rutaArchivo) {
  const token = await obtenerToken();
  if (!token) return null;

  // ✅ Descargar contenido
  const url = `https://graph.microsoft.com/v1.0/me${rutaArchivo}:/content`;

  const resp = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  if (!resp.ok) {
    console.error("❌ No se pudo obtener el archivo:", resp.status);
    return null;
  }

  return resp.blob(); // Excel, PDF, imágenes…
}

/* ======================================================================
   3) MOVER ARCHIVO (Aprobar)
   Origen: /drive/root:/Documents/.../MCI_Salidas/archivo.xlsx
   Destino: /drive/root:/Documents/.../MCI_Aprobados/archivo.xlsx
====================================================================== */
export async function moverArchivo(rutaOrigen, rutaDestino) {
  // Extraer nombre
  const partes = rutaDestino.split("/");
  const nombreArchivo = partes.pop();
  const carpetaDestino = partes.join("/");

  // Cuerpo PATCH
  const body = {
    parentReference: {
      path: `/drive/root:${carpetaDestino}`
    },
    name: nombreArchivo
  };

  const url = `https://graph.microsoft.com/v1.0/me${rutaOrigen}`;

  const resp = await graphFetch(url, "PATCH", body);

  if (!resp) {
    console.error("❌ Error moviendo archivo.");
    return false;
  }

  console.log("✅ Archivo movido correctamente.");
  return true;
}

/* ======================================================================
   4) LEER UNA CARPETA COMPLETA Y NORMALIZARLA PARA LA TABLA
====================================================================== */
export async function cargarDesdeCarpeta(modulo, esAprobados = false) {
  const ruta = esAprobados ? modulo.aprobados : modulo.pendientes;

  const archivos = await listarArchivos(ruta);
  if (!archivos || archivos.length === 0) return [];

  // ✅ Normalizar usando el módulo activo (MCI o MPR)
  return archivos.map(file =>
    modulo.normalizar({
      nombre: file.name,
      ruta: `${ruta}/${file.name}`,  // ruta completa para abrir/mover
      modificado: file.lastModifiedDateTime ?? "—",
      tamano: file.size ?? "—",
      tipo: file.file?.mimeType ?? "—"
    })
  );
}

/* ======================================================================
   5) CREAR URL TEMPORAL PARA PREVIEW (Excel, PDF, etc.)
====================================================================== */
export async function obtenerURLTemporal(rutaArchivo) {
  const blob = await obtenerArchivo(rutaArchivo);
  if (!blob) return null;

  return URL.createObjectURL(blob); // ✅ Vista previa directa
}
