/* ======================================================================
   GRAPH.JS — Panel Auditor
   Acceso a OneDrive usando Microsoft Graph API.
   Funciones esenciales:
    - listar archivos de una carpeta
    - obtener metadata
    - descargar archivo
    - mover archivo (aprobar)
   ====================================================================== */

import { obtenerToken } from "./auth.js";

/* ======================================================================
   Helper: Peticiones a Graph
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
   1) LISTAR ARCHIVOS EN UNA CARPETA
   input: ruta → "/drive/root:/Carpeta/Subcarpeta"
   ====================================================================== */
export async function listarArchivos(rutaCarpeta) {

  if (!rutaCarpeta) {
    console.warn("⚠️ Ruta vacía. Aún no configurada.");
    return [];
  }

  const url = `https://graph.microsoft.com/v1.0/me${rutaCarpeta}:/children`;

  const data = await graphFetch(url);
  if (!data || !data.value) return [];

  // Filtramos solo archivos (OneDrive retorna carpetas también)
  return data.value.filter(item => item.file);
}

/* ======================================================================
   2) OBTENER CONTENIDO/METADATA DE UN ARCHIVO
   input: rutaArchivo → "/drive/root:/Carpeta/file.json"
   ====================================================================== */
export async function obtenerArchivo(rutaArchivo) {

  const token = await obtenerToken();
  if (!token) return null;

  const url = `https://graph.microsoft.com/v1.0/me${rutaArchivo}:/content`;

  const resp = await fetch(url, {
    headers: { "Authorization": `Bearer ${token}` }
  });

  if (!resp.ok) {
    console.error("❌ No se pudo obtener el archivo:", resp.status);
    return null;
  }

  // Retorna Blob (PDF, XLSX, JSON, etc.)
  const blob = await resp.blob();
  return blob;
}

/* ======================================================================
   3) MOVER ARCHIVO (Aprobar)
   input:
    - rutaOrigen  → "/drive/root:/MCI/pendientes/archivo.pdf"
    - rutaDestino → "/drive/root:/MCI/aprobados/archivo.pdf"
   ====================================================================== */
export async function moverArchivo(rutaOrigen, rutaDestino) {

  const partes = rutaDestino.split("/");
  const nombreArchivo = partes.pop();
  const carpetaDestino = partes.join("/");

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
   4) LEER ARCHIVOS Y RETORNAR OBJETOS NORMALIZADOS PARA LA TABLA
      - módulo = MODULOS.MCI o MODULOS.MPR
      - ruta = módulo.pendientes o módulo.aprobados
   ====================================================================== */
export async function cargarDesdeCarpeta(modulo, esAprobados = false) {

  const ruta = esAprobados ? modulo.aprobados : modulo.pendientes;

  const archivos = await listarArchivos(ruta);
  if (!archivos || archivos.length === 0) {
    return [];
  }

  // Construimos array limpio para la tabla
  const resultados = archivos.map(file => {
    return modulo.normalizar({
      tecnico: file?.lastModifiedBy?.user?.displayName ?? "—",
      fecha: file?.lastModifiedDateTime ?? "—",
      cliente: file?.description ?? "—",
      ubicacion: "—",
      proyecto: file?.description ?? "—",
      zona: "—",
      archivo: {
        nombre: file.name,
        id: file.id,
        tipo: file.file.mimeType,
        ruta: `${ruta}/${file.name}`   // RUTA COMPLETA
      }
    });
  });

  return resultados;
}

/* ======================================================================
   5) DESCARGAR ARCHIVO COMO URL TEMPORAL (para vista previa PDF/Img)
   ====================================================================== */
export async function obtenerURLTemporal(rutaArchivo) {

  const blob = await obtenerArchivo(rutaArchivo);
  if (!blob) return null;

  return URL.createObjectURL(blob);
}
