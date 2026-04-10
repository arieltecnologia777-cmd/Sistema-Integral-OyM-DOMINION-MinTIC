// ======================================================================
// GRAPH PARA SHAREPOINT ONLINE — VERSIÓN CORRECTA PARA TU AUDITOR
// ======================================================================

import { SITE_ID, LIBRARY_ID } from "./modulos_v2.js";
import { obtenerToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// ======================================================================
// HELPER: LLAMAR GRAPH CON TOKEN
// ======================================================================

async function graphFetchJson(url, method = "GET", body = null) {
    const token = await obtenerToken();
    if (!token) throw new Error("❌ Token no disponible");

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
        const err = await resp.text();
        console.error("❌ Error Graph:", resp.status, err);
        throw new Error(err);
    }

    return resp.json();
}

async function graphFetchRaw(url) {
    const token = await obtenerToken();
    return fetch(url, {
        headers: { "Authorization": `Bearer ${token}` }
    });
}


// ======================================================================
// FUNCIÓN COMPATIBLE — (NO SE USA CON SHAREPOINT, SOLO SE DEJA POR LEGADO)
// ======================================================================

export async function listarArchivos() {
    return [];
}


// ======================================================================
// OBTENER URL TEMPORAL (VIEW LINK) PARA PREVIEW
// ======================================================================

export async function obtenerURLTemporal(rutaItem) {
    const url = `${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}${rutaItem}/createLink`;

    const json = await graphFetchJson(url, "POST", {
        type: "view",
        scope: "anonymous"
    });

    return json?.link?.webUrl ?? null;
}


// ======================================================================
// MOVER ARCHIVO (APROBAR)
// ======================================================================

export async function moverArchivo(rutaOrigen, rutaDestino) {

    const nombreNuevo = rutaDestino.split("/").pop();
    const carpetaDestino = rutaDestino.replace(`/${nombreNuevo}`, "");

    const body = {
        parentReference: {
            driveId: LIBRARY_ID,
            id: carpetaDestino
        },
        name: nombreNuevo
    };

    const url = `${GRAPH_BASE}/sites/${SITE_ID}/drives/${LIBRARY_ID}${rutaOrigen}`;

    try {
        await graphFetchJson(url, "PATCH", body);
        return true;
    } catch (err) {
        console.error("❌ Error moviendo archivo:", err);
        return false;
    }
}


// ======================================================================
// CARGAR ARCHIVOS DESDE CARPETA (NO SE USA PARA LISTAR MCI)
// PERO SE DEJA POR COMPATIBILIDAD
// ======================================================================

export async function cargarDesdeCarpeta(modulo) {
    // Tu app ya usa listarArchivosMCI() desde modulos_v2.js
    return [];
}
