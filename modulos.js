/* ============================================================
   MODULOS.JS — Configuración de módulos del Panel Auditor
   ============================================================ */

export const MODULOS = {

  /* ============================================================
     ✅ MÓDULO MCI (OneDrive con driveId real)
     ============================================================ */
  MCI: {
    id: "mci",
    nombre: "Auditor — MCI",

    // ✅ Estas son las rutas correctas usando driveId real
    pendientes: "/drives/b!qDLeuVb8dE-_ocg255wGZSbL4Q0zxaNDvZnBorHVVnQq_CH66fH5Q6vXRgtmy0ua/root:/Documents/Base MCI - Proyecto automatización/MCI_Salidas",
    aprobados:  "/drives/b!qDLeuVb8dE-_ocg255wGZSbL4Q0zxaNDvZnBorHVVnQq_CH66fH5Q6vXRgtmy0ua/root:/Documents/Base MCI - Proyecto automatización/MCI_Aprobados",

    columnas: [
      { id: "tecnico",   label: "Técnico" },
      { id: "fecha",     label: "Fecha" },
      { id: "cliente",   label: "Cliente" },
      { id: "ubicacion", label: "Ubicación" }
    ],

    normalizar(item) {
      return {
        tecnico:   item.nombre ?? "—",
        fecha:     item.modificado || "—",
        cliente:   "—",
        ubicacion: "—",
        archivo: {
          nombre: item.nombre,
          ruta:   item.ruta,
          tamano: item.tamano,
          tipo:   item.tipo
        }
      };
    }
  },

  /* ============================================================
     ✅ MÓDULO MPR (cuando toque)
     ============================================================ */
  MPR: {
    id: "mpr",
    nombre: "Auditor — MPR",

    pendientes: null,
    aprobados: null,

    columnas: [
      { id: "tecnico",   label: "Técnico" },
      { id: "fecha",     label: "Fecha" },
      { id: "proyecto",  label: "Proyecto" },
      { id: "zona",      label: "Zona" }
    ],

    normalizar(item) {
      return {
        tecnico:   item.nombre ?? "—",
        fecha:     item.modificado || "—",
        proyecto:  "—",
        zona:      "—",
        archivo: {
          nombre: item.nombre,
          ruta:   item.ruta,
          tamano: item.tamano,
          tipo:   item.tipo
        }
      };
    }
  }

};

/* ============================================================
   🔧 FUNCIÓN AUXILIAR
   ============================================================ */
export function obtenerModulo(id) {
  return MODULOS[id.toUpperCase()] ?? null;
}
