/* ============================================================
   MODULOS.JS
   Configuración de módulos del Panel Auditor
   Sistema escalable sin tocar flows actuales

   Cada módulo define:
   - Nombre visible
   - Carpetas en OneDrive (pendientes/aprobados)
   - Columnas a mostrar en la tabla
   - Normalización opcional (si un módulo usa campos distintos)
   ============================================================ */

export const MODULOS = {

  /* ============================================================
     ✅ MÓDULO MCI
     ------------------------------------------------------------
     Este módulo debe respetar EXACTAMENTE las rutas existentes
     que usa tu Power Automate. NO las llenamos aquí aún.
     ============================================================ */
  MCI: {
    id: "mci",
    nombre: "Auditor — MCI",

    pendientes: "/drive/root:/Documents/Base MCI - Proyecto automatización/MCI_Salidas",
    aprobados: "/drive/root:/Documents/Base MCI - Proyecto automatización/MCI_Aprobados",

    columnas: [
      { id: "tecnico",     label: "Técnico" },
      { id: "fecha",       label: "Fecha" },
      { id: "cliente",     label: "Cliente" },
      { id: "ubicacion",   label: "Ubicación" }
    ],

    // Normalización de datos (placeholder)
    normalizar(item) {
      return {
        tecnico:   item?.tecnico ?? "—",
        fecha:     item?.fecha ?? "—",
        cliente:   item?.cliente ?? "—",
        ubicacion: item?.ubicacion ?? "—",
        archivo:   item?.archivo ?? null
      };
    }
  },

  /* ============================================================
     ✅ MÓDULO MPR
     ------------------------------------------------------------
     Similar a MCI. Aún sin rutas reales.
     ============================================================ */
  MPR: {
    id: "mpr",
    nombre: "Auditor — MPR",

    pendientes: null, // las pediré cuando toque
    aprobados: null,

    columnas: [
      { id: "tecnico",   label: "Técnico" },
      { id: "fecha",     label: "Fecha" },
      { id: "proyecto",  label: "Proyecto" },
      { id: "zona",      label: "Zona" }
    ],

    normalizar(item) {
      return {
        tecnico:  item?.tecnico ?? "—",
        fecha:    item?.fecha ?? "—",
        proyecto: item?.proyecto ?? "—",
        zona:     item?.zona ?? "—",
        archivo:  item?.archivo ?? null
      };
    }
  }

};

/* ============================================================
   🔧 FUNCIÓN AUXILIAR: obtener módulo activo
   (Se usará desde app.js)
   ============================================================ */
export function obtenerModulo(id) {
  return MODULOS[id.toUpperCase()] ?? null;
}
