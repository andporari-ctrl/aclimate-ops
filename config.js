// ── Configuración global AClímate ────────────────────────────────────
// IMPORTANTE: Reemplaza GAS_URL con el URL real de tu Google Apps Script
// después de publicarlo (ver gas/Code.gs para instrucciones).

const CONFIG = {
  // URL del Google Apps Script desplegado (reemplazar después de publicar)
  GAS_URL: 'https://script.google.com/macros/s/AKfycbwcizqOIHoKEXdaqqDRHwTgo7a135kfMUPUQNxmDSfQzBz8mC5CIv4CRMMbRh5GtFrkuA/exec',

  // Equipo operativo de campo
  EQUIPO: [
    'Ericsen Aguirre',
    'Javier Quesada'
  ],

  // Nombre de la empresa para los formularios
  EMPRESA: 'AV Electromecánica S.A. — AClímate',

  // Tiempo de auto-guardado en milisegundos (3 segundos)
  AUTOSAVE_INTERVAL: 3000,

  // Calidad de compresión de fotos (0-1, donde 1 = máxima calidad)
  FOTO_QUALITY: 0.7,

  // Dimensión máxima de fotos en píxeles (ancho o alto)
  FOTO_MAX_DIM: 1200,
};
