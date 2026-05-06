// Configuración del botón "Actualizar".
// El botón dispara un GitHub Actions workflow que corre build-data.js y sube
// data.json + data.js al hosting por FTP. Cuando termina, el dashboard recarga.
//
// Para activar el botón completá los 4 valores. Dejá githubToken en '' si todavía
// no lo configuraste; el botón quedará oculto.
window.__CONFIG__ = {
  githubOwner:    '',           // ej: 'melimaiolo'
  githubRepo:     '',           // ej: 'panier-dashboard'
  githubWorkflow: 'refresh-data.yml',
  githubBranch:   'main',
  // ⚠️ Token expuesto en el navegador. Usar un fine-grained PAT con
  // permiso "Actions: Read and write" SOLO sobre este repo. Sin otros scopes.
  githubToken:    '',
};
