/**
 * Logo embutido (opcional).
 * Cole aqui o PNG da sua organização em base64 (versão para FUNDO ESCURO, idealmente < 70 KB):
 *   base64 -w0 logo.png        (Linux)      |   [Convert]::ToBase64String([IO.File]::ReadAllBytes("logo.png"))   (PowerShell)
 * Sem logo, o portal exibe um escudo neutro. Alternativa sem editar código: Config > LOGO_FILE_ID (arquivo no Drive).
 */
const LOGO_PADRAO = {
  mime: 'image/png',
  b64: '',
};
