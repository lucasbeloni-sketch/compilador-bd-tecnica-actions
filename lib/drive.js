const { google } = require("googleapis");
const { withRetry } = require("./google");

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

// Lista os .xlsx (nao-lixeira) da pasta. supportsAllDrives + includeItemsFromAllDrives
// sao OBRIGATORIOS porque a pasta vive num Shared Drive ("Drives compartilhados").
// Sem eles a API retorna lista vazia silenciosamente.
async function listXlsxInFolder(drive, folderId) {
  const files = [];
  let pageToken;
  do {
    const resp = await withRetry(
      () =>
        drive.files.list({
          q: `'${folderId}' in parents and mimeType = '${XLSX_MIME}' and trashed = false`,
          fields: "nextPageToken, files(id, name, modifiedTime)",
          pageSize: 1000,
          orderBy: "name",
          supportsAllDrives: true,
          includeItemsFromAllDrives: true,
          corpora: "allDrives",
          spaces: "drive",
          pageToken,
        }),
      { label: "drive.files.list" }
    );
    files.push(...(resp.data.files || []));
    pageToken = resp.data.nextPageToken;
  } while (pageToken);
  return files;
}

// Baixa o conteudo bruto do arquivo como Buffer (pra passar pro SheetJS).
async function downloadFile(drive, fileId) {
  const resp = await withRetry(
    () =>
      drive.files.get(
        { fileId, alt: "media", supportsAllDrives: true },
        { responseType: "arraybuffer" }
      ),
    { label: `drive.download-${fileId.slice(0, 8)}` }
  );
  return Buffer.from(resp.data);
}

module.exports = { listXlsxInFolder, downloadFile, XLSX_MIME };
