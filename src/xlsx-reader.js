const textDecoder = new TextDecoder("utf-8");
const MAX_ENTRY_SIZE = 25 * 1024 * 1024;

function findEndOfCentralDirectory(view) {
  const minimumOffset = Math.max(0, view.byteLength - 65_557);
  for (let offset = view.byteLength - 22; offset >= minimumOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("O arquivo não é uma planilha XLSX válida.");
}

function readZipEntries(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  const endOffset = findEndOfCentralDirectory(view);
  const entriesCount = view.getUint16(endOffset + 10, true);
  let offset = view.getUint32(endOffset + 16, true);
  const entries = new Map();

  for (let index = 0; index < entriesCount; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error("A estrutura interna da planilha está corrompida.");
    }

    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const nameBytes = new Uint8Array(arrayBuffer, offset + 46, nameLength);
    const name = textDecoder.decode(nameBytes);

    if (flags & 0x1) throw new Error("Planilhas protegidas por senha não são suportadas.");
    if (uncompressedSize > MAX_ENTRY_SIZE) throw new Error("A planilha é grande demais para importação no navegador.");

    entries.set(name, { method, compressedSize, localOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

async function extractEntry(arrayBuffer, entry) {
  const view = new DataView(arrayBuffer);
  if (view.getUint32(entry.localOffset, true) !== 0x04034b50) {
    throw new Error("Não foi possível ler uma parte da planilha.");
  }
  const nameLength = view.getUint16(entry.localOffset + 26, true);
  const extraLength = view.getUint16(entry.localOffset + 28, true);
  const dataOffset = entry.localOffset + 30 + nameLength + extraLength;
  const compressed = new Uint8Array(arrayBuffer, dataOffset, entry.compressedSize);

  if (entry.method === 0) return compressed;
  if (entry.method !== 8 || typeof DecompressionStream === "undefined") {
    throw new Error("Este navegador não oferece suporte à compactação usada no arquivo.");
  }

  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function readXml(arrayBuffer, entries, path, required = true) {
  const entry = entries.get(path);
  if (!entry) {
    if (!required) return null;
    throw new Error(`A planilha não contém ${path}.`);
  }
  const xmlText = textDecoder.decode(await extractEntry(arrayBuffer, entry));
  const document = new DOMParser().parseFromString(xmlText, "application/xml");
  if (document.getElementsByTagName("parsererror").length) {
    throw new Error("A planilha contém XML inválido.");
  }
  return document;
}

function resolveSheetPath(workbook, relationships) {
  const sheet = workbook.getElementsByTagName("sheet")[0];
  if (!sheet) throw new Error("A planilha não possui uma aba legível.");
  const relationshipId = sheet.getAttribute("r:id") || sheet.getAttribute("id");
  const relationship = [...relationships.getElementsByTagName("Relationship")].find(
    (item) => item.getAttribute("Id") === relationshipId,
  );
  const target = relationship?.getAttribute("Target") || "worksheets/sheet1.xml";
  if (target.startsWith("/")) return target.slice(1);
  const parts = `xl/${target}`.split("/");
  const normalized = [];
  parts.forEach((part) => {
    if (part === "..") normalized.pop();
    else if (part !== ".") normalized.push(part);
  });
  return normalized.join("/");
}

function getSharedStrings(document) {
  if (!document) return [];
  return [...document.getElementsByTagName("si")].map((item) =>
    [...item.getElementsByTagName("t")].map((text) => text.textContent || "").join(""),
  );
}

function columnIndex(reference) {
  const letters = reference.match(/[A-Z]+/i)?.[0]?.toUpperCase() || "A";
  return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function sheetToRows(document, sharedStrings) {
  const rows = [];
  for (const row of document.getElementsByTagName("row")) {
    const rowNumber = Number(row.getAttribute("r"));
    if (rowNumber < 25) continue;
    const values = [];
    for (const cell of row.getElementsByTagName("c")) {
      const index = columnIndex(cell.getAttribute("r") || "A1");
      if (index > 16) continue;
      const type = cell.getAttribute("t");
      const valueNode = cell.getElementsByTagName("v")[0];
      let value = valueNode?.textContent || "";
      if (type === "s" && value !== "") value = sharedStrings[Number(value)] ?? "";
      if (type === "inlineStr") {
        value = [...cell.getElementsByTagName("t")].map((text) => text.textContent || "").join("");
      }
      values[index] = value;
    }
    rows[rowNumber - 25] = values;
  }
  return rows;
}

export async function readFirstSheetRows(arrayBuffer) {
  if (arrayBuffer.byteLength > 30 * 1024 * 1024) {
    throw new Error("A planilha é grande demais para importação no navegador.");
  }
  const entries = readZipEntries(arrayBuffer);
  const [workbook, relationships, sharedStringsDocument] = await Promise.all([
    readXml(arrayBuffer, entries, "xl/workbook.xml"),
    readXml(arrayBuffer, entries, "xl/_rels/workbook.xml.rels"),
    readXml(arrayBuffer, entries, "xl/sharedStrings.xml", false),
  ]);
  const sheetPath = resolveSheetPath(workbook, relationships);
  const sheet = await readXml(arrayBuffer, entries, sheetPath);
  return sheetToRows(sheet, getSharedStrings(sharedStringsDocument));
}
