import cashReceiptTemplateUrl from "../../xlsx_template/현금영수증 일괄발급 양식.xlsx?url";
import { NormalizedRoster, normalizeText, safeFileSegment } from "./workflow";

const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export type CashReceiptTraining = {
  trainingName: string;
  phones: string[];
};

export function buildCashReceiptTrainings(rosterRows: NormalizedRoster[]): CashReceiptTraining[] {
  const phonesByTraining = new Map<string, string[]>();
  rosterRows.forEach((row) => {
    const phone = formatCashReceiptPhone(row.phone);
    if (!phone) return;

    [row.course1, row.course2].forEach((trainingName) => {
      const cleanTrainingName = normalizeText(trainingName);
      if (!cleanTrainingName) return;
      const phones = phonesByTraining.get(cleanTrainingName) ?? [];
      phones.push(phone);
      phonesByTraining.set(cleanTrainingName, phones);
    });
  });

  return Array.from(phonesByTraining.entries())
    .map(([trainingName, phones]) => ({
      trainingName,
      phones: Array.from(new Set(phones)),
    }))
    .sort((a, b) => a.trainingName.localeCompare(b.trainingName, "ko-KR"));
}

export function buildCashReceiptFilename(trainingName: string): string {
  const cleanTrainingName = safeFileSegment(trainingName) || "현금영수증";
  const datePrefix = cleanTrainingName.match(/^(\d{6})\s+/)?.[1];
  return datePrefix
    ? `${datePrefix}_${cleanTrainingName.replace(/^\d{6}\s+/, "")}_현금영수증_일괄발급.xlsx`
    : `${cleanTrainingName}_현금영수증_일괄발급.xlsx`;
}

export async function createCashReceiptWorkbook(phones: string[]): Promise<Blob> {
  const response = await fetch(cashReceiptTemplateUrl);
  if (!response.ok) throw new Error("현금영수증 엑셀 양식을 불러오지 못했습니다.");
  const files = await unzipWorkbook(new Uint8Array(await response.arrayBuffer()));
  const sheetPath = "xl/worksheets/sheet1.xml";
  files[sheetPath] = encodeUtf8(fillCashReceiptSheet(decodeUtf8(files[sheetPath]), phones));
  return new Blob([createZip(files)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function formatCashReceiptPhone(value: string): string {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!/^010\d{8}$/.test(digits)) return "";
  return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`;
}

function fillCashReceiptSheet(sheetXml: string, phones: string[]): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(sheetXml, "application/xml");
  const sheetData = doc.getElementsByTagNameNS(SPREADSHEET_NS, "sheetData")[0];
  if (!sheetData) throw new Error("현금영수증 엑셀 양식의 시트를 읽지 못했습니다.");

  Array.from(sheetData.getElementsByTagNameNS(SPREADSHEET_NS, "row")).forEach((row) => {
    const rowNumber = Number(row.getAttribute("r") || "0");
    if (rowNumber >= 7) sheetData.removeChild(row);
  });

  phones.forEach((phone, index) => {
    const rowNumber = index + 7;
    const row = doc.createElementNS(SPREADSHEET_NS, "row");
    row.setAttribute("r", String(rowNumber));
    row.setAttribute("spans", "1:5");
    setNumberCell(doc, row, `A${rowNumber}`, "0");
    setTextCell(doc, row, `B${rowNumber}`, phone);
    setTextCell(doc, row, `C${rowNumber}`, "50,000");
    setTextCell(doc, row, `D${rowNumber}`, "50,000");
    setTextCell(doc, row, `E${rowNumber}`, "0");
    sheetData.appendChild(row);
  });

  const lastRow = Math.max(6, phones.length + 6);
  const dimension = doc.getElementsByTagNameNS(SPREADSHEET_NS, "dimension")[0];
  if (dimension) dimension.setAttribute("ref", `A1:J${lastRow}`);

  return new XMLSerializer().serializeToString(doc);
}

function setNumberCell(doc: XMLDocument, row: Element, ref: string, value: string) {
  const cell = doc.createElementNS(SPREADSHEET_NS, "c");
  cell.setAttribute("r", ref);
  const v = doc.createElementNS(SPREADSHEET_NS, "v");
  v.textContent = value;
  cell.appendChild(v);
  row.appendChild(cell);
}

function setTextCell(doc: XMLDocument, row: Element, ref: string, value: string) {
  const cell = doc.createElementNS(SPREADSHEET_NS, "c");
  cell.setAttribute("r", ref);
  cell.setAttribute("t", "inlineStr");
  const inline = doc.createElementNS(SPREADSHEET_NS, "is");
  const text = doc.createElementNS(SPREADSHEET_NS, "t");
  text.textContent = value;
  inline.appendChild(text);
  cell.appendChild(inline);
  row.appendChild(cell);
}

async function unzipWorkbook(bytes: Uint8Array): Promise<Record<string, Uint8Array>> {
  const files: Record<string, Uint8Array> = {};
  const decoder = new TextDecoder();
  let offset = 0;

  while (offset + 30 <= bytes.length) {
    const view = new DataView(bytes.buffer, bytes.byteOffset + offset);
    if (view.getUint32(0, true) !== 0x04034b50) break;
    const method = view.getUint16(8, true);
    const compressedSize = view.getUint32(18, true);
    const fileNameLength = view.getUint16(26, true);
    const extraLength = view.getUint16(28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + fileNameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    const name = decoder.decode(bytes.slice(nameStart, nameStart + fileNameLength));
    const compressed = bytes.slice(dataStart, dataEnd);
    if (method === 0) {
      files[name] = compressed;
    } else if (method === 8) {
      files[name] = await inflateRaw(compressed);
    } else {
      throw new Error("현금영수증 엑셀 양식 압축 형식을 읽지 못했습니다.");
    }
    offset = dataEnd;
  }

  return files;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function encodeUtf8(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function decodeUtf8(value: Uint8Array | undefined): string {
  if (!value) throw new Error("현금영수증 엑셀 양식에 필요한 시트가 없습니다.");
  return new TextDecoder().decode(value);
}

function createZip(files: Record<string, Uint8Array>): Uint8Array {
  const encoder = new TextEncoder();
  const fileRecords = Object.entries(files).map(([name, data]) => ({
    name,
    nameBytes: encoder.encode(name),
    data,
  }));
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;

  fileRecords.forEach((file) => {
    const crc = crc32(file.data);
    const local = new Uint8Array(30 + file.nameBytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, 0, true);
    localView.setUint16(12, 0, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, file.data.length, true);
    localView.setUint32(22, file.data.length, true);
    localView.setUint16(26, file.nameBytes.length, true);
    local.set(file.nameBytes, 30);
    chunks.push(local, file.data);

    const central = new Uint8Array(46 + file.nameBytes.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 0, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, file.data.length, true);
    centralView.setUint32(24, file.data.length, true);
    centralView.setUint16(28, file.nameBytes.length, true);
    centralView.setUint32(42, offset, true);
    central.set(file.nameBytes, 46);
    centralDirectory.push(central);
    offset += local.length + file.data.length;
  });

  const centralOffset = offset;
  const centralSize = centralDirectory.reduce((sum, chunk) => sum + chunk.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, fileRecords.length, true);
  endView.setUint16(10, fileRecords.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, centralOffset, true);
  return concatUint8Arrays([...chunks, ...centralDirectory, end]);
}

let crcTable: Uint32Array | null = null;

function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[i] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  data.forEach((byte) => {
    crc = crcTable![byte ^ (crc & 0xff)] ^ (crc >>> 8);
  });
  return (crc ^ 0xffffffff) >>> 0;
}

function concatUint8Arrays(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((sum, array) => sum + array.length, 0);
  const result = new Uint8Array(total);
  let offset = 0;
  arrays.forEach((array) => {
    result.set(array, offset);
    offset += array.length;
  });
  return result;
}
