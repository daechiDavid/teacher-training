import { NormalizedCompletion, normalizeText, safeFileSegment } from "./workflow";

const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export type CompletionDocumentForm = {
  trainingName: string;
  institute: string;
  trainingDate: string;
  totalHours: string;
};

export type CompletionDocumentRow = {
  sequence: string;
  niceNumber: string;
  trainingName: string;
  institute: string;
  startDate: string;
  endDate: string;
  trainingCategory: string;
  educationTypeCode: string;
  totalHours: string;
  score: string;
  jobRelated: string;
  credit: string;
  completionNumber: string;
  name: string;
  birthDate: string;
  schoolName: string;
  schoolLevel: string;
  trainingClassCode: string;
  certificateNumber: string;
  mandatoryYn: string;
  mandatoryCode: string;
  educationFormatCode: string;
};

export const COMPLETION_DOCUMENT_HEADERS = [
  "연번",
  "나이스개인번호",
  "연수과정",
  "연수기관",
  "연수시작일",
  "연수종료일",
  "연수구분",
  "교육유형구분코드",
  "연수시간",
  "성적",
  "직무관련성",
  "평점학점",
  "이수번호",
  "성명",
  "생년월일",
  "학교명",
  "초/중등",
  "연수분류코드",
  "합격증번호",
  "법정의무여부",
  "법정의무코드",
  "교육형태코드",
] as const;

export function getTrainingNameOptions(rows: NormalizedCompletion[]): string[] {
  return Array.from(new Set(rows.map((row) => row.trainingName).filter(Boolean)));
}

export function buildCompletionDocumentRows(
  rows: NormalizedCompletion[],
  form: CompletionDocumentForm,
): CompletionDocumentRow[] {
  const trainingDate = compactDate(form.trainingDate);
  return rows.map((row, index) => ({
    sequence: String(index + 1),
    niceNumber: row.niceNumber,
    trainingName: normalizeText(form.trainingName),
    institute: normalizeText(form.institute),
    startDate: trainingDate,
    endDate: trainingDate,
    trainingCategory: "",
    educationTypeCode: "",
    totalHours: normalizeText(form.totalHours),
    score: "",
    jobRelated: "Y",
    credit: "1",
    completionNumber: "",
    name: row.name,
    birthDate: row.birthDate,
    schoolName: normalizeText(row.school),
    schoolLevel: row.schoolLevel.slice(0, 2),
    trainingClassCode: "",
    certificateNumber: "",
    mandatoryYn: "N",
    mandatoryCode: "00",
    educationFormatCode: "",
  }));
}

export function buildCompletionDocumentFilename(form: CompletionDocumentForm): string {
  const training = safeFileSegment(form.trainingName) || "직무연수";
  const date = safeFileSegment(form.trainingDate) || "이수자명단";
  return `${training}_${date}_직무연수_이수자_명단.xlsx`;
}

function compactDate(value: string): string {
  return normalizeText(value).replace(/\D/g, "").slice(0, 8);
}

export async function createCompletionDocumentWorkbook(rows: CompletionDocumentRow[]): Promise<Blob> {
  const templateResponse = await fetch("/completion_template.xlsx");
  if (!templateResponse.ok) {
    throw new Error("이수자 명단 양식을 불러오지 못했습니다. 다시 시도해 주세요.");
  }
  const files = await unzipWorkbook(new Uint8Array(await templateResponse.arrayBuffer()));
  const sheetXml = decodeUtf8(files["xl/worksheets/sheet1.xml"]);
  files["xl/worksheets/sheet1.xml"] = encodeUtf8(fillTemplateSheet(sheetXml, rows));
  return new Blob([createZip(files)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function fillTemplateSheet(sheetXml: string, rows: CompletionDocumentRow[]): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(sheetXml, "application/xml");
  const sheetData = doc.getElementsByTagNameNS(SPREADSHEET_NS, "sheetData")[0];
  if (!sheetData) throw new Error("이수자 명단 양식을 읽지 못했습니다. 다시 시도해 주세요.");

  const dataRowsNeeded = Math.max(rows.length, 0);
  const existingRows = Array.from(sheetData.getElementsByTagNameNS(SPREADSHEET_NS, "row"));
  const templateRow = existingRows.find((row) => row.getAttribute("r") === "2") ?? existingRows[1];
  const totalRows = Math.max(existingRows.length, dataRowsNeeded + 1);
  for (let rowNumber = 2; rowNumber <= totalRows; rowNumber += 1) {
    const rowElement = ensureRow(doc, sheetData, rowNumber, templateRow);
    const values = rows[rowNumber - 2] ? rowToValues(rows[rowNumber - 2]) : Array(COMPLETION_DOCUMENT_HEADERS.length).fill("");
    values.forEach((value, columnIndex) => {
      setCellValue(doc, rowElement, `${columnName(columnIndex + 1)}${rowNumber}`, value);
    });
  }

  const dimension = doc.getElementsByTagNameNS(SPREADSHEET_NS, "dimension")[0];
  if (dimension) dimension.setAttribute("ref", `A1:V${Math.max(1, totalRows)}`);
  const autoFilter = doc.getElementsByTagNameNS(SPREADSHEET_NS, "autoFilter")[0];
  if (autoFilter) autoFilter.setAttribute("ref", `A1:V${Math.max(1, totalRows)}`);

  return new XMLSerializer().serializeToString(doc);
}

function rowToValues(row: CompletionDocumentRow): string[] {
  return [
    row.sequence,
    row.niceNumber,
    row.trainingName,
    row.institute,
    row.startDate,
    row.endDate,
    row.trainingCategory,
    row.educationTypeCode,
    row.totalHours,
    row.score,
    row.jobRelated,
    row.credit,
    row.completionNumber,
    row.name,
    row.birthDate,
    row.schoolName,
    row.schoolLevel,
    row.trainingClassCode,
    row.certificateNumber,
    row.mandatoryYn,
    row.mandatoryCode,
    row.educationFormatCode,
  ];
}

function ensureRow(
  doc: XMLDocument,
  sheetData: Element,
  rowNumber: number,
  templateRow?: Element,
): Element {
  const existing = Array.from(sheetData.getElementsByTagNameNS(SPREADSHEET_NS, "row")).find(
    (row) => row.getAttribute("r") === String(rowNumber),
  );
  if (existing) return existing;

  const row = templateRow ? (templateRow.cloneNode(true) as Element) : doc.createElementNS(SPREADSHEET_NS, "row");
  row.setAttribute("r", String(rowNumber));
  Array.from(row.getElementsByTagNameNS(SPREADSHEET_NS, "c")).forEach((cell, index) => {
    cell.setAttribute("r", `${columnName(index + 1)}${rowNumber}`);
  });
  sheetData.appendChild(row);
  return row;
}

function setCellValue(doc: XMLDocument, rowElement: Element, ref: string, value: string) {
  const cell = ensureCell(doc, rowElement, ref);
  Array.from(cell.childNodes).forEach((child) => cell.removeChild(child));
  cell.setAttribute("t", "inlineStr");
  const inline = doc.createElementNS(SPREADSHEET_NS, "is");
  const text = doc.createElementNS(SPREADSHEET_NS, "t");
  text.textContent = value;
  inline.appendChild(text);
  cell.appendChild(inline);
}

function ensureCell(doc: XMLDocument, rowElement: Element, ref: string): Element {
  const existing = Array.from(rowElement.getElementsByTagNameNS(SPREADSHEET_NS, "c")).find(
    (cell) => cell.getAttribute("r") === ref,
  );
  if (existing) return existing;

  const cell = doc.createElementNS(SPREADSHEET_NS, "c");
  cell.setAttribute("r", ref);
  rowElement.appendChild(cell);
  return cell;
}

function columnName(index: number): string {
  let name = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    current = Math.floor((current - 1) / 26);
  }
  return name;
}

async function unzipWorkbook(bytes: Uint8Array): Promise<Record<string, Uint8Array>> {
  const files: Record<string, Uint8Array> = {};
  const decoder = new TextDecoder();
  let offset = 0;

  while (offset + 30 <= bytes.length) {
    const view = new DataView(bytes.buffer, bytes.byteOffset + offset);
    const signature = view.getUint32(0, true);
    if (signature !== 0x04034b50) break;

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
      throw new Error("이수자 명단 파일을 만들지 못했습니다. 다른 환경에서 다시 시도해 주세요.");
    }
    offset = dataEnd;
  }

  return files;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("현재 환경에서 이수자 명단 파일 생성을 지원하지 않습니다.");
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function encodeUtf8(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function decodeUtf8(value: Uint8Array | undefined): string {
  if (!value) throw new Error("이수자 명단 양식에 필요한 파일이 없습니다.");
  return new TextDecoder().decode(value);
}

function createZip(files: Record<string, string | Uint8Array>): Uint8Array {
  const encoder = new TextEncoder();
  const fileRecords = Object.entries(files).map(([name, content]) => ({
    name,
    nameBytes: encoder.encode(name),
    data: typeof content === "string" ? encoder.encode(content) : content,
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
      for (let k = 0; k < 8; k += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
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
