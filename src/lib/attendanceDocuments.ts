import { readSheet } from "read-excel-file/browser";
import {
  CompletionRow,
  normalizeText,
  NormalizedCompletion,
  safeFileSegment,
} from "./workflow";
import {
  buildCompletionDocumentRows,
  CompletionDocumentForm,
  CompletionDocumentRow,
  createCompletionDocumentWorkbook,
} from "./completionDocument";

export type AttendanceBaseForm = CompletionDocumentForm & {
  trainingDate: string;
  startTime: string;
  endTime: string;
  period1Label: string;
  period2Label: string;
  instructorName: string;
};

export type CaptureMode = "camera" | "chat";

export type CaptureEvidenceImage = {
  name: string;
  dataUrl: string;
};

export type CaptureEvidenceRow = {
  id: string;
  period: 1 | 2;
  mode: CaptureMode;
  cameraImage: CaptureEvidenceImage | null;
  chatImages: CaptureEvidenceImage[];
};

export type AttendancePerson = {
  sequence: string;
  name: string;
  niceNumber: string;
  schoolName: string;
  source: NormalizedCompletion;
};

export type CaptureAttendanceRow = AttendancePerson & {
  period1: "O" | "X";
  period2: "O" | "X";
  result: "인정" | "미인정";
};

export type ZoomAttendanceRow = AttendancePerson & {
  entryTime: string;
  exitTime: string;
  rawMinutes: number;
  effectiveMinutes: number;
  result: "인정" | "미인정";
  warning?: string;
};

export type SummaryAttendanceRow = AttendancePerson & {
  result1: "인정" | "미인정";
  result2: "인정" | "미인정";
  result3: "이수" | "미이수";
};

export type EvaluationSummary = {
  respondentCount: number;
  averages: string[];
  opinion1: string;
  opinion2: string;
};

export type AttendanceDocumentBundle = {
  captureRows: CaptureAttendanceRow[];
  zoomRows: ZoomAttendanceRow[];
  summaryRows: SummaryAttendanceRow[];
  completionRows: CompletionDocumentRow[];
  evaluation: EvaluationSummary | null;
  completionCount: number;
  incompleteCount: number;
};

const HWPX_TEMPLATES = {
  capture: "/templates/attendance_capture.hwpx",
  zoom: "/templates/attendance_time.hwpx",
  summary: "/templates/attendance_summary.hwpx",
  evaluation: "/templates/evaluation_summary.hwpx",
} as const;

type HwpxFiles = Record<string, Uint8Array>;

type EmbeddedImageRef = {
  refId: string;
  picId: number;
  name: string;
  path: string;
};

type EmbeddedEvidenceRow = {
  period: 1 | 2;
  mode: CaptureMode;
  sequence: number;
  cameraImage: EmbeddedImageRef | null;
  chatImages: EmbeddedImageRef[];
};

export function buildAttendancePeople(rows: NormalizedCompletion[]): AttendancePerson[] {
  return [...rows]
    .sort((a, b) => a.name.localeCompare(b.name, "ko-KR"))
    .map((row, index) => ({
      sequence: String(index + 1),
      name: row.name,
      niceNumber: row.niceNumber,
      schoolName: buildSchoolName(row.region, row.school),
      source: row,
    }));
}

export function buildDefaultCaptureRows(people: AttendancePerson[]): CaptureAttendanceRow[] {
  return people.map((person) => ({
    ...person,
    period1: "O",
    period2: "O",
    result: "인정",
  }));
}

export function updateCaptureResult(row: CaptureAttendanceRow): CaptureAttendanceRow {
  return {
    ...row,
    result: row.period1 === "O" && row.period2 === "O" ? "인정" : "미인정",
  };
}

export async function parseZoomAttendanceWorkbook(file: File, people: AttendancePerson[], form: AttendanceBaseForm): Promise<ZoomAttendanceRow[]> {
  const rawRows = await readSheet(file);
  const records = rawRows.slice(1).map((row) => ({
    zoomName: normalizeText(row[16]),
    entry: normalizeTime(row[18]),
    exit: normalizeTime(row[19]),
    minutes: Number.parseFloat(normalizeText(row[20])) || 0,
  }));
  const start = parseMinutes(form.startTime);
  const end = parseMinutes(form.endTime);

  return people.map((person) => {
    const matches = records.filter((record) => isZoomNameMatch(record.zoomName, person.name));
    const entryMinutes = matches.map((match) => parseMinutes(match.entry)).filter((value) => value != null) as number[];
    const exitMinutes = matches.map((match) => parseMinutes(match.exit)).filter((value) => value != null) as number[];
    const rawMinutes = Math.round(matches.reduce((sum, match) => sum + match.minutes, 0));
    const clippedMinutes = matches.reduce((sum, match) => {
      const entry = parseMinutes(match.entry);
      const exit = parseMinutes(match.exit);
      if (entry == null || exit == null || start == null || end == null) return sum;
      return sum + Math.max(0, Math.min(exit, end) - Math.max(entry, start));
    }, 0);
    const effectiveMinutes = Math.max(0, Math.round(clippedMinutes - 10));
    const result = effectiveMinutes >= 80 ? "인정" : "미인정";
    return {
      ...person,
      entryTime: entryMinutes.length ? formatMinutes(Math.min(...entryMinutes)) : "",
      exitTime: exitMinutes.length ? formatMinutes(Math.max(...exitMinutes)) : "",
      rawMinutes,
      effectiveMinutes,
      result,
      warning: !matches.length ? "줌 접속기록 없음" : rawMinutes < 80 ? "U열 기간 합계 80분 미만 확인 필요" : undefined,
    };
  });
}

export function buildSummaryRows(captureRows: CaptureAttendanceRow[], zoomRows: ZoomAttendanceRow[]): SummaryAttendanceRow[] {
  const zoomByName = new Map(zoomRows.map((row) => [row.name, row]));
  return captureRows.map((row) => {
    const zoom = zoomByName.get(row.name);
    const result2 = zoom?.result ?? "미인정";
    return {
      ...row,
      result1: row.result,
      result2,
      result3: row.result === "인정" && result2 === "인정" ? "이수" : "미이수",
    };
  });
}

export function countSummary(rows: SummaryAttendanceRow[]) {
  const completionCount = rows.filter((row) => row.result3 === "이수").length;
  return { completionCount, incompleteCount: rows.length - completionCount };
}

export function parseEvaluationRows(rawRows: unknown[][]): EvaluationSummary {
  if (rawRows.length < 2) {
    return { respondentCount: 0, averages: Array(11).fill(""), opinion1: "", opinion2: "" };
  }
  const last = rawRows[rawRows.length - 1] ?? [];
  const dataRows = rawRows.slice(1, -1);
  return {
    respondentCount: dataRows.length,
    averages: Array.from({ length: 11 }, (_, index) => normalizeText(last[index + 1])),
    opinion1: dataRows.map((row) => normalizeText(row[12])).filter(Boolean).join("\n"),
    opinion2: dataRows.map((row) => normalizeText(row[13])).filter(Boolean).join("\n"),
  };
}

export function buildCompletionRowsWithZoomMinutes(
  people: AttendancePerson[],
  form: CompletionDocumentForm,
): CompletionDocumentRow[] {
  return buildCompletionDocumentRows(people.map((person) => person.source), form);
}

export async function createCaptureHwpx(form: AttendanceBaseForm, rows: CaptureAttendanceRow[], evidenceRows: CaptureEvidenceRow[]): Promise<Blob> {
  return createHwpx("capture", form, rows, {
    "시작시간": form.startTime,
    "종료시간": form.endTime,
    "1교시": form.period1Label,
    "2교시": form.period2Label,
  }, evidenceRows);
}

export async function createZoomHwpx(form: AttendanceBaseForm, rows: ZoomAttendanceRow[]): Promise<Blob> {
  const replacements = {
    "시작시간": form.startTime,
    "종료시간": form.endTime,
    "시작마감시간": addMinutesText(form.startTime, 20),
    "종료마감시간": addMinutesText(form.endTime, 10),
  };
  return createHwpx("zoom", form, rows, replacements);
}

export async function createSummaryHwpx(form: AttendanceBaseForm, rows: SummaryAttendanceRow[]): Promise<Blob> {
  const counts = countSummary(rows);
  return createHwpx("summary", form, rows, {
    "시작시간": form.startTime,
    "종료시간": form.endTime,
    "이수자수": String(counts.completionCount),
    "미이수자수": String(counts.incompleteCount),
  });
}

export async function createEvaluationHwpx(
  form: AttendanceBaseForm,
  evaluation: EvaluationSummary,
  completionCount: number,
  incompleteCount: number,
): Promise<Blob> {
  const replacements: Record<string, string> = {
    "등록자": String(completionCount + incompleteCount),
    "이수자": String(completionCount),
    "미이수자": String(incompleteCount),
    "강사명": form.instructorName,
    "종합의견_1": evaluation.opinion1,
    "종합의견_2": evaluation.opinion2,
  };
  evaluation.averages.forEach((value, index) => {
    replacements[`응답_${index + 1}`] = value;
  });
  return createHwpx("evaluation", form, [], replacements);
}

export async function createCompletionXlsx(people: AttendancePerson[], form: CompletionDocumentForm, zoomRows: ZoomAttendanceRow[]): Promise<Blob> {
  void zoomRows;
  return createCompletionDocumentWorkbook(buildCompletionRowsWithZoomMinutes(people, form));
}

export function buildAttendanceFilename(form: AttendanceBaseForm, label: string, extension: "hwpx" | "xlsx") {
  void form;
  return `${safeFileSegment(label) || label}.${extension}`;
}

async function createHwpx(
  template: keyof typeof HWPX_TEMPLATES,
  form: AttendanceBaseForm,
  rows: Array<CaptureAttendanceRow | ZoomAttendanceRow | SummaryAttendanceRow>,
  replacements: Record<string, string> = {},
  evidenceRows: CaptureEvidenceRow[] = [],
): Promise<Blob> {
  const response = await fetch(HWPX_TEMPLATES[template]);
  if (!response.ok) throw new Error(`HWPX 템플릿을 불러오지 못했습니다. HTTP ${response.status}`);
  const files = await unzipArchive(new Uint8Array(await response.arrayBuffer()));
  const embeddedEvidenceRows = template === "capture" ? embedEvidenceImages(files, evidenceRows) : [];
  let section = decodeUtf8(files["Contents/section0.xml"]);
  section = renderBlock(section, "연수생", rows.map(rowToTemplateValues));
  section = renderEvidence(section, embeddedEvidenceRows, form);
  section = replaceTokens(section, {
    ...baseTemplateValues(form),
    ...replacements,
  });
  files["Contents/section0.xml"] = encodeUtf8(section);
  updatePreviewText(files, section);
  return new Blob([createZip(files)], { type: "application/x-hwpml-package" });
}

function rowToTemplateValues(row: CaptureAttendanceRow | ZoomAttendanceRow | SummaryAttendanceRow): Record<string, string> {
  return {
    "연번": row.sequence,
    "성명": row.name,
    "나이스번호": row.niceNumber,
    "학교명": row.schoolName,
    "1교시": "period1" in row ? row.period1 : "",
    "2교시": "period2" in row ? row.period2 : "",
    "입장시간": "entryTime" in row ? row.entryTime : "",
    "퇴장시간": "exitTime" in row ? row.exitTime : "",
    "결과": "result" in row ? row.result : "",
    "결과_1": "result1" in row ? row.result1 : "",
    "결과_2": "result2" in row ? row.result2 : "",
    "결과_3": "result3" in row ? row.result3 : "",
  };
}

function renderEvidence(section: string, rows: EmbeddedEvidenceRow[], form: AttendanceBaseForm): string {
  const open = "{{#증빙묶음}}";
  const close = "{{/증빙묶음}}";
  const start = section.indexOf(open);
  const end = section.indexOf(close);
  if (start < 0 || end < 0 || end < start) return section;

  const rowStart = section.lastIndexOf("<hp:tr", start);
  const closeRowEnd = section.indexOf("</hp:tr>", end);
  if (rowStart < 0 || closeRowEnd < 0) return section;
  const blockEnd = closeRowEnd + "</hp:tr>".length;

  // The block holds two template rows: a wide camera row whose first cell is the
  // 교시 label (vertically merged), and a chat row with four narrow image cells.
  const templateRows = section.slice(rowStart, blockEnd).match(/<hp:tr\b[\s\S]*?<\/hp:tr>/g) ?? [];
  const cameraRow = templateRows[0];
  const chatRow = templateRows[1];
  if (!cameraRow || !chatRow) return section;
  const cameraCells = cameraRow.replace(open, "").match(/<hp:tc\b[\s\S]*?<\/hp:tc>/g) ?? [];
  const chatCells = chatRow.replace(close, "").match(/<hp:tc\b[\s\S]*?<\/hp:tc>/g) ?? [];
  const labelCell = cameraCells[0];
  const cameraContentCells = cameraCells.slice(1);
  if (!labelCell || !cameraContentCells.length || !chatCells.length) return section;

  const rowHeight = readCellHeight(cameraContentCells[0] ?? "") || 7456;
  const baseRow = Number(cameraRow.match(/rowAddr="(\d+)"/)?.[1] ?? "0");

  const dataRows: EmbeddedEvidenceRow[] = rows.length ? rows : [
    { period: 1, mode: "camera", sequence: 1, cameraImage: null, chatImages: [] },
    { period: 2, mode: "camera", sequence: 1, cameraImage: null, chatImages: [] },
  ];
  const periods = [...new Set(dataRows.map((row) => row.period))].sort((a, b) => a - b);

  const physicalRows: string[] = [];
  periods.forEach((period) => {
    const group = dataRows.filter((row) => row.period === period);
    const periodLabel = period === 1 ? form.period1Label : form.period2Label;
    group.forEach((row, index) => {
      const cells: string[] = [];
      // The 교시 label cell appears once per period and spans the whole group.
      if (index === 0) cells.push(renderEvidenceLabelCell(labelCell, periodLabel, group.length, rowHeight));
      const contentCells = row.mode === "camera" ? cameraContentCells : chatCells;
      const values = evidenceRowValues(row);
      contentCells.forEach((cell) => cells.push(replaceTokens(cell, values)));
      const rowAddr = baseRow + physicalRows.length;
      physicalRows.push(`<hp:tr>${cells.map((cell) => setCellRowAddr(cell, rowAddr)).join("")}</hp:tr>`);
    });
  });

  const rowDelta = physicalRows.length - 2;
  const rendered = physicalRows.join("");
  return `${updateLastTableRowCount(section.slice(0, rowStart), rowDelta)}${rendered}${shiftFollowingTableRows(section.slice(blockEnd), rowDelta)}`;
}

function renderEvidenceLabelCell(template: string, label: string, span: number, rowHeight: number): string {
  return replaceTokens(template, { "교시라벨": label })
    .replace(/(<hp:cellSpan colSpan="\d+" rowSpan=)"\d+"/, `$1"${span}"`)
    .replace(/(<hp:cellSz width="\d+" height=)"\d+"/, `$1"${rowHeight * span}"`);
}

function evidenceRowValues(row: EmbeddedEvidenceRow): Record<string, string> {
  return {
    "증빙번호": `${row.period}-${row.sequence}`,
    "캠화면이미지": row.mode === "camera" && row.cameraImage ? renderImageXml(row.cameraImage, "wide") : "",
    "채팅화면_1": row.mode === "chat" && row.chatImages[0] ? renderImageXml(row.chatImages[0], "narrow") : "",
    "채팅화면_2": row.mode === "chat" && row.chatImages[1] ? renderImageXml(row.chatImages[1], "narrow") : "",
    "채팅화면_3": row.mode === "chat" && row.chatImages[2] ? renderImageXml(row.chatImages[2], "narrow") : "",
    "채팅화면_4": row.mode === "chat" && row.chatImages[3] ? renderImageXml(row.chatImages[3], "narrow") : "",
  };
}

function readCellHeight(cell: string): number {
  return Number(cell.match(/<hp:cellSz width="\d+" height="(\d+)"/)?.[1] ?? "0");
}

function setCellRowAddr(cell: string, rowAddr: number): string {
  return cell.replace(/rowAddr="\d+"/, `rowAddr="${rowAddr}"`);
}

function renderBlock(section: string, name: string, rows: Record<string, string>[]): string {
  const open = `{{#${name}}}`;
  const close = `{{/${name}}}`;
  const start = section.indexOf(open);
  const end = section.indexOf(close);
  if (start < 0 || end < 0 || end < start) return section;

  const rowStart = section.lastIndexOf("<hp:tr", start);
  const rowEnd = section.indexOf("</hp:tr>", end);
  if (rowStart >= 0 && rowEnd >= 0) {
    const blockEnd = rowEnd + "</hp:tr>".length;
    const block = section.slice(rowStart, blockEnd);
    const rowCount = countOccurrences(block, "<hp:tr");
    const rowDelta = Math.max(0, rows.length - 1) * rowCount;
    const rendered = rows.map((row, index) =>
      shiftTableRowAddresses(replaceTokens(block, row).replace(open, "").replace(close, ""), index * rowCount),
    ).join("");
    return `${updateLastTableRowCount(section.slice(0, rowStart), rowDelta)}${rendered}${shiftFollowingTableRows(section.slice(blockEnd), rowDelta)}`;
  }

  const blockEnd = end + close.length;
  const block = section.slice(start, blockEnd);
  const rendered = rows.map((row) => replaceTokens(block, row).replace(open, "").replace(close, "")).join("");
  return `${section.slice(0, start)}${rendered}${section.slice(blockEnd)}`;
}

function countOccurrences(value: string, pattern: string): number {
  return value.split(pattern).length - 1;
}

function shiftTableRowAddresses(value: string, offset: number): string {
  if (!offset) return value;
  return value.replace(/rowAddr="(\d+)"/g, (_, rowAddress: string) => `rowAddr="${Number(rowAddress) + offset}"`);
}

function shiftFollowingTableRows(suffix: string, rowDelta: number): string {
  if (!rowDelta) return suffix;
  const tableEnd = suffix.indexOf("</hp:tbl>");
  if (tableEnd < 0) return suffix;
  const tableTailEnd = tableEnd + "</hp:tbl>".length;
  return `${shiftTableRowAddresses(suffix.slice(0, tableTailEnd), rowDelta)}${suffix.slice(tableTailEnd)}`;
}

function updateLastTableRowCount(prefix: string, rowDelta: number): string {
  if (!rowDelta) return prefix;
  const tableStart = prefix.lastIndexOf("<hp:tbl");
  if (tableStart < 0) return prefix;
  const before = prefix.slice(0, tableStart);
  const table = prefix.slice(tableStart);
  return `${before}${table.replace(/rowCnt="(\d+)"/, (_, count: string) => `rowCnt="${Number(count) + rowDelta}"`)}`;
}

function replaceTokens(value: string, replacements: Record<string, string>): string {
  return value.replace(/<hp:t>\{\{([^}]+)\}\}<\/hp:t>/g, (match, key: string) => {
    const replacement = replacements[key.trim()] ?? "";
    return replacement.startsWith("<hp:pic ") ? replacement : match;
  }).replace(/\{\{([^}]+)\}\}/g, (_, key: string) => escapeXml(replacements[key.trim()] ?? ""));
}

function embedEvidenceImages(files: HwpxFiles, rows: CaptureEvidenceRow[]): EmbeddedEvidenceRow[] {
  const meaningfulRows = rows.filter((row) =>
    row.mode === "camera" ? Boolean(row.cameraImage) : row.chatImages.length > 0,
  );
  let imageIndex = 0;
  const sequenceByPeriod = new Map<1 | 2, number>();
  const embeddedRows = meaningfulRows.map((row) => {
    const nextSequence = (sequenceByPeriod.get(row.period) ?? 0) + 1;
    sequenceByPeriod.set(row.period, nextSequence);
    const embed = (image: CaptureEvidenceImage): EmbeddedImageRef => {
      imageIndex += 1;
      const extension = imageExtension(image.dataUrl, image.name);
      const refId = `image${imageIndex}`;
      const path = `BinData/${refId}.${extension}`;
      files[path] = dataUrlToBytes(image.dataUrl);
      return {
        refId,
        picId: 1000 + imageIndex,
        name: image.name,
        path,
      };
    };
    return {
      period: row.period,
      mode: row.mode,
      sequence: nextSequence,
      cameraImage: row.mode === "camera" && row.cameraImage ? embed(row.cameraImage) : null,
      chatImages: row.mode === "chat" ? row.chatImages.slice(0, 4).map(embed) : [],
    };
  });

  const refs = embeddedRows.flatMap((row) => [row.cameraImage, ...row.chatImages].filter((image): image is EmbeddedImageRef => Boolean(image)));
  if (!refs.length) return embeddedRows;

  const headerPath = "Contents/header.xml";
  let header = decodeUtf8(files[headerPath]);
  const binDataList = `<hh:binDataList itemCnt="${refs.length}">${refs
    .map((ref) => `<hh:binData id="${ref.refId}" type="EMBEDDING" embedding="${ref.path}" compression="0"/>`)
    .join("")}</hh:binDataList>`;
  if (header.includes("<hh:binDataList")) {
    header = header.replace(/<hh:binDataList[\s\S]*?<\/hh:binDataList>/, binDataList);
  } else {
    header = header.replace("</hh:refList>", `${binDataList}</hh:refList>`);
  }
  files[headerPath] = encodeUtf8(header);

  const contentPath = "Contents/content.hpf";
  if (files[contentPath]) {
    let content = decodeUtf8(files[contentPath]);
    const items = refs.map((ref) => `<opf:item id="${ref.refId}" href="${ref.path}" media-type="${imageMime(ref.path)}" isEmbeded="1"/>`).join("");
    content = content.replace("</opf:manifest>", `${items}</opf:manifest>`);
    files[contentPath] = encodeUtf8(content);
  }
  return embeddedRows;
}

function renderImageXml(image: EmbeddedImageRef, size: "wide" | "narrow"): string {
  const width = size === "wide" ? 30000 : 7600;
  const height = 6500;
  const identityMatrix = `e1="1.000000" e2="0.000000" e3="0.000000" e4="0.000000" e5="1.000000" e6="0.000000"`;
  return (
    `<hp:pic id="${image.picId}" zOrder="0" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None">` +
    `<hp:offset x="0" y="0"/>` +
    `<hp:orgSz width="${width}" height="${height}"/>` +
    `<hp:curSz width="${width}" height="${height}"/>` +
    `<hp:flip horizontal="0" vertical="0"/>` +
    `<hp:rotationInfo angle="0" centerX="0" centerY="0" rotateImage="1"/>` +
    `<hp:renderingInfo>` +
      `<hc:transMatrix ${identityMatrix}/>` +
      `<hc:scaMatrix ${identityMatrix}/>` +
      `<hc:rotMatrix ${identityMatrix}/>` +
    `</hp:renderingInfo>` +
    `<hp:sz width="${width}" height="${height}" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" protect="0"/>` +
    `<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
    `<hp:outMargin left="0" right="0" top="0" bottom="0"/>` +
    `<hp:inMargin left="0" right="0" top="0" bottom="0"/>` +
    `<hp:imgDim dimwidth="${width}" dimheight="${height}"/>` +
    `<hp:imgRect>` +
      `<hc:pt0 x="0" y="0"/>` +
      `<hc:pt1 x="${width}" y="0"/>` +
      `<hc:pt2 x="${width}" y="${height}"/>` +
      `<hc:pt3 x="0" y="${height}"/>` +
    `</hp:imgRect>` +
    `<hp:imgClip left="0" right="0" top="0" bottom="0"/>` +
    `<hp:effects/>` +
    `<hc:img binaryItemIDRef="${image.refId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>` +
    `</hp:pic><hp:t/>`
  );
}

function updatePreviewText(files: HwpxFiles, section: string) {
  const path = "Preview/PrvText.txt";
  if (!files[path]) return;
  const text = Array.from(section.matchAll(/<hp:t>([\s\S]*?)<\/hp:t>/g))
    .map((match) => decodeXmlText(match[1]))
    .join("\n")
    .replace(/\{\{[^}]+\}\}/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  files[path] = encodeUtf8(text);
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'");
}

function baseTemplateValues(form: AttendanceBaseForm): Record<string, string> {
  const date = splitDate(form.trainingDate);
  return {
    "연수명": form.trainingName,
    "기관명": form.institute,
    "연수일자_년": date.year,
    "연수일자_월": date.month,
    "연수일자_일": date.day,
  };
}

function splitDate(value: string) {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return { year: "", month: "", day: "" };
  return { year: match[1], month: String(Number(match[2])), day: String(Number(match[3])) };
}

function buildSchoolName(region: string, school: string): string {
  const cleanRegion = normalizeText(region);
  const cleanSchool = normalizeText(school);
  if (!cleanRegion) return cleanSchool;
  if (!cleanSchool) return cleanRegion;
  if (cleanSchool.startsWith(cleanRegion)) return cleanSchool;
  return `${cleanRegion} ${cleanSchool}`;
}

function normalizeTime(value: unknown): string {
  const text = normalizeText(value);
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  return `${match[1].padStart(2, "0")}:${match[2]}`;
}

function parseMinutes(value: string): number | null {
  const match = normalizeTime(value).match(/^(\d{2}):(\d{2})$/);
  if (!match) return null;
  return Number(match[1]) * 60 + Number(match[2]);
}

function formatMinutes(value: number): string {
  const normalized = ((value % 1440) + 1440) % 1440;
  return `${String(Math.floor(normalized / 60)).padStart(2, "0")}:${String(normalized % 60).padStart(2, "0")}`;
}

function addMinutesText(value: string, amount: number): string {
  const minutes = parseMinutes(value);
  return minutes == null ? "" : formatMinutes(minutes + amount);
}

function isZoomNameMatch(zoomName: string, name: string): boolean {
  const compact = zoomName.replace(/\s+/g, "");
  return Boolean(name && compact.includes(name.replace(/\s+/g, "")));
}

function escapeXml(value: string): string {
  return normalizeText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function dataUrlToBytes(dataUrl: string): Uint8Array {
  const base64 = dataUrl.split(",")[1] ?? "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function imageExtension(dataUrl: string, name: string): "png" | "jpg" | "jpeg" | "webp" {
  const mime = dataUrl.match(/^data:image\/([^;,]+)/)?.[1]?.toLowerCase();
  if (mime === "jpeg" || mime === "jpg" || mime === "webp" || mime === "png") return mime;
  const extension = name.split(".").pop()?.toLowerCase();
  if (extension === "jpeg" || extension === "jpg" || extension === "webp" || extension === "png") return extension;
  return "png";
}

function imageMime(path: string): string {
  if (path.endsWith(".jpg") || path.endsWith(".jpeg")) return "image/jpeg";
  if (path.endsWith(".webp")) return "image/webp";
  return "image/png";
}

async function unzipArchive(bytes: Uint8Array): Promise<HwpxFiles> {
  const files: HwpxFiles = {};
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
    files[name] = method === 0 ? compressed : await inflateRaw(compressed);
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
  if (!value) throw new Error("HWPX 템플릿 본문을 찾지 못했습니다.");
  return new TextDecoder().decode(value);
}

function createZip(files: HwpxFiles): Uint8Array {
  const encoder = new TextEncoder();
  const records = Object.entries(files).map(([name, data]) => ({ name, nameBytes: encoder.encode(name), data }));
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;
  records.forEach((file) => {
    const crc = crc32(file.data);
    const local = new Uint8Array(30 + file.nameBytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
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
  endView.setUint16(8, records.length, true);
  endView.setUint16(10, records.length, true);
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
