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
  cameraImages: CaptureEvidenceImage[];
  chatImages: CaptureEvidenceImage[];
  /** Links rows automatically created from one multi-image selection. */
  imageBatchId?: string;
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

export type ZoomChatAttendanceApplyResult = {
  rows: CaptureAttendanceRow[];
  startMatches: number;
  endMatches: number;
  anchorTime: string;
  startWindowLabel: string;
  endWindowLabel: string;
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

export type HwpxPreliminaryTrainingRow = {
  number: string;
  name: string;
  instructor?: string;
};

const HWPX_TEMPLATES = {
  capture: "/templates/attendance_capture.hwpx",
  zoom: "/templates/attendance_time.hwpx",
  summary: "/templates/attendance_summary.hwpx",
  evaluation: "/templates/evaluation_summary.hwpx",
} as const;

const ZIP_ENTRIES = Symbol("hwpxZipEntries");

type ZipEntryMeta = {
  name: string;
  method: number;
  flags: number;
  versionMadeBy: number;
  versionNeeded: number;
  modifiedTime: number;
  modifiedDate: number;
  localExtra: Uint8Array;
  centralExtra: Uint8Array;
  comment: Uint8Array;
  internalAttrs: number;
  externalAttrs: number;
};

type HwpxFiles = Record<string, Uint8Array> & {
  [ZIP_ENTRIES]?: ZipEntryMeta[];
};

type EmbeddedImageRef = {
  refId: string;
  picId: number;
  name: string;
  path: string;
  naturalWidth: number;
  naturalHeight: number;
};

type EmbeddedEvidenceRow = {
  period: 1 | 2;
  mode: CaptureMode;
  sequence: number;
  cameraImages: EmbeddedImageRef[];
  chatImages: EmbeddedImageRef[];
};

type EvidenceCellLayout = {
  colAddr: number;
  colSpan: number;
  cellWidth: number;
  imageWidth: number;
};

const EVIDENCE_IMAGE_START_COLUMN = 4;
const EVIDENCE_IMAGE_COLUMN_SPAN = 11;
const EVIDENCE_IMAGE_TOTAL_WIDTH = 32113;
const SINGLE_CAMERA_IMAGE_WIDTH = 30000;
const IMAGE_TO_CELL_WIDTH_RATIO = 7600 / 8028;

export function buildAttendancePeople(rows: NormalizedCompletion[]): AttendancePerson[] {
  return [...rows]
    .sort((a, b) => a.name.localeCompare(b.name, "ko-KR"))
    .map((row, index) => ({
      sequence: String(index + 1),
      name: row.name,
      niceNumber: row.niceNumber,
      schoolName: normalizeText(row.school),
      source: row,
    }));
}

export function buildDefaultCaptureRows(people: AttendancePerson[]): CaptureAttendanceRow[] {
  return people.map((person) => ({
    ...person,
    period1: "X",
    period2: "X",
    result: "미인정",
  }));
}

export function updateCaptureResult(row: CaptureAttendanceRow): CaptureAttendanceRow {
  return {
    ...row,
    result: row.period1 === "O" && row.period2 === "O" ? "인정" : "미인정",
  };
}

/** Derive document 1 without writing automatic O marks into chat/manual attendance. */
export function applyRecognizedZoomRowsToCaptureRows(
  captureRows: CaptureAttendanceRow[],
  zoomRows: ZoomAttendanceRow[],
): CaptureAttendanceRow[] {
  const recognizedKeys = new Set(
    zoomRows.filter((row) => row.result === "인정").map(attendancePersonKey),
  );
  return captureRows.map((row) => recognizedKeys.has(attendancePersonKey(row))
    ? updateCaptureResult({ ...row, period1: "O", period2: "O" })
    : row);
}

export async function applyZoomChatAttendanceText(
  file: File,
  rows: CaptureAttendanceRow[],
  form: AttendanceBaseForm,
): Promise<ZoomChatAttendanceApplyResult> {
  const text = await readFileAsText(file);
  const records = parseZoomChatRecords(text);
  const anchor = findZoomChatClockAnchor(records);
  if (!anchor) {
    throw new Error("보조강사가 실제 시각(HH:mm)을 입력한 채팅을 찾지 못했습니다.");
  }
  const start = parseMinutes(form.startTime);
  const end = parseMinutes(form.endTime);
  if (start == null || end == null) {
    throw new Error("연수 시작시간과 종료시간을 HH:mm 형식으로 입력하세요.");
  }
  if (!records.length) {
    throw new Error("줌 채팅기록에서 읽을 수 있는 대화 내용을 찾지 못했습니다.");
  }
  const normalizedEnd = end < start ? end + 1440 : end;
  const meetingStart = normalizeMinuteNearRange(anchor.clockMinutes - anchor.elapsedMinutes, start, normalizedEnd);
  const recordTimes = records.map((record) => normalizeMinuteNearRange(anchor.clockMinutes + (record.elapsedMinutes - anchor.elapsedMinutes), start, normalizedEnd));
  const meetingEnd = Math.max(...recordTimes);
  const startWindowEnd = start + 30;
  const endWindowStart = meetingEnd - 30;
  const startNames = new Set<string>();
  const endNames = new Set<string>();

  records.forEach((record, index) => {
    if (!record.message.includes("출석")) return;
    const actual = recordTimes[index];
    if (actual >= start && actual <= startWindowEnd) startNames.add(record.speaker);
    if (actual >= endWindowStart && actual <= meetingEnd) endNames.add(record.speaker);
  });

  let startMatches = 0;
  let endMatches = 0;
  const nextRows = rows.map((row) => {
    const period1 = [...startNames].some((speaker) => isZoomAttendanceRecordMatch(speaker, row)) ? "O" : row.period1;
    const period2 = [...endNames].some((speaker) => isZoomAttendanceRecordMatch(speaker, row)) ? "O" : row.period2;
    if (period1 === "O" && row.period1 !== "O") startMatches += 1;
    if (period2 === "O" && row.period2 !== "O") endMatches += 1;
    return updateCaptureResult({ ...row, period1, period2 });
  });

  return {
    rows: nextRows,
    startMatches,
    endMatches,
    anchorTime: formatMinutes(anchor.clockMinutes),
    startWindowLabel: `${formatMinutes(start)}~${formatMinutes(startWindowEnd)}`,
    endWindowLabel: `${formatMinutes(endWindowStart)}~${formatMinutes(meetingEnd)}`,
  };
}

export async function parseZoomAttendanceWorkbook(file: File, people: AttendancePerson[], form: AttendanceBaseForm): Promise<ZoomAttendanceRow[]> {
  const rawRows = await readZoomRows(file);
  const records = rawRows.slice(1).map((row) => ({
    zoomName: normalizeText(row[16]),
    entry: normalizeTime(row[18]),
    exit: normalizeTime(row[19]),
    minutes: Number.parseFloat(normalizeText(row[20])) || 0,
  }));
  const start = parseMinutes(form.startTime);
  const end = parseMinutes(form.endTime);
  const entryFloor = start == null ? null : start - 10;

  return people.map((person) => {
    const matches = records.filter((record) => isZoomAttendanceRecordMatch(record.zoomName, person));
    const entryMinutes = matches.map((match) => parseMinutes(match.entry)).filter((value) => value != null) as number[];
    const exitMinutes = matches.map((match) => parseMinutes(match.exit)).filter((value) => value != null) as number[];
    const rawMinutes = Math.round(matches.reduce((sum, match) => sum + match.minutes, 0));
    const intervals = matches.reduce<Array<[number, number]>>((result, match) => {
      const entry = parseMinutes(match.entry);
      const exit = parseMinutes(match.exit);
      if (entry != null && exit != null) result.push([entry, exit]);
      return result;
    }, []);
    const individualEffectiveMinutes = start == null || end == null
      ? 0
      : intervals.reduce((sum, [entry, exit]) => sum + calculateEffectiveZoomMinutes(entry, exit, start, end), 0);
    const effectiveMinutes = start == null || end == null
      ? 0
      : calculateMergedEffectiveZoomMinutes(intervals, start, end);
    const excludedOverlapMinutes = Math.max(0, individualEffectiveMinutes - effectiveMinutes);
    const result = effectiveMinutes >= 80 ? "인정" : "미인정";
    const earliestEntry = entryMinutes.length ? Math.min(...entryMinutes) : null;
    const latestExit = exitMinutes.length ? Math.max(...exitMinutes) : null;
    return {
      ...person,
      entryTime: earliestEntry == null ? "-" : formatMinutes(entryFloor == null ? earliestEntry : Math.max(earliestEntry, entryFloor)),
      exitTime: latestExit == null ? "-" : formatMinutes(latestExit),
      rawMinutes,
      effectiveMinutes,
      result,
      warning: buildZoomAttendanceWarning(matches.length, effectiveMinutes, excludedOverlapMinutes),
    };
  });
}

function buildZoomAttendanceWarning(
  matchCount: number,
  effectiveMinutes: number,
  excludedOverlapMinutes: number,
): string | undefined {
  if (matchCount === 0) return "줌 접속기록 없음";
  const warnings: string[] = [];
  if (excludedOverlapMinutes > 0) warnings.push(`중복 접속 ${excludedOverlapMinutes}분 제외`);
  if (effectiveMinutes < 80) warnings.push("인정시간 80분 미만 확인 필요");
  return warnings.length ? warnings.join(" · ") : undefined;
}

/**
 * Merge overlapping device sessions for one participant before calculating
 * attendance, so the same minute is counted at most once.
 */
export function calculateMergedEffectiveZoomMinutes(
  intervals: Array<[number, number]>,
  start: number,
  end: number,
): number {
  const normalizedEnd = end < start ? end + 1440 : end;
  const clippedIntervals = intervals.reduce<Array<[number, number]>>((result, [entry, exit]) => {
    const normalizedEntry = normalizeMinuteNearRange(entry, start, normalizedEnd);
    let normalizedExit = normalizeMinuteNearRange(exit, start, normalizedEnd);
    if (normalizedExit < normalizedEntry) normalizedExit += 1440;

    const clippedStart = Math.max(normalizedEntry, start);
    const clippedEnd = Math.min(normalizedExit, normalizedEnd);
    if (clippedEnd > clippedStart) result.push([clippedStart, clippedEnd]);
    return result;
  }, []);

  clippedIntervals.sort(([leftStart], [rightStart]) => leftStart - rightStart);
  const mergedIntervals: Array<[number, number]> = [];
  clippedIntervals.forEach(([entry, exit]) => {
    const previous = mergedIntervals[mergedIntervals.length - 1];
    if (!previous || entry > previous[1]) {
      mergedIntervals.push([entry, exit]);
      return;
    }
    previous[1] = Math.max(previous[1], exit);
  });

  return mergedIntervals.reduce(
    (sum, [entry, exit]) => sum + calculateEffectiveZoomMinutes(entry, exit, start, end),
    0,
  );
}

/**
 * Count only the part of a Zoom session inside the training interval and
 * remove the single 10-minute break between the two 50-minute periods.
 */
export function calculateEffectiveZoomMinutes(
  entry: number,
  exit: number,
  start: number,
  end: number,
): number {
  const normalizedEnd = end < start ? end + 1440 : end;
  const normalizedEntry = normalizeMinuteNearRange(entry, start, normalizedEnd);
  let normalizedExit = normalizeMinuteNearRange(exit, start, normalizedEnd);
  if (normalizedExit < normalizedEntry) normalizedExit += 1440;

  const clippedStart = Math.max(normalizedEntry, start);
  const clippedEnd = Math.min(normalizedExit, normalizedEnd);
  const clippedMinutes = Math.max(0, clippedEnd - clippedStart);
  if (clippedMinutes === 0) return 0;

  const breakStart = start + 50;
  const breakEnd = start + 60;
  const breakMinutes = Math.max(
    0,
    Math.min(clippedEnd, breakEnd) - Math.max(clippedStart, breakStart),
  );
  return Math.max(0, Math.round(clippedMinutes - breakMinutes));
}

async function readZoomRows(file: File): Promise<unknown[][]> {
  const isCsv = /\.csv$/i.test(file.name) || file.type === "text/csv";
  if (isCsv) {
    return parseCsv(await readFileAsText(file));
  }
  return readSheet(file);
}

async function readFileAsText(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  // Zoom CSV exports may be UTF-8 with a BOM; strip it so the first cell parses cleanly.
  const text = new TextDecoder("utf-8").decode(buffer);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (inQuotes) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n" || char === "\r") {
      if (char === "\r" && text[i + 1] === "\n") i += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

type ZoomChatRecord = {
  elapsedMinutes: number;
  speaker: string;
  message: string;
};

function parseZoomChatRecords(text: string): ZoomChatRecord[] {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  const records: ZoomChatRecord[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const inlineRecord = parseInlineZoomChatRecord(lines[index]);
    if (inlineRecord) {
      records.push(inlineRecord);
      continue;
    }
    const elapsedMinutes = parseElapsedMinutes(lines[index]);
    if (elapsedMinutes == null) continue;
    const speakerLine = normalizeText(lines[index + 1] ?? "");
    const speaker = parseZoomChatSpeaker(speakerLine);
    if (!speaker) continue;
    const messageLines: string[] = [];
    index += 2;
    while (index < lines.length && parseElapsedMinutes(lines[index]) == null) {
      messageLines.push(lines[index]);
      index += 1;
    }
    index -= 1;
    records.push({
      elapsedMinutes,
      speaker,
      message: normalizeText(messageLines.join("\n")),
    });
  }
  return records;
}

function parseInlineZoomChatRecord(value: string): ZoomChatRecord | null {
  const columns = value.split("\t").map(normalizeText);
  if (columns.length >= 3) {
    const elapsedMinutes = parseElapsedMinutes(columns[0]);
    const speaker = columns[1].replace(/:$/, "").trim();
    const message = columns.slice(2).join("\t").trim();
    if (elapsedMinutes != null && speaker) {
      return { elapsedMinutes, speaker, message };
    }
  }

  const match = normalizeText(value).match(/^(\d{1,2}):(\d{2}):(\d{2})\s+(.+?):\s*(.*)$/);
  if (!match) return null;
  return {
    elapsedMinutes: Number(match[1]) * 60 + Number(match[2]) + Number(match[3]) / 60,
    speaker: match[4].trim(),
    message: match[5].trim(),
  };
}

function parseZoomChatSpeaker(value: string): string {
  const clean = normalizeText(value);
  if (!clean) return "";
  if (clean.endsWith(":")) return clean.slice(0, -1).trim();
  const fromMatch = clean.match(/^(?:From|발신자)\s+(.+?)\s+(?:to|에게|님께서|→)/i);
  if (fromMatch?.[1]) return fromMatch[1].trim();
  return "";
}

function findZoomChatClockAnchor(records: ZoomChatRecord[]): { elapsedMinutes: number; clockMinutes: number } | null {
  for (const record of records) {
    const combined = `${record.speaker}\n${record.message}`;
    if (!combined.includes("보조강사")) continue;
    const match = combined.match(/(?:^|[^\d])([01]?\d|2[0-3]):([0-5]\d)(?:[^\d]|$)/);
    if (!match) continue;
    return {
      elapsedMinutes: record.elapsedMinutes,
      clockMinutes: Number(match[1]) * 60 + Number(match[2]),
    };
  }
  return null;
}

function parseElapsedMinutes(value: string): number | null {
  const match = normalizeText(value).match(/^(\d{1,2}):(\d{2}):(\d{2})(?:\s|$)/);
  if (!match) return null;
  return Number(match[1]) * 60 + Number(match[2]) + Number(match[3]) / 60;
}

export function buildSummaryRows(captureRows: CaptureAttendanceRow[], zoomRows: ZoomAttendanceRow[]): SummaryAttendanceRow[] {
  const zoomByPerson = new Map(zoomRows.map((row) => [attendancePersonKey(row), row]));
  return applyRecognizedZoomRowsToCaptureRows(captureRows, zoomRows).map((row) => {
    const zoom = zoomByPerson.get(attendancePersonKey(row));
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
  const lastDataRowIndex = findLastMeaningfulRowIndex(rawRows);
  if (lastDataRowIndex < 1) {
    return { respondentCount: 0, averages: Array(11).fill(""), opinion1: "", opinion2: "" };
  }
  const last = rawRows[lastDataRowIndex] ?? [];
  const dataRows = rawRows.slice(1, lastDataRowIndex);
  return {
    respondentCount: Math.max(0, lastDataRowIndex - 1),
    averages: Array.from({ length: 11 }, (_, index) => normalizeText(last[index + 1])),
    opinion1: dataRows.map((row) => normalizeText(row[12])).filter(Boolean).join("\n"),
    opinion2: dataRows.map((row) => normalizeText(row[13])).filter(Boolean).join("\n"),
  };
}

function findLastMeaningfulRowIndex(rows: unknown[][]): number {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if ((rows[index] ?? []).some((cell) => normalizeText(cell))) return index;
  }
  return -1;
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
    "시작시간": addMinutesText(form.startTime, -10),
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
    "평가참여자": String(evaluation.respondentCount),
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

export async function createHwpxFromTemplateUrl(
  templateUrl: string,
  replacements: Record<string, string>,
  options: { preliminaryTrainingRows?: HwpxPreliminaryTrainingRow[] } = {},
): Promise<Blob> {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error("한글 문서 양식을 불러오지 못했습니다. 다시 시도해 주세요.");
  const files = await unzipArchive(new Uint8Array(await response.arrayBuffer()));
  let section = decodeUtf8(files["Contents/section0.xml"]);
  if (options.preliminaryTrainingRows) {
    section = renderPreliminaryTrainingTable(section, options.preliminaryTrainingRows);
  }
  section = replaceTokensAndInvalidateLineSegments(section, replacements);
  files["Contents/section0.xml"] = encodeUtf8(section);
  updatePreviewText(files, section);
  return new Blob([await createZip(files)], { type: "application/x-hwpml-package" });
}

function renderPreliminaryTrainingTable(
  section: string,
  rows: HwpxPreliminaryTrainingRow[],
): string {
  if (!rows.length) throw new Error("사전 문서에 넣을 연수 정보가 없습니다.");

  const firstToken = "{{연수1_연수명}}";
  const secondToken = "{{연수2_연수명}}";
  const firstTokenIndex = section.indexOf(firstToken);
  const secondTokenIndex = section.indexOf(secondToken, firstTokenIndex + firstToken.length);
  if (firstTokenIndex < 0 || secondTokenIndex < 0) {
    throw new Error("사전 문서 양식의 연수 목록 표를 찾지 못했습니다.");
  }

  const firstRowStart = section.lastIndexOf("<hp:tr", firstTokenIndex);
  const firstRowClose = section.indexOf("</hp:tr>", firstTokenIndex);
  const secondRowStart = section.lastIndexOf("<hp:tr", secondTokenIndex);
  const secondRowClose = section.indexOf("</hp:tr>", secondTokenIndex);
  if (firstRowStart < 0 || firstRowClose < 0 || secondRowStart <= firstRowStart || secondRowClose < 0) {
    throw new Error("사전 문서 양식의 연수 목록 행을 찾지 못했습니다.");
  }

  const firstRow = section.slice(firstRowStart, firstRowClose + "</hp:tr>".length);
  const renderedRows = rows.map((row, index) => {
    const numberedRow = replacePreliminaryTrainingRowNumber(firstRow, row.number);
    return shiftTableRowAddresses(
      replaceTokensAndInvalidateLineSegments(numberedRow, {
        "연수1_연수명": row.name,
        "연수1_강사": row.instructor ?? "",
      }),
      index,
    );
  }).join("");
  const rowDelta = rows.length - 2;
  const prefix = updateLastTableRowCount(section.slice(0, firstRowStart), rowDelta);
  const suffix = section.slice(secondRowClose + "</hp:tr>".length);
  return `${prefix}${renderedRows}${suffix}`;
}

function replacePreliminaryTrainingRowNumber(row: string, number: string): string {
  return row.replace(/(<hp:t>)1(<\/hp:t>)/, `$1${escapeXml(number)}$2`);
}

async function createHwpx(
  template: keyof typeof HWPX_TEMPLATES,
  form: AttendanceBaseForm,
  rows: Array<CaptureAttendanceRow | ZoomAttendanceRow | SummaryAttendanceRow>,
  replacements: Record<string, string> = {},
  evidenceRows: CaptureEvidenceRow[] = [],
): Promise<Blob> {
  const response = await fetch(HWPX_TEMPLATES[template]);
  if (!response.ok) throw new Error("한글 문서 양식을 불러오지 못했습니다. 다시 시도해 주세요.");
  const files = await unzipArchive(new Uint8Array(await response.arrayBuffer()));
  const embeddedEvidenceRows = template === "capture" ? await embedEvidenceImages(files, evidenceRows) : [];
  let section = decodeUtf8(files["Contents/section0.xml"]);
  section = renderBlock(section, "연수생", rows.map(rowToTemplateValues));
  section = renderEvidence(section, embeddedEvidenceRows, form);
  section = replaceTokens(section, {
    ...baseTemplateValues(form),
    ...replacements,
  });
  files["Contents/section0.xml"] = encodeUtf8(section);
  updatePreviewText(files, section);
  return new Blob([await createZip(files)], { type: "application/x-hwpml-package" });
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

  const defaultRowHeight = readCellHeight(cameraContentCells[0] ?? "") || 7456;
  const baseRow = Number(cameraRow.match(/rowAddr="(\d+)"/)?.[1] ?? "0");

  const dataRows: EmbeddedEvidenceRow[] = rows.length ? rows : [
    { period: 1, mode: "camera", sequence: 1, cameraImages: [], chatImages: [] },
    { period: 2, mode: "camera", sequence: 1, cameraImages: [], chatImages: [] },
  ];
  const periods = [...new Set(dataRows.map((row) => row.period))].sort((a, b) => a - b);

  const physicalRows: string[] = [];
  periods.forEach((period) => {
    const group = dataRows.filter((row) => row.period === period);
    const periodLabel = period === 1 ? form.period1Label : form.period2Label;
    const rowHeights = group.map((row) => computeEvidenceRowHeight(row, defaultRowHeight));
    group.forEach((row, index) => {
      const rowHeight = rowHeights[index];
      const cells: string[] = [];
      // The 교시 label cell appears once per period and spans the whole group.
      if (index === 0) cells.push(renderEvidenceLabelCell(labelCell, periodLabel, group.length, rowHeights));
      const contentCells = row.mode === "camera"
        ? renderCameraEvidenceCells(cameraContentCells, chatCells, row, rowHeight)
        : chatCells.map((cell) => setCellHeight(replaceTokens(cell, evidenceRowValues(row)), rowHeight));
      cells.push(...contentCells);
      const rowAddr = baseRow + physicalRows.length;
      physicalRows.push(`<hp:tr>${cells.map((cell) => setCellRowAddr(cell, rowAddr)).join("")}</hp:tr>`);
    });
  });

  const rowDelta = physicalRows.length - 2;
  const rendered = physicalRows.join("");
  return `${updateLastTableRowCount(section.slice(0, rowStart), rowDelta)}${rendered}${shiftFollowingTableRows(section.slice(blockEnd), rowDelta)}`;
}

function renderEvidenceLabelCell(template: string, label: string, span: number, rowHeights: number[]): string {
  const totalHeight = rowHeights.reduce((sum, h) => sum + h, 0);
  return replaceTokens(template, { "교시라벨": label })
    .replace(/(<hp:cellSpan colSpan="\d+" rowSpan=)"\d+"/, `$1"${span}"`)
    .replace(/(<hp:cellSz width="\d+" height=)"\d+"/, `$1"${totalHeight}"`);
}

function evidenceRowValues(row: EmbeddedEvidenceRow): Record<string, string> {
  return {
    "증빙번호": `${row.period}-${row.sequence}`,
    "캠화면이미지": row.mode === "camera" && row.cameraImages[0] ? renderImageXml(row.cameraImages[0], SINGLE_CAMERA_IMAGE_WIDTH) : "",
    "채팅화면_1": row.mode === "chat" && row.chatImages[0] ? renderImageXml(row.chatImages[0], 7600) : "",
    "채팅화면_2": row.mode === "chat" && row.chatImages[1] ? renderImageXml(row.chatImages[1], 7600) : "",
    "채팅화면_3": row.mode === "chat" && row.chatImages[2] ? renderImageXml(row.chatImages[2], 7600) : "",
    "채팅화면_4": row.mode === "chat" && row.chatImages[3] ? renderImageXml(row.chatImages[3], 7600) : "",
  };
}

function renderCameraEvidenceCells(
  cameraContentCells: string[],
  chatCells: string[],
  row: EmbeddedEvidenceRow,
  rowHeight: number,
): string[] {
  if (row.cameraImages.length <= 1) {
    return cameraContentCells.map((cell) =>
      setCellHeight(replaceTokens(cell, evidenceRowValues(row)), rowHeight),
    );
  }

  const numberCell = cameraContentCells[0];
  const typeCell = cameraContentCells.at(-1);
  const imageCellTemplates = chatCells.slice(1, -1);
  if (!numberCell || !typeCell || imageCellTemplates.length < row.cameraImages.length) {
    return cameraContentCells.map((cell) =>
      setCellHeight(replaceTokens(cell, evidenceRowValues(row)), rowHeight),
    );
  }

  const renderedNumberCell = setCellHeight(replaceTokens(numberCell, {
    "증빙번호": `${row.period}-${row.sequence}`,
  }), rowHeight);
  const renderedImageCells = calculateEvidenceCellLayout(row.cameraImages.length).map((layout, index) => {
    const rendered = replaceTokens(imageCellTemplates[index], {
      [`채팅화면_${index + 1}`]: renderImageXml(row.cameraImages[index], layout.imageWidth),
    });
    return configureEvidenceImageCell(rendered, layout, rowHeight);
  });
  return [renderedNumberCell, ...renderedImageCells, setCellHeight(typeCell, rowHeight)];
}

export function calculateEvidenceCellLayout(imageCount: number): EvidenceCellLayout[] {
  const count = Math.max(1, Math.min(4, Math.trunc(imageCount) || 1));
  const baseSpan = Math.floor(EVIDENCE_IMAGE_COLUMN_SPAN / count);
  const spanRemainder = EVIDENCE_IMAGE_COLUMN_SPAN % count;
  const baseWidth = Math.floor(EVIDENCE_IMAGE_TOTAL_WIDTH / count);
  const widthRemainder = EVIDENCE_IMAGE_TOTAL_WIDTH % count;
  let colAddr = EVIDENCE_IMAGE_START_COLUMN;
  return Array.from({ length: count }, (_, index) => {
    const colSpan = baseSpan + (index < spanRemainder ? 1 : 0);
    const cellWidth = baseWidth + (index < widthRemainder ? 1 : 0);
    const layout = {
      colAddr,
      colSpan,
      cellWidth,
      imageWidth: count === 1
        ? SINGLE_CAMERA_IMAGE_WIDTH
        : Math.round(cellWidth * IMAGE_TO_CELL_WIDTH_RATIO),
    };
    colAddr += colSpan;
    return layout;
  });
}

function configureEvidenceImageCell(cell: string, layout: EvidenceCellLayout, height: number): string {
  return setCellHeight(cell, height)
    .replace(/(<hp:cellAddr colAddr=)"\d+"/, `$1"${layout.colAddr}"`)
    .replace(/(<hp:cellSpan colSpan=)"\d+"/, `$1"${layout.colSpan}"`)
    .replace(/(<hp:cellSz width=)"\d+"/, `$1"${layout.cellWidth}"`);
}

function readCellHeight(cell: string): number {
  return Number(cell.match(/<hp:cellSz width="\d+" height="(\d+)"/)?.[1] ?? "0");
}

function setCellHeight(cell: string, height: number): string {
  return cell.replace(/(<hp:cellSz width="\d+" height=)"\d+"/, `$1"${height}"`);
}

function computeEvidenceRowHeight(row: EmbeddedEvidenceRow, defaultHeight: number): number {
  const NARROW_WIDTH = 7600;
  const PADDING = 800;
  if (row.mode === "camera" && row.cameraImages.length > 0) {
    const [{ imageWidth }] = calculateEvidenceCellLayout(row.cameraImages.length);
    const imageHeights = row.cameraImages
      .filter((image) => image.naturalWidth > 0)
      .map((image) => Math.round(imageWidth * (image.naturalHeight / image.naturalWidth)));
    if (imageHeights.length) return Math.max(defaultHeight, Math.max(...imageHeights) + PADDING);
  }
  if (row.mode === "chat" && row.chatImages.length > 0) {
    const maxImgHeight = Math.max(...row.chatImages
      .filter((img) => img.naturalWidth > 0)
      .map((img) => Math.round(NARROW_WIDTH * (img.naturalHeight / img.naturalWidth))));
    if (maxImgHeight > 0) return Math.max(defaultHeight, maxImgHeight + PADDING);
  }
  return defaultHeight;
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
    const block = removeParagraphsContaining(section.slice(rowStart, blockEnd), [open, close]);
    const rowCount = countOccurrences(block, "<hp:tr");
    const rowDelta = Math.max(0, rows.length - 1) * rowCount;
    const rendered = rows.map((row, index) =>
      shiftTableRowAddresses(replaceTokens(block, row), index * rowCount),
    ).join("");
    return `${updateLastTableRowCount(section.slice(0, rowStart), rowDelta)}${rendered}${shiftFollowingTableRows(section.slice(blockEnd), rowDelta)}`;
  }

  const blockEnd = end + close.length;
  const block = removeParagraphsContaining(section.slice(start, blockEnd), [open, close]);
  const rendered = rows.map((row) => replaceTokens(block, row)).join("");
  return `${section.slice(0, start)}${rendered}${section.slice(blockEnd)}`;
}

function removeParagraphsContaining(value: string, markers: string[]): string {
  return markers.reduce((current, marker) => {
    let next = current;
    while (next.includes(marker)) {
      const markerIndex = next.indexOf(marker);
      const paragraphStart = next.lastIndexOf("<hp:p", markerIndex);
      const paragraphEnd = next.indexOf("</hp:p>", markerIndex);
      if (paragraphStart < 0 || paragraphEnd < 0) {
        next = next.replace(marker, "");
      } else {
        next = `${next.slice(0, paragraphStart)}${next.slice(paragraphEnd + "</hp:p>".length)}`;
      }
    }
    return next;
  }, value);
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

function replaceTokensAndInvalidateLineSegments(value: string, replacements: Record<string, string>): string {
  const withParagraphs = value.replace(/<hp:p\b[\s\S]*?<\/hp:p>/g, (paragraph) => {
    const rendered = replaceTokens(paragraph, replacements);
    if (rendered === paragraph) return paragraph;
    return rendered.replace(/<hp:linesegarray>[\s\S]*?<\/hp:linesegarray>/g, "");
  });
  return replaceTokens(withParagraphs, replacements);
}

async function embedEvidenceImages(files: HwpxFiles, rows: CaptureEvidenceRow[]): Promise<EmbeddedEvidenceRow[]> {
  const meaningfulRows = rows.filter((row) =>
    row.mode === "camera" ? row.cameraImages.length > 0 : row.chatImages.length > 0,
  );
  let imageIndex = 0;
  const sequenceByPeriod = new Map<1 | 2, number>();
  const embeddedRows: EmbeddedEvidenceRow[] = [];
  for (const row of meaningfulRows) {
    const nextSequence = (sequenceByPeriod.get(row.period) ?? 0) + 1;
    sequenceByPeriod.set(row.period, nextSequence);
    const embed = async (image: CaptureEvidenceImage): Promise<EmbeddedImageRef> => {
      imageIndex += 1;
      const extension = imageExtension(image.dataUrl, image.name);
      const refId = `image${imageIndex}`;
      const path = `BinData/${refId}.${extension}`;
      files[path] = dataUrlToBytes(image.dataUrl);
      const { width: naturalWidth, height: naturalHeight } = await getImageDimensions(image.dataUrl);
      return {
        refId,
        picId: 1000 + imageIndex,
        name: image.name,
        path,
        naturalWidth,
        naturalHeight,
      };
    };
    const cameraImages = row.mode === "camera" ? await Promise.all(row.cameraImages.slice(0, 4).map(embed)) : [];
    const chatImages = row.mode === "chat" ? await Promise.all(row.chatImages.slice(0, 4).map(embed)) : [];
    embeddedRows.push({
      period: row.period,
      mode: row.mode,
      sequence: nextSequence,
      cameraImages,
      chatImages,
    });
  }

  const refs = embeddedRows.flatMap((row) => [...row.cameraImages, ...row.chatImages]);
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

function renderImageXml(image: EmbeddedImageRef, width: number): string {
  const height = image.naturalWidth > 0 && image.naturalHeight > 0
    ? Math.round(width * (image.naturalHeight / image.naturalWidth))
    : 6500;
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

function normalizeTime(value: unknown): string {
  const text = normalizeText(value);
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  let hour = Number(match[1]);
  const minute = match[2];
  // Zoom exports record times on a 12-hour clock (e.g. "8:00 PM"); convert to 24-hour
  // so downstream minute math uses a single consistent scale. Also handles 오전/오후.
  const meridiem = text.match(/(a\.?m\.?|p\.?m\.?)/i)?.[1]?.toLowerCase() ?? (text.includes("오후") ? "pm" : text.includes("오전") ? "am" : "");
  if (meridiem.startsWith("p") && hour < 12) hour += 12;
  if (meridiem.startsWith("a") && hour === 12) hour = 0;
  return `${String(hour).padStart(2, "0")}:${minute}`;
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

function normalizeMinuteNearRange(value: number, start: number, end: number): number {
  let normalized = value;
  while (normalized < start - 720) normalized += 1440;
  while (normalized > end + 720) normalized -= 1440;
  return normalized;
}

function addMinutesText(value: string, amount: number): string {
  const minutes = parseMinutes(value);
  return minutes == null ? "" : formatMinutes(minutes + amount);
}

function isZoomNameMatch(zoomName: string, name: string): boolean {
  const compact = compactMatchText(stripZoomOriginalNameSuffix(zoomName));
  return Boolean(name && compact.includes(compactMatchText(name)));
}

function isZoomAttendanceRecordMatch(zoomName: string, person: AttendancePerson): boolean {
  const parsed = parseZoomNameAndSchool(zoomName);
  if (parsed.name) {
    if (parsed.name !== compactMatchText(person.name)) return false;
    return !parsed.school || isSchoolNameMatch(parsed.school, person.source.school);
  }

  if (!isZoomNameMatch(zoomName, person.name)) return false;
  const zoomSchool = extractZoomSchoolToken(zoomName, person.name);
  if (!zoomSchool) return true;
  return isSchoolNameMatch(zoomSchool, person.source.school);
}

function parseZoomNameAndSchool(zoomName: string): { name: string; school: string } {
  const parts = normalizeText(stripZoomOriginalNameSuffix(zoomName))
    .split(/[_/()\-\s]+/)
    .map(compactMatchText)
    .filter(Boolean);
  return parts.length >= 2 ? { name: parts[0], school: parts.slice(1).join("") } : { name: "", school: "" };
}

function extractZoomSchoolToken(zoomName: string, personName: string): string {
  const compactName = compactMatchText(personName);
  if (!compactName) return "";
  const strippedZoomName = stripZoomOriginalNameSuffix(zoomName);
  const parts = normalizeText(strippedZoomName)
    .split(/[_/()\-\s]+/)
    .map(compactMatchText)
    .filter(Boolean);
  if (parts.length >= 2 && (parts[0] === compactName || parts[0].includes(compactName) || compactName.includes(parts[0]))) {
    return parts.slice(1).join("");
  }

  const compactZoomName = compactMatchText(strippedZoomName);
  const nameIndex = compactZoomName.indexOf(compactName);
  if (nameIndex < 0) return "";
  return compactZoomName.slice(nameIndex + compactName.length);
}

function stripZoomOriginalNameSuffix(value: string): string {
  return normalizeText(value).replace(/\s+\(.*\)\s*$/, "");
}

function compactMatchText(value: string): string {
  return normalizeText(value).replace(/\s+/g, "");
}

function isSchoolNameMatch(zoomSchool: string, rosterSchool: string): boolean {
  const zoom = normalizeSchoolMatchText(zoomSchool);
  const roster = normalizeSchoolMatchText(rosterSchool);
  if (!zoom || !roster) return false;
  return roster.includes(zoom) || zoom.includes(roster);
}

function normalizeSchoolMatchText(value: string): string {
  return normalizeUniversityAttachedSchoolName(compactMatchText(value))
    // Zoom names often shorten a school-affiliated kindergarten to 병설유/병유
    // and may omit the parent elementary school's grade marker.
    .replace(/병설유치원|병설유|병유/g, "병설유치원")
    .replace(/컴퍼스|캠퍼스/g, "캠")
    .replace(/공업고등학교/g, "공고")
    .replace(/공업고/g, "공고")
    .replace(/외국어고등학교/g, "외고")
    .replace(/외국어고/g, "외고")
    .replace(/초등학교/g, "초")
    .replace(/중학교/g, "중")
    .replace(/고등학교/g, "고")
    .replace(/여자/g, "여")
    .replace(/남자/g, "남")
    .replace(/(?:초|중|고)병설유치원/g, "병설유치원")
    .replace(/학교/g, "");
}

function normalizeUniversityAttachedSchoolName(value: string): string {
  const match = value.match(/^(.+?)(?:여자대학교|대학교|여자대|대)(?:사범대학)?(?:부속|부설)(.+)$/);
  if (!match) return value;
  const university = match[1];
  const school = match[2];
  const universityAbbreviation = universityAbbreviations[university] ?? `${university.slice(0, 1)}대`;
  return `${universityAbbreviation}부${normalizeAttachedSchoolLevel(school)}`;
}

function normalizeAttachedSchoolLevel(value: string): string {
  return value
    .replace(/초등학교/g, "초")
    .replace(/중학교/g, "중")
    .replace(/고등학교/g, "고")
    .replace(/유치원/g, "유")
    .replace(/학교/g, "");
}

const universityAbbreviations: Record<string, string> = {
  한양: "한대",
  단국: "단대",
  이화: "이대",
  이화여자: "이대",
};

function attendancePersonKey(person: AttendancePerson): string {
  return person.niceNumber && person.niceNumber !== "없음"
    ? person.niceNumber
    : `${person.sequence}::${person.name}::${person.schoolName}`;
}

function escapeXml(value: string): string {
  return normalizeText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function getImageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 0, height: 0 });
    img.src = dataUrl;
  });
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
  const entries: ZipEntryMeta[] = [];
  const centralEntries = readCentralDirectory(bytes);
  const decoder = new TextDecoder();
  let offset = 0;
  while (offset + 30 <= bytes.length) {
    const view = new DataView(bytes.buffer, bytes.byteOffset + offset);
    if (view.getUint32(0, true) !== 0x04034b50) break;
    const flags = view.getUint16(6, true);
    const method = view.getUint16(8, true);
    const modifiedTime = view.getUint16(10, true);
    const modifiedDate = view.getUint16(12, true);
    const compressedSize = view.getUint32(18, true);
    const fileNameLength = view.getUint16(26, true);
    const extraLength = view.getUint16(28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + fileNameLength + extraLength;
    const dataEnd = dataStart + compressedSize;
    const name = decoder.decode(bytes.slice(nameStart, nameStart + fileNameLength));
    const localExtra = bytes.slice(nameStart + fileNameLength, dataStart);
    const compressed = bytes.slice(dataStart, dataEnd);
    const central = centralEntries.get(name);
    files[name] = method === 0 ? compressed : await inflateRaw(compressed);
    entries.push({
      name,
      method,
      flags,
      versionMadeBy: central?.versionMadeBy ?? 20,
      versionNeeded: central?.versionNeeded ?? 20,
      modifiedTime,
      modifiedDate,
      localExtra,
      centralExtra: central?.centralExtra ?? localExtra,
      comment: central?.comment ?? new Uint8Array(),
      internalAttrs: central?.internalAttrs ?? 0,
      externalAttrs: central?.externalAttrs ?? 0,
    });
    offset = dataEnd;
  }
  files[ZIP_ENTRIES] = entries;
  return files;
}

function readCentralDirectory(bytes: Uint8Array): Map<string, Omit<ZipEntryMeta, "method" | "flags" | "modifiedTime" | "modifiedDate" | "localExtra">> {
  const entries = new Map<string, Omit<ZipEntryMeta, "method" | "flags" | "modifiedTime" | "modifiedDate" | "localExtra">>();
  const decoder = new TextDecoder();
  const maxCommentLength = 0xffff;
  const eocdMinLength = 22;
  const start = Math.max(0, bytes.length - eocdMinLength - maxCommentLength);
  for (let offset = bytes.length - eocdMinLength; offset >= start; offset -= 1) {
    const view = new DataView(bytes.buffer, bytes.byteOffset + offset);
    if (view.getUint32(0, true) !== 0x06054b50) continue;
    const centralCount = view.getUint16(10, true);
    let centralOffset = view.getUint32(16, true);
    for (let index = 0; index < centralCount && centralOffset + 46 <= bytes.length; index += 1) {
      const centralView = new DataView(bytes.buffer, bytes.byteOffset + centralOffset);
      if (centralView.getUint32(0, true) !== 0x02014b50) break;
      const nameLength = centralView.getUint16(28, true);
      const extraLength = centralView.getUint16(30, true);
      const commentLength = centralView.getUint16(32, true);
      const nameStart = centralOffset + 46;
      const extraStart = nameStart + nameLength;
      const commentStart = extraStart + extraLength;
      const name = decoder.decode(bytes.slice(nameStart, nameStart + nameLength));
      entries.set(name, {
        name,
        versionMadeBy: centralView.getUint16(4, true),
        versionNeeded: centralView.getUint16(6, true),
        centralExtra: bytes.slice(extraStart, commentStart),
        comment: bytes.slice(commentStart, commentStart + commentLength),
        internalAttrs: centralView.getUint16(36, true),
        externalAttrs: centralView.getUint32(38, true),
      });
      centralOffset = commentStart + commentLength;
    }
    break;
  }
  return entries;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
  if (!data.length) return data;
  const stream = new Blob([data]).stream().pipeThrough(new CompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function encodeUtf8(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function decodeUtf8(value: Uint8Array | undefined): string {
  if (!value) throw new Error("한글 문서 양식을 읽지 못했습니다. 다시 시도해 주세요.");
  return new TextDecoder().decode(value);
}

async function createZip(files: HwpxFiles): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const known = new Set<string>();
  const records = [
    ...(files[ZIP_ENTRIES] ?? []).flatMap((meta) => {
      const data = files[meta.name];
      if (!data) return [];
      known.add(meta.name);
      return [{ ...meta, data }];
    }),
    ...Object.entries(files)
      .filter(([name]) => !known.has(name))
      .map(([name, data]) => ({
        name,
        method: name === "mimetype" ? 0 : 8,
        flags: 0,
        versionMadeBy: 20,
        versionNeeded: 20,
        modifiedTime: 0,
        modifiedDate: 0,
        localExtra: new Uint8Array(),
        centralExtra: new Uint8Array(),
        comment: new Uint8Array(),
        internalAttrs: 0,
        externalAttrs: 0,
        data,
      })),
  ].map((entry) => ({ ...entry, nameBytes: encoder.encode(entry.name) }));
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;
  for (const file of records) {
    const crc = crc32(file.data);
    const method = file.method === 0 ? 0 : 8;
    const flags = method === 8 ? file.flags & ~0x6 : file.flags;
    const payload = method === 0 ? file.data : await deflateRaw(file.data);
    const local = new Uint8Array(30 + file.nameBytes.length + file.localExtra.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, file.versionNeeded, true);
    localView.setUint16(6, flags, true);
    localView.setUint16(8, method, true);
    localView.setUint16(10, file.modifiedTime, true);
    localView.setUint16(12, file.modifiedDate, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, payload.length, true);
    localView.setUint32(22, file.data.length, true);
    localView.setUint16(26, file.nameBytes.length, true);
    localView.setUint16(28, file.localExtra.length, true);
    local.set(file.nameBytes, 30);
    local.set(file.localExtra, 30 + file.nameBytes.length);
    chunks.push(local, payload);

    const central = new Uint8Array(46 + file.nameBytes.length + file.centralExtra.length + file.comment.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, file.versionMadeBy, true);
    centralView.setUint16(6, file.versionNeeded, true);
    centralView.setUint16(8, flags, true);
    centralView.setUint16(10, method, true);
    centralView.setUint16(12, file.modifiedTime, true);
    centralView.setUint16(14, file.modifiedDate, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, payload.length, true);
    centralView.setUint32(24, file.data.length, true);
    centralView.setUint16(28, file.nameBytes.length, true);
    centralView.setUint16(30, file.centralExtra.length, true);
    centralView.setUint16(32, file.comment.length, true);
    centralView.setUint16(36, file.internalAttrs, true);
    centralView.setUint32(38, file.externalAttrs, true);
    centralView.setUint32(42, offset, true);
    central.set(file.nameBytes, 46);
    central.set(file.centralExtra, 46 + file.nameBytes.length);
    central.set(file.comment, 46 + file.nameBytes.length + file.centralExtra.length);
    centralDirectory.push(central);
    offset += local.length + payload.length;
  }
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
