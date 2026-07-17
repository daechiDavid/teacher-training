import { readSheet } from "read-excel-file/browser";
import {
  COMPLETION_HEADERS,
  CompletionRow,
  findMissingHeaders,
  mapRowsByHeaders,
  ROSTER_HEADERS,
  RosterRow,
} from "./workflow";

export type ParsedSheet<T> = {
  headers: string[];
  rows: T[];
  missingHeaders: string[];
};

export async function readWorkbookRows(file: File): Promise<unknown[][]> {
  const rows = await readSheet(file);
  if (!rows.length) {
    throw new Error("엑셀 파일에서 읽을 수 있는 내용이 없습니다.");
  }
  return rows.map((row: unknown[]) => row.map((value: unknown) => cellToString(value)));
}

function cellToString(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object" && "text" in value) return String((value as { text?: unknown }).text ?? "");
  if (typeof value === "object" && "result" in value) return String((value as { result?: unknown }).result ?? "");
  if (typeof value === "object" && "richText" in value) {
    const richText = (value as { richText?: Array<{ text?: unknown }> }).richText ?? [];
    return richText.map((part) => String(part.text ?? "")).join("");
  }
  return String(value);
}

export function parseCompletionWorkbookRows(rawRows: unknown[][]): ParsedSheet<CompletionRow> {
  const headers = rawRows[0]?.map((value) => String(value ?? "").trim()) ?? [];
  return {
    headers,
    rows: mapRowsByHeaders(COMPLETION_HEADERS, rawRows) as CompletionRow[],
    missingHeaders: findMissingHeaders(headers, COMPLETION_HEADERS),
  };
}

export function parseRosterWorkbookRows(rawRows: unknown[][]): ParsedSheet<RosterRow> {
  const headers = rawRows[0]?.map((value) => String(value ?? "").trim()) ?? [];
  const fixedRows = mapRowsByHeaders(ROSTER_HEADERS, rawRows) as RosterRow[];
  return {
    headers,
    rows: fixedRows.map((row, index) => ({
      ...row,
      __trainingRecords: headers.slice(12).map((header, offset) => ({
        header,
        value: String(rawRows[index + 1]?.[offset + 12] ?? "").trim(),
      })),
    })),
    missingHeaders: findMissingHeaders(headers, ROSTER_HEADERS),
  };
}
