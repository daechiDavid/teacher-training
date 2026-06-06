export const COMPLETION_HEADERS = [
  "연번",
  "연수과정명",
  "성명",
  "생년월일",
  "연락처",
  "이메일",
  "근무지역",
  "학교유형",
  "학교급",
  "학교명",
  "나이스번호",
  "접속시간",
  "이수결과",
  "비고",
] as const;

export const ROSTER_HEADERS = [
  "연번",
  "학교명",
  "이름",
  "전화번호",
  "영수증발급횟수",
  "과정명1",
  "발급날짜1",
  "링크1",
  "과정명2",
  "발급날짜2",
  "링크2",
] as const;

export type CompletionHeader = (typeof COMPLETION_HEADERS)[number];
export type RosterHeader = (typeof ROSTER_HEADERS)[number];

export type CompletionRow = Record<CompletionHeader, string>;
export type RosterRow = Record<RosterHeader, string>;

export type ValidationIssue = {
  severity: "error" | "warning";
  row?: number;
  field?: string;
  message: string;
};

export type MatchResult =
  | {
      status: "eligible";
      completion: NormalizedCompletion;
      roster: NormalizedRoster;
      nextSlot: 1 | 2;
      receiptFilename: string;
    }
  | {
      status: "manual-review" | "excluded";
      completion: NormalizedCompletion;
      reason: string;
      roster?: NormalizedRoster;
    };

export type NormalizedCompletion = {
  rowNumber: number;
  sequence: string;
  trainingName: string;
  name: string;
  birthDate: string;
  email: string;
  region: string;
  schoolLevel: string;
  school: string;
  niceNumber: string;
  phone: string;
  phoneLast4: string;
  completionStatus: string;
};

export type NormalizedRoster = {
  rowNumber: number;
  sequence: string;
  name: string;
  school: string;
  phone: string;
  issueCount: number;
  course1: string;
  issuedAt1: string;
  link1: string;
  course2: string;
  issuedAt2: string;
  link2: string;
};

export function normalizeText(value: unknown): string {
  return String(value ?? "").normalize("NFC").trim();
}

export function normalizePhone(value: unknown): string {
  return String(value ?? "").replace(/\D/g, "");
}

export function normalizeComparablePhone(value: unknown): string {
  const digits = normalizePhone(value);
  return /^\d{11}$/.test(digits) ? digits : "";
}

export function isValidMobilePhone(phone: string): boolean {
  return /^010\d{8}$/.test(normalizeComparablePhone(phone));
}

export function isValidNiceNumber(niceNumber: string): boolean {
  return /^[A-Za-z]\d{9}$/.test(normalizeText(niceNumber));
}

export function safeFileSegment(value: string): string {
  return normalizeText(value)
    .replace(/[\\/:*?"<>|#%{}[\]^~`]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function buildFolderName(trainingName: string, issuedDate: string): string {
  return `${safeFileSegment(trainingName)}_${safeFileSegment(issuedDate)}`;
}

export function buildReceiptFilename(
  name: string,
  school: string,
  phone: string,
  hasSameNameAndSchool: boolean,
): string {
  const base = `${safeFileSegment(name)}_${safeFileSegment(school)}`;
  return hasSameNameAndSchool ? `${base}_${phone.slice(-4)}.pdf` : `${base}.pdf`;
}

export function findMissingHeaders(headers: string[], required: readonly string[]): string[] {
  const present = new Set(headers.map(normalizeText));
  return required.filter((header) => !present.has(header));
}

export function mapRowsByHeaders<T extends string>(
  headers: readonly T[],
  rawRows: unknown[][],
): Record<T, string>[] {
  const headerRow = rawRows[0]?.map(normalizeText) ?? [];
  const indexes = new Map<T, number>();
  headers.forEach((header) => indexes.set(header, headerRow.indexOf(header)));

  return rawRows.slice(1).map((row) => {
    const mapped = {} as Record<T, string>;
    headers.forEach((header) => {
      const index = indexes.get(header) ?? -1;
      mapped[header] = index >= 0 ? normalizeText(row[index]) : "";
    });
    return mapped;
  });
}

export function normalizeCompletionRows(rows: CompletionRow[]): NormalizedCompletion[] {
  return normalizeCompletionRowsBase(rows).filter((row) => row.completionStatus === "이수");
}

export function normalizeApplicantRows(rows: CompletionRow[]): NormalizedCompletion[] {
  return normalizeCompletionRowsBase(rows);
}

function normalizeCompletionRowsBase(rows: CompletionRow[]): NormalizedCompletion[] {
  return rows
    .map((row, index) => ({
      rowNumber: index + 2,
      sequence: normalizeText(row["연번"]),
      trainingName: normalizeText(row["연수과정명"]),
      name: normalizeText(row["성명"]),
      birthDate: normalizeText(row["생년월일"]),
      email: normalizeText(row["이메일"]),
      region: normalizeText(row["근무지역"]),
      schoolLevel: normalizeText(row["학교급"]),
      school: normalizeText(row["학교명"]),
      niceNumber: normalizeText(row["나이스번호"]),
      phone: normalizePhone(row["연락처"]),
      phoneLast4: normalizePhone(row["연락처"]).slice(-4),
      completionStatus: normalizeText(row["이수결과"]),
    }))
    .filter((row) => row.name || row.niceNumber || row.school || row.trainingName);
}

export function normalizeRosterRows(rows: RosterRow[]): NormalizedRoster[] {
  return rows.map((row, index) => ({
    rowNumber: index + 2,
    sequence: normalizeText(row["연번"]),
    school: normalizeText(row["학교명"]),
    name: normalizeText(row["이름"]),
    phone: normalizePhone(row["전화번호"]),
    issueCount: Number.parseInt(normalizeText(row["영수증발급횟수"]) || "0", 10),
    course1: normalizeText(row["과정명1"]),
    issuedAt1: normalizeText(row["발급날짜1"]),
    link1: normalizeText(row["링크1"]),
    course2: normalizeText(row["과정명2"]),
    issuedAt2: normalizeText(row["발급날짜2"]),
    link2: normalizeText(row["링크2"]),
  }));
}

export function validateRosterIntegrity(rows: NormalizedRoster[]): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  rows.forEach((row) => {
    if (!Number.isInteger(row.issueCount) || row.issueCount < 0) {
      issues.push({
        severity: "error",
        row: row.rowNumber,
        field: "영수증발급횟수",
        message: "영수증발급횟수는 0 이상의 숫자여야 합니다.",
      });
      return;
    }

    const slot1Filled = Boolean(row.course1 && row.issuedAt1 && row.link1);
    const slot1Any = Boolean(row.course1 || row.issuedAt1 || row.link1);
    const slot2Filled = Boolean(row.course2 && row.issuedAt2 && row.link2);
    const slot2Any = Boolean(row.course2 || row.issuedAt2 || row.link2);

    if (row.issueCount === 0 && (slot1Any || slot2Any)) {
      issues.push({
        severity: "error",
        row: row.rowNumber,
        message: "영수증발급횟수 0인 행에는 과정명/발급날짜/링크 기록이 없어야 합니다.",
      });
    }

    if (row.issueCount === 1 && (!slot1Filled || slot2Any)) {
      issues.push({
        severity: "error",
        row: row.rowNumber,
        message: "영수증발급횟수 1인 행은 1차 기록만 완전해야 합니다.",
      });
    }

    if (row.issueCount === 2 && (!slot1Filled || !slot2Filled)) {
      issues.push({
        severity: "error",
        row: row.rowNumber,
        message: "영수증발급횟수 2인 행은 1차와 2차 기록이 모두 완전해야 합니다.",
      });
    }
  });

  return issues;
}

export function validateCompletionNiceNumbers(rows: NormalizedCompletion[]): ValidationIssue[] {
  return rows
    .filter((row) => !isValidNiceNumber(row.niceNumber))
    .map((row) => ({
      severity: "error",
      row: row.rowNumber,
      field: "나이스번호",
      message: `${row.name || "이름 없음"}: 나이스번호가 알파벳 1자리 + 숫자 9자리 형식이 아닙니다.`,
    }));
}

export function matchRecipients(
  completions: NormalizedCompletion[],
  rosterRows: NormalizedRoster[],
): MatchResult[] {
  const rosterByKey = new Map<string, NormalizedRoster[]>();
  const nameSchoolCounts = new Map<string, number>();

  rosterRows.forEach((row) => {
    const key = personKey(row.name, row.phone);
    const existing = rosterByKey.get(key) ?? [];
    existing.push(row);
    rosterByKey.set(key, existing);

    const nameSchoolKey = `${row.name}::${row.school}`;
    nameSchoolCounts.set(nameSchoolKey, (nameSchoolCounts.get(nameSchoolKey) ?? 0) + 1);
  });

  return completions.map((completion) => {
    if (!completion.trainingName || !completion.name || !completion.phone) {
      return {
        status: "manual-review",
        completion,
        reason: "연수과정명, 성명, 연락처는 필수입니다.",
      };
    }

    if (!isValidMobilePhone(completion.phone)) {
      return {
        status: "manual-review",
        completion,
        reason: "전화번호가 010으로 시작하는 11자리 휴대전화 형식이 아닙니다.",
      };
    }

    if (!isValidNiceNumber(completion.niceNumber)) {
      return {
        status: "manual-review",
        completion,
        reason: "나이스번호가 알파벳 1자리 + 숫자 9자리 형식이 아닙니다.",
      };
    }

    const matches = rosterByKey.get(personKey(completion.name, completion.phone)) ?? [];

    if (matches.length !== 1) {
      return {
        status: "manual-review",
        completion,
        reason: matches.length === 0 ? "명단 탭에서 일치하는 행을 찾지 못했습니다." : "명단 탭에서 중복 행이 발견되었습니다.",
      };
    }

    const roster = matches[0];

    if (roster.issueCount >= 2) {
      return {
        status: "excluded",
        completion,
        roster,
        reason: "이미 영수증발급횟수가 2회 이상입니다.",
      };
    }

    if ([roster.course1, roster.course2].includes(completion.trainingName)) {
      return {
        status: "excluded",
        completion,
        roster,
        reason: "같은 연수과정명으로 이미 발급된 기록이 있습니다.",
      };
    }

    const hasSameNameAndSchool = (nameSchoolCounts.get(`${completion.name}::${completion.school}`) ?? 0) > 1;
    return {
      status: "eligible",
      completion,
      roster,
      nextSlot: roster.issueCount === 0 ? 1 : 2,
      receiptFilename: buildReceiptFilename(
        completion.name,
        completion.school,
        completion.phone,
        hasSameNameAndSchool,
      ),
    };
  });
}

export function personKey(name: string, phone: string): string {
  return `${normalizeText(name)}::${normalizeComparablePhone(phone)}`;
}
