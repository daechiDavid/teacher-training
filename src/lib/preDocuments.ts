import doc1TemplateUrl from "../../hwpx_template/doc1.hwpx?url";
import doc3TemplateUrl from "../../hwpx_template/doc3.hwpx?url";
import doc4TemplateUrl from "../../hwpx_template/doc4.hwpx?url";
import { createHwpxFromTemplateUrl } from "./attendanceDocuments";

export type PreliminaryTrainingSet = {
  trainingName: string;
  trainingDate: string;
  instructorName: string;
};

export type PreliminaryDocumentForm = {
  training1: PreliminaryTrainingSet;
  training2: PreliminaryTrainingSet;
};

export type PreliminaryGeneratedDocument = {
  label: string;
  filename: string;
  blob: Blob;
};

export async function createPreliminaryDocuments(
  form: PreliminaryDocumentForm,
): Promise<PreliminaryGeneratedDocument[]> {
  const date1 = parseTrainingDate(form.training1.trainingDate);
  const date2 = parseTrainingDate(form.training2.trainingDate);
  const common = {
    yyyy: date1.year,
    M: date1.month,
    "연수1_연수명": form.training1.trainingName,
    "연수1_강사": form.training1.instructorName,
    "연수2_연수명": form.training2.trainingName,
    "연수2_강사": form.training2.instructorName,
  };

  const doc1 = await createHwpxFromTemplateUrl(doc1TemplateUrl, common);
  const doc3A = await createHwpxFromTemplateUrl(doc3TemplateUrl, {
    ...common,
    yyyy: date1.year,
    MM: date1.month.padStart(2, "0"),
    dd: date1.day.padStart(2, "0"),
    연수명: form.training1.trainingName,
  });
  const doc3B = await createHwpxFromTemplateUrl(doc3TemplateUrl, {
    ...common,
    yyyy: date2.year,
    MM: date2.month.padStart(2, "0"),
    dd: date2.day.padStart(2, "0"),
    연수명: form.training2.trainingName,
  });
  const doc4 = await createHwpxFromTemplateUrl(doc4TemplateUrl, common);

  return [
    {
      label: "1. 실시간 쌍방향 연수과정 연수 계획서 제출 공문",
      filename: "1. 실시간 쌍방향 연수과정 연수 계획서 제출 공문.hwpx",
      blob: doc1,
    },
    {
      label: "3-1. 확약서",
      filename: "3-1. 확약서.hwpx",
      blob: doc3A,
    },
    {
      label: "3-2. 확약서",
      filename: "3-2. 확약서.hwpx",
      blob: doc3B,
    },
    {
      label: "4. 실시간쌍방향 연수 심의결과서",
      filename: "4. 실시간쌍방향 연수 심의결과서.hwpx",
      blob: doc4,
    },
  ];
}

export function buildPreliminaryFolderName(trainingDate: string): string {
  const date = parseTrainingDate(trainingDate);
  return `${date.shortYear}.${date.month.padStart(2, "0")}.${date.day.padStart(2, "0")}.(${date.dow})`;
}

export function buildConsentFormTitle(trainingDate: string): string {
  const date = parseTrainingDate(trainingDate);
  return `${date.month}월 ${date.day}일(${date.dow}) 쌍방향 ZOOM 원격직무연수 동의서`;
}

export function parseTrainingDate(value: string): {
  year: string;
  shortYear: string;
  month: string;
  day: string;
  dow: string;
} {
  const trimmed = value.trim();
  const match = trimmed.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (!match) {
    throw new Error("연수 날짜는 YYYY-MM-DD 형식으로 입력하세요.");
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    throw new Error("연수 날짜가 올바르지 않습니다.");
  }
  const dows = ["일", "월", "화", "수", "목", "금", "토"];
  return {
    year: String(year),
    shortYear: String(year).slice(2),
    month: String(month),
    day: String(day),
    dow: dows[date.getDay()],
  };
}
