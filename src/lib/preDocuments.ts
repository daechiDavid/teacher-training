import doc1TemplateUrl from "../../hwpx_template/doc1.hwpx?url";
import doc3TemplateUrl from "../../hwpx_template/doc3.hwpx?url";
import doc4TemplateUrl from "../../hwpx_template/doc4.hwpx?url";
import {
  createHwpxFromTemplateUrl,
  HwpxPreliminaryTrainingRow,
} from "./attendanceDocuments";

export type PreliminaryTrainingSet = {
  id?: string;
  trainingName: string;
  trainingDate: string;
  instructorName: string;
};

export type PreliminaryDocumentForm = {
  instituteName: string;
  trainings: PreliminaryTrainingSet[];
};

export type PreliminaryGeneratedDocument = {
  label: string;
  filename: string;
  blob: Blob;
};

export async function createPreliminaryDocuments(
  form: PreliminaryDocumentForm,
): Promise<PreliminaryGeneratedDocument[]> {
  if (!form.trainings.length) {
    throw new Error("최소 1개의 연수를 입력하세요.");
  }
  const firstTraining = form.trainings[0];
  const firstDate = parseTrainingDate(firstTraining.trainingDate);
  const common = {
    yyyy: firstDate.year,
    M: firstDate.month,
    "연수원명": form.instituteName.trim(),
  };
  const trainingRows: HwpxPreliminaryTrainingRow[] = form.trainings.map((training, index) => ({
    number: String(index + 1),
    name: training.trainingName,
    instructor: training.instructorName,
  }));

  const doc1 = await createHwpxFromTemplateUrl(doc1TemplateUrl, common, {
    preliminaryTrainingRows: trainingRows,
  });
  const doc3Documents = await Promise.all(form.trainings.map(async (training) => {
    const date = parseTrainingDate(training.trainingDate);
    return createHwpxFromTemplateUrl(doc3TemplateUrl, {
      ...common,
      yyyy: date.year,
      M: date.month,
      MM: date.month.padStart(2, "0"),
      // 확약서 본문의 제출일은 연수 예정 월의 첫날로 고정한다.
      dd: "01",
      연수명: training.trainingName,
    });
  }));
  const doc4 = await createHwpxFromTemplateUrl(doc4TemplateUrl, common, {
    preliminaryTrainingRows: trainingRows,
  });

  return [
    {
      label: "1. 실시간 쌍방향 연수과정 연수 계획서 제출 공문",
      filename: "1. 실시간 쌍방향 연수과정 연수 계획서 제출 공문.hwpx",
      blob: doc1,
    },
    ...doc3Documents.map((blob, index) => ({
      label: `3-${index + 1}. 확약서`,
      filename: `3-${index + 1}. 확약서.hwpx`,
      blob,
    })),
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

export function buildEvaluationFormTitle(trainingDate: string): string {
  const date = parseTrainingDate(trainingDate);
  return `화상원격연수 평가서(${date.month}.${date.day}.)`;
}

export function buildPreliminaryAssetFolderName(
  trainingDate: string,
  trainingIndex: number,
  trainings: PreliminaryTrainingSet[],
): string {
  const baseName = buildPreliminaryFolderName(trainingDate);
  const sameDateCount = trainings.filter((training) => {
    try {
      return buildPreliminaryFolderName(training.trainingDate) === baseName;
    } catch {
      return false;
    }
  }).length;
  return sameDateCount > 1 ? `${baseName} - ${trainingIndex + 1}` : baseName;
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
