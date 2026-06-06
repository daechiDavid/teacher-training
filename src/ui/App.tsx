import {
  AlertCircle,
  ArrowLeft,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  FolderUp,
  LockKeyhole,
  ReceiptText,
  Settings,
  UploadCloud,
} from "lucide-react";
import { save } from "@tauri-apps/plugin-dialog";
import { writeFile } from "@tauri-apps/plugin-fs";
import { useEffect, useMemo, useState } from "react";
import {
  matchRecipients,
  normalizeApplicantRows,
  normalizeCompletionRows,
  personKey,
  normalizeRosterRows,
  validateCompletionNiceNumbers,
  validateRosterIntegrity,
  NormalizedCompletion,
  NormalizedRoster,
  ValidationIssue,
} from "../lib/workflow";
import { parseCompletionWorkbookRows, parseRosterWorkbookRows } from "../lib/sheet";
import {
  CompletionDocumentRow,
  getTrainingNameOptions,
} from "../lib/completionDocument";
import {
  AttendanceBaseForm,
  AttendancePerson,
  applyZoomChatAttendanceText,
  buildAttendanceFilename,
  buildAttendancePeople,
  buildCompletionRowsWithZoomMinutes,
  buildDefaultCaptureRows,
  buildSummaryRows,
  CaptureAttendanceRow,
  CaptureEvidenceImage,
  CaptureEvidenceRow,
  CaptureMode,
  countSummary,
  createCaptureHwpx,
  createCompletionXlsx,
  createEvaluationHwpx,
  createSummaryHwpx,
  createZoomHwpx,
  EvaluationSummary,
  parseEvaluationRows,
  parseZoomAttendanceWorkbook,
  SummaryAttendanceRow,
  updateCaptureResult,
  ZoomAttendanceRow,
} from "../lib/attendanceDocuments";
import {
  appendJobLog,
  batchUpdateSheet,
  batchUpdateGoogleSheet,
  createDriveTrainingFolder,
  getAppStatus,
  getGoogleConfigStatus,
  GoogleConfigStatus,
  readGoogleRosterValues,
  readGoogleSheetValues,
  resolveGoogleSheetTitle,
  saveGoogleConfig,
  startGoogleOAuth,
  googleLogout,
  uploadPdfToDrive,
} from "../lib/google";
import { generateReceiptPdf, todayLocalDate } from "../lib/receipt";

type LoadedFileState = {
  name: string;
  rowCount: number;
  missingHeaders: string[];
};

type WorkflowStep = "settings" | "upload" | "review" | "issue";
type DocumentWorkflowStep = "source" | "doc1" | "doc2" | "doc3" | "doc5" | "doc7" | "done";
type ActiveTask = "completion" | "receipt" | null;

type GoogleUrlForm = {
  spreadsheetUrl: string;
  driveFolderUrl: string;
};

type CompletionSheetForm = {
  spreadsheetUrl: string;
};

type EvaluationSheetForm = {
  spreadsheetUrl: string;
};

type GeneratedDocument = {
  label: string;
  filename: string;
  blob: Blob;
};

type DocumentSheetSource = {
  spreadsheetId: string;
  sheetName: string;
};

function fileExtensionFilter(filename: string): { name: string; extensions: string[] } | undefined {
  const ext = filename.split(".").pop()?.toLowerCase();
  if (ext === "hwpx") return { name: "한글 문서", extensions: ["hwpx"] };
  if (ext === "xlsx") return { name: "엑셀 문서", extensions: ["xlsx"] };
  if (ext === "pdf") return { name: "영수증 파일", extensions: ["pdf"] };
  return undefined;
}

async function saveDocumentAs(doc: GeneratedDocument): Promise<void> {
  const filter = fileExtensionFilter(doc.filename);
  const path = await save({
    defaultPath: doc.filename,
    filters: filter ? [filter] : [],
  });
  if (!path) return;
  const bytes = new Uint8Array(await doc.blob.arrayBuffer());
  await writeFile(path, bytes);
}

type IssueProgressState = {
  current: number;
  total: number;
  label: string;
};

type IssueCompletionState = {
  count: number;
  rosterUrl: string;
  folderLinks: Array<{ name: string; url: string }>;
};

function createEmptyCaptureEvidenceRow(): CaptureEvidenceRow {
  return {
    id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random()}`,
    period: 1,
    mode: "camera",
    cameraImage: null,
    chatImages: [],
  };
}

export function App() {
  const [step, setStep] = useState<WorkflowStep>("settings");
  const [activeTask, setActiveTask] = useState<ActiveTask>(null);
  const [completionFile, setCompletionFile] = useState<LoadedFileState | null>(null);
  const [rosterFile, setRosterFile] = useState<LoadedFileState | null>(null);
  const [completionRows, setCompletionRows] = useState<ReturnType<typeof normalizeCompletionRows>>([]);
  const [rosterRows, setRosterRows] = useState<ReturnType<typeof normalizeRosterRows>>([]);
  const [rosterHeaders, setRosterHeaders] = useState<string[]>([]);
  const [documentSourceFile, setDocumentSourceFile] = useState<LoadedFileState | null>(null);
  const [documentRows, setDocumentRows] = useState<ReturnType<typeof normalizeApplicantRows>>([]);
  const [documentSheetSource, setDocumentSheetSource] = useState<DocumentSheetSource | null>(null);
  const [documentForm, setDocumentForm] = useState<AttendanceBaseForm>({
    trainingName: "",
    institute: "",
    trainingDate: "",
    totalHours: "",
    startTime: "",
    endTime: "",
    period1Label: "",
    period2Label: "",
    instructorName: "",
  });
  const [documentPreviewRows, setDocumentPreviewRows] = useState<CompletionDocumentRow[]>([]);
  const [documentPeople, setDocumentPeople] = useState<AttendancePerson[]>([]);
  const [captureRows, setCaptureRows] = useState<CaptureAttendanceRow[]>([]);
  const [zoomRows, setZoomRows] = useState<ZoomAttendanceRow[]>([]);
  const [summaryRows, setSummaryRows] = useState<SummaryAttendanceRow[]>([]);
  const [captureEvidenceRows, setCaptureEvidenceRows] = useState<CaptureEvidenceRow[]>([createEmptyCaptureEvidenceRow()]);
  const [zoomChatFile, setZoomChatFile] = useState<LoadedFileState | null>(null);
  const [zoomFile, setZoomFile] = useState<LoadedFileState | null>(null);
  const [evaluationSheetForm, setEvaluationSheetForm] = useState<EvaluationSheetForm>({ spreadsheetUrl: "" });
  const [evaluationSummary, setEvaluationSummary] = useState<EvaluationSummary | null>(null);
  const [documentWorkflowStep, setDocumentWorkflowStep] = useState<DocumentWorkflowStep>("source");
  const [currentDocument, setCurrentDocument] = useState<GeneratedDocument | null>(null);
  const [currentDocumentDownloaded, setCurrentDocumentDownloaded] = useState(false);
  const [rosterIssues, setRosterIssues] = useState<ValidationIssue[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [googleStatus, setGoogleStatus] = useState<GoogleConfigStatus | null>(null);
  const [googleForm, setGoogleForm] = useState<GoogleUrlForm>({
    spreadsheetUrl: "",
    driveFolderUrl: "",
  });
  const [completionSheetForm, setCompletionSheetForm] = useState<CompletionSheetForm>({
    spreadsheetUrl: "",
  });
  const [samplePdfUrl, setSamplePdfUrl] = useState<string | null>(null);
  const [issueProgress, setIssueProgress] = useState<IssueProgressState | null>(null);
  const [issueCompletion, setIssueCompletion] = useState<IssueCompletionState | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    void refreshGoogleStatus({ initialize: true });
  }, []);

  useEffect(() => {
    return () => { void currentDocument; };
  }, [currentDocument]);

  const results = useMemo(
    () => matchRecipients(completionRows, rosterRows),
    [completionRows, rosterRows],
  );
  const eligibleResults = useMemo(
    () => results.filter((result) => result.status === "eligible"),
    [results],
  );
  const manualReviewResults = useMemo(
    () => results.filter((result) => result.status === "manual-review"),
    [results],
  );
  const summary = useMemo(
    () => ({
      completed: completionRows.length,
      eligible: eligibleResults.length,
      manual: manualReviewResults.length + rosterIssues.length,
      excluded: results.filter((result) => result.status === "excluded").length,
    }),
    [completionRows.length, eligibleResults.length, manualReviewResults.length, results, rosterIssues.length],
  );
  const documentIssues = useMemo(
    () => validateCompletionNiceNumbers(documentRows),
    [documentRows],
  );

  useEffect(() => {
    if (step !== "issue" || !eligibleResults[0]) {
      setSamplePdfUrl((current) => {
        if (current) URL.revokeObjectURL(current);
        return null;
      });
      return;
    }

    let cancelled = false;
    async function createSample() {
      const first = eligibleResults[0];
      if (!first || first.status !== "eligible") return;
      const pdfBytes = await generateReceiptPdf({
        completion: first.completion,
        roster: first.roster,
        issuedDate: todayLocalDate(),
      });
      if (cancelled) return;
      const nextUrl = URL.createObjectURL(new Blob([pdfBytes], { type: "application/pdf" }));
      setSamplePdfUrl((current) => {
        if (current) URL.revokeObjectURL(current);
        return nextUrl;
      });
    }

    createSample().catch((sampleError) => {
      setError(sampleError instanceof Error ? sampleError.message : "샘플 영수증을 생성하지 못했습니다.");
    });

    return () => {
      cancelled = true;
    };
  }, [eligibleResults, step]);

  async function refreshGoogleStatus(options?: { initialize?: boolean }) {
    const status = await getGoogleConfigStatus();
    setGoogleStatus(status);
    setGoogleForm({
      spreadsheetUrl: status.spreadsheet_id,
      driveFolderUrl: status.drive_parent_folder_id,
    });
    if (options?.initialize && status.authenticated && status.configured) {
      setStep("upload");
    }
    return status;
  }

  async function connectGoogle() {
    setError(null);
    setNotice("브라우저에서 구글 로그인을 완료한 뒤 이 화면으로 돌아오세요.");
    setBusy(true);
    try {
      await startGoogleOAuth();
      const status = await refreshGoogleStatus();
      setStep(status.configured ? "upload" : "settings");
      setActiveTask(null);
      setNotice("구글 로그인이 완료되었습니다.");
    } catch (authError) {
      setError(authError instanceof Error ? authError.message : String(authError));
    } finally {
      setBusy(false);
    }
  }

  async function disconnectGoogle() {
    setError(null);
    setBusy(true);
    try {
      await googleLogout();
      await refreshGoogleStatus();
      setStep("settings");
      setActiveTask(null);
      setNotice("구글 로그아웃이 완료되었습니다. 다른 계정을 사용하려면 다시 로그인하세요.");
    } catch (logoutError) {
      setError(logoutError instanceof Error ? logoutError.message : String(logoutError));
    } finally {
      setBusy(false);
    }
  }

  function chooseTask(task: Exclude<ActiveTask, null>) {
    setError(null);
    setNotice(null);
    setActiveTask(task);
    if (task === "receipt") {
      setStep(googleStatus?.configured ? "upload" : "settings");
    }
  }

  function returnToMenu() {
    setError(null);
    setNotice(null);
    setActiveTask(null);
  }

  async function loadCompletionDocumentSource() {
    setError(null);
    setNotice(null);
    setBusy(true);
    try {
      const spreadsheetId = extractSpreadsheetId(completionSheetForm.spreadsheetUrl, "연수 신청자 명단");
      const gid = extractGoogleSheetGid(completionSheetForm.spreadsheetUrl);
      const sheetName = await resolveGoogleSheetTitle(spreadsheetId, gid);
      const rawRows = await readGoogleSheetValues(spreadsheetId, gid);
      const parsed = parseCompletionWorkbookRows(rawRows);
      if (parsed.missingHeaders.length) {
        throw new Error(`연수 신청자 명단에서 필수 항목을 찾지 못했습니다: ${parsed.missingHeaders.join(", ")}`);
      }
      const normalized = normalizeApplicantRows(parsed.rows);
      const trainingNameOptions = getTrainingNameOptions(normalized);
      const people = buildAttendancePeople(normalized);
      setDocumentRows(normalized);
      setDocumentSheetSource({ spreadsheetId, sheetName });
      setDocumentPeople(people);
      setCaptureRows(buildDefaultCaptureRows(people));
      setCaptureEvidenceRows([createEmptyCaptureEvidenceRow()]);
      setZoomChatFile(null);
      setZoomRows([]);
      setZoomFile(null);
      setSummaryRows([]);
      setDocumentPreviewRows([]);
      setEvaluationSheetForm({ spreadsheetUrl: "" });
      setEvaluationSummary(null);
      setDocumentSourceFile({
        name: `연수 신청자 명단: ${sheetName}`,
        rowCount: normalized.length,
        missingHeaders: parsed.missingHeaders,
      });
      setDocumentForm((current) => ({
        ...current,
        trainingName: current.trainingName || trainingNameOptions[0] || "",
      }));
      setDocumentWorkflowStep("doc1");
      clearCurrentDocument();
      setNotice(`${normalized.length.toLocaleString()}명의 연수 신청자를 불러왔습니다.`);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : String(loadError));
    } finally {
      setBusy(false);
    }
  }

  function validateDocumentBasics() {
    if (!documentRows.length) {
      setError("연수 신청자 명단을 먼저 불러오세요.");
      return false;
    }
    if (documentIssues.length) {
      setError("나이스번호 형식 오류가 남아 있습니다.");
      return false;
    }
    if (!documentForm.trainingName.trim() || !documentForm.institute.trim() || !documentForm.trainingDate.trim() || !documentForm.totalHours.trim() || !documentForm.startTime.trim() || !documentForm.endTime.trim()) {
      setError("연수과정명, 연수기관, 연수날짜, 연수 총 시간, 시작시간, 종료시간을 모두 입력하세요.");
      return false;
    }
    if (!isValidTimeText(documentForm.startTime) || !isValidTimeText(documentForm.endTime)) {
      setError("시작시간과 종료시간은 HH:mm 형식으로 입력하세요.");
      return false;
    }
    setError(null);
    return true;
  }

  async function createCurrentDocument() {
    if (!validateDocumentBasics()) return;
    if (documentWorkflowStep === "doc2" && !zoomRows.length) {
      setError("2번 문서를 만들려면 줌 접속기록 파일을 먼저 불러오세요.");
      return;
    }
    if (documentWorkflowStep === "doc3" && !zoomRows.length) {
      setError("3번 문서를 만들려면 줌 접속기록을 먼저 반영하세요.");
      return;
    }
    if (documentWorkflowStep === "doc5" && !evaluationSummary) {
      setError("5번 평가서를 만들려면 평가 결과를 먼저 불러오세요.");
      return;
    }
    if (documentWorkflowStep === "doc7" && !summaryRows.length) {
      setError("7번 명단을 만들려면 3번 종합 출결자료를 먼저 작성하세요.");
      return;
    }
    setBusy(true);
    try {
      const built = await buildDocumentForCurrentStep();
      clearCurrentDocument();
      setCurrentDocument({
        label: built.label,
        filename: built.filename,
        blob: built.blob,
      });
      setCurrentDocumentDownloaded(false);
      setError(null);
      setNotice(`${built.label} 파일을 만들었습니다. 파일을 저장한 뒤 다음 문서로 이동하세요.`);
    } catch (createError) {
      setError(createError instanceof Error ? createError.message : String(createError));
    } finally {
      setBusy(false);
    }
  }

  async function buildDocumentForCurrentStep() {
    if (documentWorkflowStep === "doc1") {
      return {
        label: "1. (출결) 연수생 화면 캡처 출결자료",
        filename: buildAttendanceFilename(documentForm, "1. (출결) 연수생 화면 캡처 출결자료", "hwpx"),
        blob: await createCaptureHwpx(documentForm, captureRows, captureEvidenceRows),
      };
    }
    if (documentWorkflowStep === "doc2") {
      if (documentSheetSource) {
        await batchUpdateGoogleSheet(
          documentSheetSource.spreadsheetId,
          documentSheetSource.sheetName,
          zoomRows.map((row) => ({
            range: `L${row.source.rowNumber}`,
            values: [[String(row.effectiveMinutes)]],
          })),
        );
      }
      return {
        label: "2. (출결) 연수 입장 및 퇴장 시간 기록자료",
        filename: buildAttendanceFilename(documentForm, "2. (출결) 연수 입장 및 퇴장 시간 기록자료", "hwpx"),
        blob: await createZoomHwpx(documentForm, zoomRows),
      };
    }
    if (documentWorkflowStep === "doc3") {
      const nextSummaryRows = summaryRows.length ? summaryRows : buildSummaryRows(captureRows, zoomRows);
      setSummaryRows(nextSummaryRows);
      if (documentSheetSource) {
        await batchUpdateGoogleSheet(
          documentSheetSource.spreadsheetId,
          documentSheetSource.sheetName,
          nextSummaryRows.map((row) => ({
            range: `M${row.source.rowNumber}`,
            values: [[row.result3]],
          })),
        );
      }
      return {
        label: "3. (출결) 쌍방향 ZOOM 연수 교육 종합 출결 자료",
        filename: buildAttendanceFilename(documentForm, "3. (출결) 쌍방향 ZOOM 연수 교육 종합 출결 자료", "hwpx"),
        blob: await createSummaryHwpx(documentForm, nextSummaryRows),
      };
    }
    if (documentWorkflowStep === "doc5") {
      const counts = countSummary(summaryRows);
      return {
        label: "5. 쌍방향원격교육종합평가서",
        filename: buildAttendanceFilename(documentForm, "5. 쌍방향원격교육종합평가서", "hwpx"),
        blob: await createEvaluationHwpx(documentForm, evaluationSummary!, counts.completionCount, counts.incompleteCount),
      };
    }
    if (documentWorkflowStep === "doc7") {
      const summaryByNiceNumber = new Map(summaryRows.map((row) => [row.niceNumber, row]));
      const completedPeople = documentPeople.filter((person) =>
        summaryByNiceNumber.get(person.niceNumber)?.result3 === "이수",
      );
      const completionRows = buildCompletionRowsWithZoomMinutes(completedPeople, documentForm);
      setDocumentPreviewRows(completionRows);
      return {
        label: "7. 직무연수 이수자 명단",
        filename: buildAttendanceFilename(documentForm, "7. 직무연수 이수자 명단", "xlsx"),
        blob: await createCompletionXlsx(completedPeople, documentForm, zoomRows),
      };
    }
    throw new Error("현재 단계에서 생성할 문서가 없습니다.");
  }

  function updateCompletionDocumentForm(nextForm: AttendanceBaseForm) {
    const timeChanged = documentForm.startTime !== nextForm.startTime || documentForm.endTime !== nextForm.endTime;
    setDocumentForm(nextForm);
    setDocumentPreviewRows([]);
    if (timeChanged) {
      setZoomChatFile(null);
      setZoomRows([]);
      setZoomFile(null);
      setSummaryRows([]);
    }
    clearCurrentDocument();
  }

  function clearCurrentDocument() {
    setCurrentDocument(null);
    setCurrentDocumentDownloaded(false);
  }

  async function loadZoomFile(file: File | null) {
    if (!file) return;
    if (!isValidTimeText(documentForm.startTime) || !isValidTimeText(documentForm.endTime)) {
      setError("줌 접속기록을 불러오기 전에 연수 시작시간과 종료시간을 HH:mm 형식으로 입력하세요.");
      return;
    }
    setError(null);
    setBusy(true);
    try {
      const rows = await parseZoomAttendanceWorkbook(file, documentPeople, documentForm);
      const nextSummaryRows = buildSummaryRows(captureRows, rows);
      setZoomRows(rows);
      setSummaryRows(nextSummaryRows);
      setZoomFile({ name: file.name, rowCount: rows.length, missingHeaders: [] });
      clearCurrentDocument();
      setNotice("줌 접속기록을 2번 문서에 반영했습니다.");
    } catch (zoomError) {
      setError(zoomError instanceof Error ? zoomError.message : String(zoomError));
    } finally {
      setBusy(false);
    }
  }

  async function loadZoomChatFile(file: File | null) {
    if (!file) return;
    if (!isValidTimeText(documentForm.startTime) || !isValidTimeText(documentForm.endTime)) {
      setError("줌 채팅기록을 불러오기 전에 연수 시작시간과 종료시간을 HH:mm 형식으로 입력하세요.");
      return;
    }
    setError(null);
    setBusy(true);
    try {
      const applied = await applyZoomChatAttendanceText(file, captureRows, documentForm);
      setCaptureRows(applied.rows);
      setSummaryRows(zoomRows.length ? buildSummaryRows(applied.rows, zoomRows) : []);
      setZoomChatFile({
        name: file.name,
        rowCount: applied.startMatches + applied.endMatches,
        missingHeaders: [],
      });
      clearCurrentDocument();
      if (applied.startMatches + applied.endMatches === 0) {
        setNotice(null);
        setError(`줌 채팅기록은 읽었지만 출결에 반영된 수강생이 없습니다. 확인 구간: 1교시 ${applied.startWindowLabel}, 2교시 ${applied.endWindowLabel}. 채팅 이름이 명단의 성명과 맞는지 확인하세요.`);
      } else {
        setNotice(`줌 채팅기록을 반영했습니다. 기준시각 ${applied.anchorTime}, 1교시 ${applied.startMatches}명, 2교시 ${applied.endMatches}명을 O로 변경했습니다. 캠으로 출석체크한 수강생은 직접 이수처리해주세요.`);
      }
    } catch (chatError) {
      setError(chatError instanceof Error ? chatError.message : String(chatError));
    } finally {
      setBusy(false);
    }
  }

  async function loadEvaluationSheet() {
    setError(null);
    setBusy(true);
    try {
      const spreadsheetId = extractSpreadsheetId(evaluationSheetForm.spreadsheetUrl, "평가 결과");
      const gid = extractGoogleSheetGid(evaluationSheetForm.spreadsheetUrl);
      const rawRows = await readGoogleSheetValues(spreadsheetId, gid);
      setEvaluationSummary(parseEvaluationRows(rawRows));
      clearCurrentDocument();
      setNotice("평가 결과를 5번 평가서에 반영했습니다.");
    } catch (evaluationError) {
      setError(evaluationError instanceof Error ? evaluationError.message : String(evaluationError));
    } finally {
      setBusy(false);
    }
  }

  async function loadCaptureEvidenceImages(id: string, kind: "camera" | "chat", files: FileList | null) {
    if (!files) return;
    const images: CaptureEvidenceImage[] = await Promise.all(Array.from(files).slice(0, kind === "camera" ? 1 : 4).map(async (file) => ({
      name: file.name.replace(/\.[^.]+$/, ".jpg"),
      dataUrl: await resizeImageFileForHwpx(file, kind),
    })));
    setCaptureEvidenceRows((current) => current.map((row) => {
      if (row.id !== id) return row;
      return kind === "camera" ? { ...row, cameraImage: images[0] ?? null } : { ...row, chatImages: images };
    }));
    clearCurrentDocument();
  }

  function updateCaptureEvidenceRow(id: string, patch: Partial<Pick<CaptureEvidenceRow, "period" | "mode">>) {
    setCaptureEvidenceRows((current) => current.map((row) => {
      if (row.id !== id) return row;
      const next = { ...row, ...patch };
      if (patch.mode && patch.mode !== row.mode) {
        return { ...next, cameraImage: null, chatImages: [] };
      }
      return next;
    }));
    clearCurrentDocument();
  }

  function addCaptureEvidenceRow() {
    setCaptureEvidenceRows((current) => [...current, createEmptyCaptureEvidenceRow()]);
    clearCurrentDocument();
  }

  function removeCaptureEvidenceRow(id: string) {
    setCaptureEvidenceRows((current) => current.length > 1 ? current.filter((row) => row.id !== id) : current);
    clearCurrentDocument();
  }

  function updateCaptureRow(index: number, patch: Partial<Pick<CaptureAttendanceRow, "period1" | "period2" | "result">>) {
    setCaptureRows((current) => {
      const next = current.map((row, rowIndex) => {
        if (rowIndex !== index) return row;
        const patched = { ...row, ...patch };
        return patch.result ? patched : updateCaptureResult(patched);
      });
      setSummaryRows(zoomRows.length ? buildSummaryRows(next, zoomRows) : []);
      return next;
    });
    clearCurrentDocument();
  }

  function updateZoomRow(index: number, patch: Partial<Pick<ZoomAttendanceRow, "result">>) {
    setZoomRows((current) => {
      const next = current.map((row, rowIndex) => rowIndex === index ? { ...row, ...patch } : row);
      setSummaryRows(captureRows.length ? buildSummaryRows(captureRows, next) : []);
      return next;
    });
    clearCurrentDocument();
  }

  function markCurrentDocumentDownloaded() {
    setCurrentDocumentDownloaded(true);
  }

  function moveToDocumentStep(step: DocumentWorkflowStep) {
    clearCurrentDocument();
    setDocumentWorkflowStep(step);
    if (step === "doc3") setSummaryRows(buildSummaryRows(captureRows, zoomRows));
    if (step === "doc7") {
      const summaryByNiceNumber = new Map(summaryRows.map((row) => [row.niceNumber, row]));
      const completedPeople = documentPeople.filter((person) =>
        summaryByNiceNumber.get(person.niceNumber)?.result3 === "이수",
      );
      setDocumentPreviewRows(buildCompletionRowsWithZoomMinutes(completedPeople, documentForm));
    }
  }

  async function saveBaseSettings() {
    setError(null);
    setNotice(null);
    setBusy(true);
    try {
      const spreadsheetId = extractSpreadsheetId(googleForm.spreadsheetUrl, "전체명단");
      const spreadsheetGid = extractGoogleSheetGid(googleForm.spreadsheetUrl);
      const sheetName = await resolveGoogleSheetTitle(spreadsheetId, spreadsheetGid);
      const driveFolderId = extractDriveFolderId(googleForm.driveFolderUrl);
      await saveGoogleConfig({
        spreadsheetId,
        sheetName,
        driveParentFolderId: driveFolderId,
      });
      await refreshGoogleStatus();
      setStep("upload");
      setNotice("기본 설정을 저장했습니다. 이후 실행에서는 이 단계를 건너뜁니다.");
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : String(saveError));
    } finally {
      setBusy(false);
    }
  }

  async function loadCompletionFromGoogleSheet() {
    setError(null);
    setNotice(null);
    setBusy(true);
    try {
      const spreadsheetId = extractSpreadsheetId(completionSheetForm.spreadsheetUrl, "이수자 명단");
      const gid = extractGoogleSheetGid(completionSheetForm.spreadsheetUrl);
      const rawRows = await readGoogleSheetValues(spreadsheetId, gid);
      const parsed = parseCompletionWorkbookRows(rawRows);
      if (parsed.missingHeaders.length) {
        throw new Error(`이수자 명단에서 필수 항목을 찾지 못했습니다: ${parsed.missingHeaders.join(", ")}`);
      }
      const normalized = normalizeCompletionRows(parsed.rows);
      setCompletionRows(normalized);
      setCompletionFile({
        name: `이수자 명단${gid == null ? "" : " 선택한 탭"}`,
        rowCount: normalized.length,
        missingHeaders: parsed.missingHeaders,
      });
      await loadRosterFromGoogle();
      setStep("review");
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : String(loadError));
    } finally {
      setBusy(false);
    }
  }

  async function loadRosterFromGoogle() {
    const rawRows = await readGoogleRosterValues();
    const parsed = parseRosterWorkbookRows(rawRows);
    if (parsed.missingHeaders.length) {
      throw new Error(`전체명단에서 필수 항목을 찾지 못했습니다: ${parsed.missingHeaders.join(", ")}`);
    }
    const normalized = normalizeRosterRows(parsed.rows);
    setRosterRows(normalized);
    setRosterHeaders(parsed.headers);
    setRosterIssues(validateRosterIntegrity(normalized));
    setRosterFile({
      name: `전체명단: ${googleStatus?.sheet_name || "명단"}`,
      rowCount: normalized.length,
      missingHeaders: parsed.missingHeaders,
    });
  }

  async function reloadRosterForReview() {
    setError(null);
    setNotice(null);
    setBusy(true);
    try {
      await loadRosterFromGoogle();
      setNotice("전체명단을 다시 불러왔습니다.");
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : String(loadError));
    } finally {
      setBusy(false);
    }
  }

  function confirmReview() {
    if (summary.manual > 0) {
      setError("수동 확인 대상 또는 전체명단 정합성 오류가 남아 있습니다.");
      return;
    }
    if (!eligibleResults.length && !completionRows.length) {
      setError("기록할 이수자가 없습니다.");
      return;
    }
    setError(null);
    setNotice(null);
    setIssueCompletion(null);
    setStep("issue");
  }

  function resetToUpload() {
    setError(null);
    setNotice(null);
    setIssueProgress(null);
    setIssueCompletion(null);
    setStep("upload");
    setCompletionFile(null);
    setCompletionRows([]);
    setCompletionSheetForm({ spreadsheetUrl: "" });
  }

  function returnToUploadFromReview() {
    setError(null);
    setNotice(null);
    setStep("upload");
    setCompletionFile(null);
    setCompletionRows([]);
  }

  async function generateUploadAndRecord() {
    if (!eligibleResults.length && !completionRows.length) {
      setError("기록할 이수자가 없습니다.");
      return;
    }
    if (!googleStatus?.configured || !googleStatus.authenticated) {
      setError("구글 로그인과 기본 설정을 먼저 완료하세요.");
      return;
    }

    setError(null);
    setNotice(null);
    setIssueCompletion(null);
    setIssueProgress({
      current: 0,
      total: eligibleResults.length,
      label: "영수증 파일 저장을 준비 중입니다.",
    });
    setBusy(true);
    try {
      const issuedDate = todayLocalDate();
      const { run_id: runId } = await getAppStatus();
      const folderByTraining = new Map<string, { id: string; name: string }>();
      const completionRecordUpdates = buildCompletionRecordUpdates(completionRows, rosterRows, rosterHeaders);
      if (completionRecordUpdates.length) {
        setIssueProgress({
          current: 0,
          total: eligibleResults.length,
          label: "전체명단 이수 기록 중",
        });
        await batchUpdateSheet(completionRecordUpdates);
      }

      for (const [index, result] of eligibleResults.entries()) {
        const trainingName = result.completion.trainingName;
        const trainingId = `${trainingName}_${issuedDate}`;
        const logPersonKey = `${result.completion.name}_${result.completion.phone}`;
        let folder = folderByTraining.get(trainingName);
        let uploaded: Awaited<ReturnType<typeof uploadPdfToDrive>> | null = null;

        try {
          setIssueProgress({
            current: index,
            total: eligibleResults.length,
            label: `${result.completion.name} 영수증 만드는 중`,
          });
          if (!folder) {
            folder = await createDriveTrainingFolder(trainingName, issuedDate);
            folderByTraining.set(trainingName, folder);
          }

          const pdfBytes = await generateReceiptPdf({
            completion: result.completion,
            roster: result.roster,
            issuedDate,
          });
          setIssueProgress({
            current: index,
            total: eligibleResults.length,
            label: `${result.completion.name} 영수증 저장 중`,
          });
          uploaded = await uploadPdfToDrive(folder.id, result.receiptFilename, Array.from(pdfBytes));
          const nextIssueCount = result.roster.issueCount + 1;
          setIssueProgress({
            current: index,
            total: eligibleResults.length,
            label: `${result.completion.name} 전체명단 기록 중`,
          });
          await batchUpdateSheet([
            {
              range: `E${result.roster.rowNumber}`,
              values: [[String(nextIssueCount)]],
            },
            {
              range:
                result.nextSlot === 1
                  ? `F${result.roster.rowNumber}:H${result.roster.rowNumber}`
                  : `I${result.roster.rowNumber}:K${result.roster.rowNumber}`,
              values: [[trainingName, issuedDate, uploaded.web_view_link]],
            },
          ]);
          await writeJobLog({
            runId,
            trainingId,
            personKey: logPersonKey,
            sheetRow: result.roster.rowNumber,
            uploaded,
            sheetUpdated: true,
            status: "completed",
          });
          setIssueProgress({
            current: index + 1,
            total: eligibleResults.length,
            label: `${index + 1}/${eligibleResults.length}건 완료`,
          });
        } catch (itemError) {
          await writeJobLog({
            runId,
            trainingId,
            personKey: logPersonKey,
            sheetRow: result.roster.rowNumber,
            uploaded,
            sheetUpdated: false,
            status: "failed",
            errorMessage: itemError instanceof Error ? itemError.message : String(itemError),
          });
          throw itemError;
        }
      }
      await loadRosterFromGoogle();
      setIssueCompletion({
        count: eligibleResults.length,
        rosterUrl: buildGoogleSheetUrl(googleStatus.spreadsheet_id),
        folderLinks: Array.from(folderByTraining.values()).map((folder) => ({
          name: folder.name,
          url: buildDriveFolderUrl(folder.id),
        })),
      });
      setIssueProgress(null);
    } catch (workflowError) {
      setError(workflowError instanceof Error ? workflowError.message : String(workflowError));
    } finally {
      setBusy(false);
    }
  }

  const isAuthenticated = Boolean(googleStatus?.authenticated);

  return (
    <main className="app-shell">
      <Header authenticated={isAuthenticated} onLogout={disconnectGoogle} />

      {error ? (
        <section className="alert error">
          <AlertCircle size={18} />
          {error}
        </section>
      ) : null}
      {notice ? (
        <section className="alert info">
          <CheckCircle2 size={18} />
          {notice}
        </section>
      ) : null}

      {!isAuthenticated ? (
        <LoginGate busy={busy} status={googleStatus} onLogin={connectGoogle} />
      ) : activeTask === null ? (
        <TaskMenu onChoose={chooseTask} />
      ) : activeTask === "completion" ? (
        <CompletionDocumentStage
          busy={busy}
          sourceFile={documentSourceFile}
          sourceRows={documentRows}
          form={documentForm}
          issues={documentIssues}
          previewRows={documentPreviewRows}
          people={documentPeople}
          captureRows={captureRows}
          zoomRows={zoomRows}
          summaryRows={summaryRows}
          captureEvidenceRows={captureEvidenceRows}
          zoomChatFile={zoomChatFile}
          zoomFile={zoomFile}
          evaluationSheetForm={evaluationSheetForm}
          evaluationSummary={evaluationSummary}
          workflowStep={documentWorkflowStep}
          currentDocument={currentDocument}
          currentDocumentDownloaded={currentDocumentDownloaded}
          sheetForm={completionSheetForm}
          onBack={returnToMenu}
          onSheetFormChange={setCompletionSheetForm}
          onFormChange={updateCompletionDocumentForm}
          onLoad={loadCompletionDocumentSource}
          onCreateDocument={createCurrentDocument}
          onMoveStep={moveToDocumentStep}
          onCaptureEvidenceRowChange={updateCaptureEvidenceRow}
          onCaptureEvidenceImagesChange={loadCaptureEvidenceImages}
          onAddCaptureEvidenceRow={addCaptureEvidenceRow}
          onRemoveCaptureEvidenceRow={removeCaptureEvidenceRow}
          onCaptureRowChange={updateCaptureRow}
          onZoomChatFileChange={loadZoomChatFile}
          onZoomFileChange={loadZoomFile}
          onZoomRowChange={updateZoomRow}
          onEvaluationSheetFormChange={setEvaluationSheetForm}
          onLoadEvaluationSheet={loadEvaluationSheet}
          onDocumentDownloaded={markCurrentDocumentDownloaded}
        />
      ) : (
        <section className="workflow-stage">
          <button className="back-button" type="button" onClick={returnToMenu} disabled={busy}>
            <ArrowLeft size={16} />
            메뉴로 돌아가기
          </button>
          <Progress current={step} configured={Boolean(googleStatus?.configured)} />
          {step === "settings" ? (
            <StepCard
              icon={<Settings />}
              title="기본 설정"
              description="전체명단 주소와 영수증 저장 폴더 주소를 한 번만 저장합니다."
            >
              <GoogleUrlSettingsForm value={googleForm} busy={busy} onChange={setGoogleForm} />
              <button className="primary-action" type="button" onClick={saveBaseSettings} disabled={busy}>
                기본 설정 저장
              </button>
            </StepCard>
          ) : null}

          {step === "upload" ? (
            <StepCard
              icon={<FileSpreadsheet />}
              title="이수자 명단 불러오기"
              description="이수자 명단 주소를 입력하면 저장된 전체명단과 자동으로 대조합니다."
            >
              <CompletionSheetUrlForm value={completionSheetForm} busy={busy} onChange={setCompletionSheetForm} />
              <button
                className="primary-action"
                type="button"
                onClick={loadCompletionFromGoogleSheet}
                disabled={busy || !completionSheetForm.spreadsheetUrl.trim()}
              >
                이수자 명단 불러오기
              </button>
              <FileStatus title="이수자 명단" state={completionFile} />
              <button className="secondary-action" type="button" onClick={() => setStep("settings")} disabled={busy}>
                기본 설정 수정
              </button>
            </StepCard>
          ) : null}

          {step === "review" ? (
            <StepCard
              icon={<ReceiptText />}
              title="검토"
              description="전체명단과 대조한 결과를 확인하고, 수동 확인 대상이 없을 때 발급 화면으로 이동합니다."
            >
              <SummaryGrid summary={summary} />
              <FileStatus title="전체명단" state={rosterFile} />
              <IssueList issues={rosterIssues} />
              <ResultTable results={results} />
              <div className="action-row">
                <button className="secondary-action" type="button" onClick={returnToUploadFromReview} disabled={busy}>
                  이전 단계로 돌아가기
                </button>
                <button className="secondary-action" type="button" onClick={reloadRosterForReview} disabled={busy}>
                  전체명단 다시 불러오기
                </button>
                <button className="primary-action" type="button" onClick={confirmReview} disabled={busy || summary.manual > 0}>
                  이상없음 확인
                </button>
              </div>
            </StepCard>
          ) : null}

          {step === "issue" ? (
            <StepCard
              icon={<FolderUp />}
              title="영수증 저장 및 기록"
              description="첫 번째 대상자의 샘플 영수증을 확인한 뒤 전체 영수증 저장과 전체명단 기록을 실행합니다."
            >
              {issueCompletion ? (
                <IssueCompletionPanel completion={issueCompletion} onReset={resetToUpload} />
              ) : (
                <>
                  <div className="notice">
                    <UploadCloud size={18} />
                    저장된 영수증은 링크가 있는 사용자만 볼 수 있도록 공유됩니다.
                  </div>
                  {issueProgress ? <IssueProgress progress={issueProgress} /> : null}
                  <SamplePreview samplePdfUrl={samplePdfUrl} firstName={eligibleResults[0]?.completion.name} />
                  <div className="action-row">
                    <button className="secondary-action" type="button" onClick={() => setStep("review")} disabled={busy}>
                      검토로 돌아가기
                    </button>
                    <button
                      className="primary-action"
                      type="button"
                      onClick={generateUploadAndRecord}
                      disabled={busy || !samplePdfUrl}
                    >
                      샘플 이상 없음 · 전체 저장 시작
                    </button>
                  </div>
                </>
              )}
            </StepCard>
          ) : null}
        </section>
      )}
    </main>
  );
}

function Header({ authenticated, onLogout }: { authenticated: boolean; onLogout: () => void }) {
  return (
    <header className="topbar">
      <div className="brand">
        <img src="/logo.png" alt="" />
        <div>
          <h1>직무연수 영수증 발급 관리</h1>
          <p>{authenticated ? "명단 확인, 영수증 미리보기, 저장 기록을 단계별로 처리합니다." : "구글 로그인 후 업무 화면을 시작합니다."}</p>
        </div>
      </div>
      <div className="security-chip">
        <LockKeyhole size={16} />
        {authenticated ? "구글 로그인됨" : "구글 로그인 필요"}
        {authenticated ? (
          <button className="logout-button" type="button" onClick={onLogout} title="구글 로그아웃">로그아웃</button>
        ) : null}
      </div>
    </header>
  );
}

function LoginGate({
  busy,
  status,
  onLogin,
}: {
  busy: boolean;
  status: GoogleConfigStatus | null;
  onLogin: () => void;
}) {
  return (
    <section className="login-gate">
      <button className="primary-action login-button" type="button" onClick={onLogin} disabled={busy || !status?.client_id}>
        구글 로그인
      </button>
      {!status?.client_id ? <p className="muted">관리자에게 구글 로그인 설정을 요청하세요.</p> : null}
    </section>
  );
}

function TaskMenu({ onChoose }: { onChoose: (task: Exclude<ActiveTask, null>) => void }) {
  return (
    <section className="task-menu">
      <button className="task-button" type="button" onClick={() => onChoose("completion")}>
        <FileSpreadsheet size={26} />
        <span>이수 서류 작성</span>
      </button>
      <button className="task-button" type="button" onClick={() => onChoose("receipt")}>
        <ReceiptText size={26} />
        <span>영수증 발급</span>
      </button>
    </section>
  );
}

function CompletionDocumentStage({
  busy,
  sourceFile,
  sourceRows,
  form,
  issues,
  previewRows,
  people,
  captureRows,
  zoomRows,
  summaryRows,
  captureEvidenceRows,
  zoomChatFile,
  zoomFile,
  evaluationSheetForm,
  evaluationSummary,
  workflowStep,
  currentDocument,
  currentDocumentDownloaded,
  sheetForm,
  onBack,
  onSheetFormChange,
  onFormChange,
  onLoad,
  onCreateDocument,
  onMoveStep,
  onCaptureEvidenceRowChange,
  onCaptureEvidenceImagesChange,
  onAddCaptureEvidenceRow,
  onRemoveCaptureEvidenceRow,
  onCaptureRowChange,
  onZoomChatFileChange,
  onZoomFileChange,
  onZoomRowChange,
  onEvaluationSheetFormChange,
  onLoadEvaluationSheet,
  onDocumentDownloaded,
}: {
  busy: boolean;
  sourceFile: LoadedFileState | null;
  sourceRows: ReturnType<typeof normalizeApplicantRows>;
  form: AttendanceBaseForm;
  issues: ValidationIssue[];
  previewRows: CompletionDocumentRow[];
  people: AttendancePerson[];
  captureRows: CaptureAttendanceRow[];
  zoomRows: ZoomAttendanceRow[];
  summaryRows: SummaryAttendanceRow[];
  captureEvidenceRows: CaptureEvidenceRow[];
  zoomChatFile: LoadedFileState | null;
  zoomFile: LoadedFileState | null;
  evaluationSheetForm: EvaluationSheetForm;
  evaluationSummary: EvaluationSummary | null;
  workflowStep: DocumentWorkflowStep;
  currentDocument: GeneratedDocument | null;
  currentDocumentDownloaded: boolean;
  sheetForm: CompletionSheetForm;
  onBack: () => void;
  onSheetFormChange: (value: CompletionSheetForm) => void;
  onFormChange: (value: AttendanceBaseForm) => void;
  onLoad: () => void;
  onCreateDocument: () => void;
  onMoveStep: (step: DocumentWorkflowStep) => void;
  onCaptureEvidenceRowChange: (id: string, patch: Partial<Pick<CaptureEvidenceRow, "period" | "mode">>) => void;
  onCaptureEvidenceImagesChange: (id: string, kind: "camera" | "chat", files: FileList | null) => void;
  onAddCaptureEvidenceRow: () => void;
  onRemoveCaptureEvidenceRow: (id: string) => void;
  onCaptureRowChange: (index: number, patch: Partial<Pick<CaptureAttendanceRow, "period1" | "period2" | "result">>) => void;
  onZoomChatFileChange: (file: File | null) => void;
  onZoomFileChange: (file: File | null) => void;
  onZoomRowChange: (index: number, patch: Partial<Pick<ZoomAttendanceRow, "result">>) => void;
  onEvaluationSheetFormChange: (value: EvaluationSheetForm) => void;
  onLoadEvaluationSheet: () => void;
  onDocumentDownloaded: () => void;
}) {
  const trainingNameOptions = getTrainingNameOptions(sourceRows);
  const nextStep = getNextDocumentStep(workflowStep);

  return (
    <section className="workflow-stage">
      <button className="back-button" type="button" onClick={onBack} disabled={busy}>
        <ArrowLeft size={16} />
        메뉴로 돌아가기
      </button>
      <StepCard
        icon={<FileSpreadsheet />}
        title="이수 서류 작성"
        description="연수 신청자 명단, 줌 접속기록, 평가 결과를 기준으로 제출용 문서를 만듭니다."
      >
        <CompletionSheetUrlForm value={sheetForm} busy={busy} onChange={onSheetFormChange} />
        <button className="primary-action" type="button" onClick={onLoad} disabled={busy || !sheetForm.spreadsheetUrl.trim()}>
          연수 신청자 명단 불러오기
        </button>
        <FileStatus title="연수 신청자 명단" state={sourceFile} />
        {sourceRows.length ? (
          <>
            <IssueList issues={issues} emptyMessage="나이스번호 형식 오류가 없습니다." />
            <DocumentStepProgress current={workflowStep} />
            <CompletionDocumentFormView
              value={form}
              trainingNameOptions={trainingNameOptions}
              busy={busy}
              onChange={onFormChange}
            />
            {workflowStep === "doc1" ? (
              <CaptureControls
                rows={captureRows}
                evidenceRows={captureEvidenceRows}
                zoomChatFile={zoomChatFile}
                busy={busy}
                onEvidenceRowChange={onCaptureEvidenceRowChange}
                onEvidenceImagesChange={onCaptureEvidenceImagesChange}
                onAddEvidenceRow={onAddCaptureEvidenceRow}
                onRemoveEvidenceRow={onRemoveCaptureEvidenceRow}
                onRowChange={onCaptureRowChange}
                onZoomChatFileChange={onZoomChatFileChange}
              />
            ) : null}
            {workflowStep === "doc2" ? (
              <ZoomControls
                busy={busy}
                zoomFile={zoomFile}
                rows={zoomRows}
                onFileChange={onZoomFileChange}
                onRowChange={onZoomRowChange}
              />
            ) : null}
            {workflowStep === "doc3" ? <SummaryAttendancePreview rows={summaryRows} /> : null}
            {workflowStep === "doc5" ? (
              <>
                <EvaluationControls
                  busy={busy}
                  value={evaluationSheetForm}
                  summary={evaluationSummary}
                  onChange={onEvaluationSheetFormChange}
                  onLoad={onLoadEvaluationSheet}
                />
                <EvaluationPreview summary={evaluationSummary} />
              </>
            ) : null}
            {workflowStep === "doc7" ? <CompletionPreviewTable rows={previewRows} /> : null}
            {workflowStep === "done" ? <p className="notice">5개 문서 저장을 완료했습니다.</p> : (
              <DocumentStepActions
                busy={busy}
                step={workflowStep}
                document={currentDocument}
                downloaded={currentDocumentDownloaded}
                nextStep={nextStep}
                onCreate={onCreateDocument}
                onDownloaded={onDocumentDownloaded}
                onNext={() => onMoveStep(nextStep ?? "done")}
              />
            )}
          </>
        ) : null}
      </StepCard>
    </section>
  );
}

function DocumentStepProgress({ current }: { current: DocumentWorkflowStep }) {
  const items: Array<{ id: DocumentWorkflowStep; label: string }> = [
    { id: "doc1", label: "1번 캡처" },
    { id: "doc2", label: "2번 입퇴장" },
    { id: "doc3", label: "3번 종합" },
    { id: "doc5", label: "5번 평가" },
    { id: "doc7", label: "7번 명단" },
  ];
  return (
    <nav className="document-stepper" aria-label="이수 서류 작성 단계">
      {items.map((item) => (
        <span key={item.id} className={current === item.id ? "active" : ""}>{item.label}</span>
      ))}
    </nav>
  );
}

function DocumentStepActions({
  busy,
  step,
  document,
  downloaded,
  nextStep,
  onCreate,
  onDownloaded,
  onNext,
}: {
  busy: boolean;
  step: DocumentWorkflowStep;
  document: GeneratedDocument | null;
  downloaded: boolean;
  nextStep: DocumentWorkflowStep | null;
  onCreate: () => void;
  onDownloaded: () => void;
  onNext: () => void;
}) {
  return (
    <div className="document-actions">
      <button className="primary-action" type="button" onClick={onCreate} disabled={busy || step === "done"}>
        {documentStepLabel(step)} 파일 만들기
      </button>
      {document ? (
        <button className="download-action" type="button" onClick={() => saveDocumentAs(document).then(onDownloaded)}>
          <Download size={17} />
          {document.label} 저장
        </button>
      ) : null}
      {nextStep ? (
        <button className="secondary-action" type="button" onClick={onNext} disabled={busy || !downloaded}>
          다음 문서로 이동
        </button>
      ) : (
        <button className="secondary-action" type="button" onClick={onNext} disabled={busy || !downloaded}>
          완료
        </button>
      )}
    </div>
  );
}

function getNextDocumentStep(step: DocumentWorkflowStep): DocumentWorkflowStep | null {
  if (step === "doc1") return "doc2";
  if (step === "doc2") return "doc3";
  if (step === "doc3") return "doc5";
  if (step === "doc5") return "doc7";
  if (step === "doc7") return null;
  return null;
}

function documentStepLabel(step: DocumentWorkflowStep) {
  if (step === "doc1") return "1번 문서";
  if (step === "doc2") return "2번 문서";
  if (step === "doc3") return "3번 문서";
  if (step === "doc5") return "5번 문서";
  if (step === "doc7") return "7번 문서";
  return "문서";
}

function CompletionDocumentFormView({
  value,
  trainingNameOptions,
  busy,
  onChange,
}: {
  value: AttendanceBaseForm;
  trainingNameOptions: string[];
  busy: boolean;
  onChange: (value: AttendanceBaseForm) => void;
}) {
  return (
    <div className="settings-form document-form">
      <label>
        <span>연수과정명</span>
        <input
          value={value.trainingName}
          onChange={(event) => onChange({ ...value, trainingName: event.target.value })}
          list="training-name-options"
          disabled={busy}
        />
        <datalist id="training-name-options">
          {trainingNameOptions.map((name) => (
            <option key={name} value={name} />
          ))}
        </datalist>
      </label>
      <label>
        <span>연수기관</span>
        <input
          value={value.institute}
          onChange={(event) => onChange({ ...value, institute: event.target.value })}
          disabled={busy}
        />
      </label>
      <label>
        <span>연수날짜</span>
        <input
          type="date"
          value={value.trainingDate}
          onChange={(event) => onChange({ ...value, trainingDate: event.target.value })}
          disabled={busy}
        />
      </label>
      <label>
        <span>연수 총 시간</span>
        <input
          value={value.totalHours}
          onChange={(event) => onChange({ ...value, totalHours: event.target.value })}
          inputMode="numeric"
          disabled={busy}
        />
      </label>
      <label>
        <span>연수 시작시간</span>
        <input
          value={value.startTime}
          onChange={(event) => onChange({ ...value, startTime: event.target.value })}
          placeholder="HH:mm"
          disabled={busy}
        />
      </label>
      <label>
        <span>연수 종료시간</span>
        <input
          value={value.endTime}
          onChange={(event) => onChange({ ...value, endTime: event.target.value })}
          placeholder="HH:mm"
          disabled={busy}
        />
      </label>
      <label>
        <span>1교시 라벨</span>
        <input
          value={value.period1Label}
          onChange={(event) => onChange({ ...value, period1Label: event.target.value })}
          placeholder="1교시 13:00부터 13:50까지"
          disabled={busy}
        />
      </label>
      <label>
        <span>2교시 라벨</span>
        <input
          value={value.period2Label}
          onChange={(event) => onChange({ ...value, period2Label: event.target.value })}
          placeholder="2교시 14:00부터 14:50까지"
          disabled={busy}
        />
      </label>
      <label>
        <span>강사명</span>
        <input
          value={value.instructorName}
          onChange={(event) => onChange({ ...value, instructorName: event.target.value })}
          disabled={busy}
        />
      </label>
    </div>
  );
}

function CompletionPreviewTable({ rows }: { rows: CompletionDocumentRow[] }) {
  if (!rows.length) {
    return <p className="muted">서류 만들기를 클릭하면 생성될 엑셀 명단 내용이 미리보기로 표시됩니다.</p>;
  }

  return (
    <div className="table-wrap document-preview">
      <table>
        <thead>
          <tr>
            <th>연번</th>
            <th>나이스개인번호</th>
            <th>연수과정</th>
            <th>연수기관</th>
            <th>연수시작일</th>
            <th>연수종료일</th>
            <th>연수구분</th>
            <th>교육유형구분코드</th>
            <th>연수시간</th>
            <th>성적</th>
            <th>직무관련성</th>
            <th>평점학점</th>
            <th>이수번호</th>
            <th>성명</th>
            <th>생년월일</th>
            <th>학교명</th>
            <th>초/중등</th>
            <th>연수분류코드</th>
            <th>합격증번호</th>
            <th>법정의무여부</th>
            <th>법정의무코드</th>
            <th>교육형태코드</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={`${row.sequence}-${row.niceNumber}-${row.name}`}>
              <td>{row.sequence}</td>
              <td>{row.niceNumber}</td>
              <td>{row.trainingName}</td>
              <td>{row.institute}</td>
              <td>{row.startDate}</td>
              <td>{row.endDate}</td>
              <td>{row.trainingCategory}</td>
              <td>{row.educationTypeCode}</td>
              <td>{row.totalHours}</td>
              <td>{row.score}</td>
              <td>{row.jobRelated}</td>
              <td>{row.credit}</td>
              <td>{row.completionNumber}</td>
              <td>{row.name}</td>
              <td>{row.birthDate}</td>
              <td>{row.schoolName}</td>
              <td>{row.schoolLevel}</td>
              <td>{row.trainingClassCode}</td>
              <td>{row.certificateNumber}</td>
              <td>{row.mandatoryYn}</td>
              <td>{row.mandatoryCode}</td>
              <td>{row.educationFormatCode}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CaptureControls({
  rows,
  evidenceRows,
  zoomChatFile,
  busy,
  onEvidenceRowChange,
  onEvidenceImagesChange,
  onAddEvidenceRow,
  onRemoveEvidenceRow,
  onRowChange,
  onZoomChatFileChange,
}: {
  rows: CaptureAttendanceRow[];
  evidenceRows: CaptureEvidenceRow[];
  zoomChatFile: LoadedFileState | null;
  busy: boolean;
  onEvidenceRowChange: (id: string, patch: Partial<Pick<CaptureEvidenceRow, "period" | "mode">>) => void;
  onEvidenceImagesChange: (id: string, kind: "camera" | "chat", files: FileList | null) => void;
  onAddEvidenceRow: () => void;
  onRemoveEvidenceRow: (id: string) => void;
  onRowChange: (index: number, patch: Partial<Pick<CaptureAttendanceRow, "period1" | "period2" | "result">>) => void;
  onZoomChatFileChange: (file: File | null) => void;
}) {
  return (
    <section className="sub-panel">
      <h3>1. 화면 캡처 출결자료</h3>
      <label className="file-picker">
        줌 채팅기록 불러오기
        <input type="file" accept=".txt,text/plain" onChange={(event) => onZoomChatFileChange(event.target.files?.[0] ?? null)} disabled={busy} />
      </label>
      <FileStatus title="줌 채팅기록" state={zoomChatFile} />
      <div className="evidence-list">
        {evidenceRows.map((evidenceRow, index) => (
          <div className="evidence-row" key={evidenceRow.id}>
            <div className="evidence-controls">
              <strong>증빙 행 {index + 1}</strong>
              <select
                value={evidenceRow.period}
                onChange={(event) => onEvidenceRowChange(evidenceRow.id, { period: Number(event.target.value) as 1 | 2 })}
                disabled={busy}
              >
                <option value={1}>1교시</option>
                <option value={2}>2교시</option>
              </select>
              <select
                value={evidenceRow.mode}
                onChange={(event) => onEvidenceRowChange(evidenceRow.id, { mode: event.target.value as CaptureMode })}
                disabled={busy}
              >
                <option value="camera">캠화면</option>
                <option value="chat">채팅화면</option>
              </select>
              <label className="file-picker">
                {evidenceRow.mode === "camera" ? "캠 이미지 1개" : "채팅 이미지 최대 4개"}
                <input
                  type="file"
                  accept="image/*"
                  multiple={evidenceRow.mode === "chat"}
                  onChange={(event) => onEvidenceImagesChange(evidenceRow.id, evidenceRow.mode, event.target.files)}
                  disabled={busy}
                />
              </label>
              <button className="secondary-action compact-button" type="button" onClick={() => onRemoveEvidenceRow(evidenceRow.id)} disabled={busy || evidenceRows.length <= 1}>
                삭제
              </button>
            </div>
            <EvidenceImagePreview row={evidenceRow} />
          </div>
        ))}
        <button className="secondary-action add-row-button" type="button" onClick={onAddEvidenceRow} disabled={busy}>
          + 증빙 행 추가
        </button>
      </div>
      <div className="table-wrap compact-preview">
        <table>
          <thead>
            <tr>
              <th>연번</th>
              <th>성명</th>
              <th>1교시</th>
              <th>2교시</th>
              <th>결과</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => (
              <tr key={`${row.sequence}-${row.name}`}>
                <td>{row.sequence}</td>
                <td>{row.name}</td>
                <td><SelectOX value={row.period1} busy={busy} onChange={(value) => onRowChange(index, { period1: value })} /></td>
                <td><SelectOX value={row.period2} busy={busy} onChange={(value) => onRowChange(index, { period2: value })} /></td>
                <td>{row.result}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function EvidenceImagePreview({ row }: { row: CaptureEvidenceRow }) {
  const [detailed, setDetailed] = useState(false);
  const images = row.mode === "camera" ? (row.cameraImage ? [row.cameraImage] : []) : row.chatImages;
  if (!images.length) return <p className="muted">첨부된 이미지가 없습니다.</p>;
  return (
    <div>
      <div className="image-preview-toggle">
        <button
          type="button"
          className={!detailed ? "active" : ""}
          onClick={() => setDetailed(false)}
        >간략히</button>
        <button
          type="button"
          className={detailed ? "active" : ""}
          onClick={() => setDetailed(true)}
        >자세히</button>
      </div>
      {detailed ? (
        <div className="image-preview-grid">
          {images.map((image, index) => (
            <figure className="image-preview-card" key={`${row.id}-${image.name}-${index}`}>
              <img src={image.dataUrl} alt={`${row.period}교시 ${image.name}`} />
              <figcaption>{row.period}교시 · {row.mode === "camera" ? "캠" : `채팅 ${index + 1}`} · {image.name}</figcaption>
            </figure>
          ))}
        </div>
      ) : (
        <ul className="image-preview-brief">
          {images.map((image, index) => (
            <li key={`${row.id}-${image.name}-${index}`}>
              <img src={image.dataUrl} alt="" />
              <span>{image.name}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function SelectOX({ value, busy, onChange }: { value: "O" | "X"; busy: boolean; onChange: (value: "O" | "X") => void }) {
  const nextValue = value === "O" ? "X" : "O";
  return (
    <button
      type="button"
      className={`ox-toggle ${value === "O" ? "on" : "off"}`}
      onClick={() => onChange(nextValue)}
      disabled={busy}
      aria-label={`${value}에서 ${nextValue}로 변경`}
    >
      {value}
    </button>
  );
}

function ZoomControls({
  busy,
  zoomFile,
  rows,
  onFileChange,
  onRowChange,
}: {
  busy: boolean;
  zoomFile: LoadedFileState | null;
  rows: ZoomAttendanceRow[];
  onFileChange: (file: File | null) => void;
  onRowChange: (index: number, patch: Partial<Pick<ZoomAttendanceRow, "result">>) => void;
}) {
  return (
    <section className="sub-panel">
      <h3>2. 입장 및 퇴장 시간 기록자료</h3>
      <label className="file-picker">
        줌 접속기록 파일 불러오기
        <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => onFileChange(event.target.files?.[0] ?? null)} disabled={busy} />
      </label>
      <FileStatus title="줌 접속기록" state={zoomFile} />
      {rows.length ? (
        <div className="table-wrap compact-preview">
          <table>
            <thead>
              <tr>
                <th>성명</th>
                <th>입장</th>
                <th>퇴장</th>
                <th>인정분</th>
                <th>결과</th>
                <th>경고</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={`${row.sequence}-${row.name}`}>
                  <td>{row.name}</td>
                  <td>{row.entryTime || "-"}</td>
                  <td>{row.exitTime || "-"}</td>
                  <td>{row.effectiveMinutes}</td>
                  <td>
                    <select
                      value={row.result}
                      onChange={(event) => onRowChange(index, { result: event.target.value as "인정" | "미인정" })}
                      disabled={busy}
                    >
                      <option value="인정">인정</option>
                      <option value="미인정">미인정</option>
                    </select>
                  </td>
                  <td>{row.warning || ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </section>
  );
}

function EvaluationControls({
  busy,
  value,
  summary,
  onChange,
  onLoad,
}: {
  busy: boolean;
  value: EvaluationSheetForm;
  summary: EvaluationSummary | null;
  onChange: (value: EvaluationSheetForm) => void;
  onLoad: () => void;
}) {
  return (
    <section className="sub-panel">
      <h3>5. 종합평가서</h3>
      <div className="settings-form">
        <label>
          <span>평가 결과 주소</span>
          <input
            value={value.spreadsheetUrl}
            onChange={(event) => onChange({ spreadsheetUrl: event.target.value })}
            placeholder="구글 스프레드시트 주소를 붙여넣으세요"
            disabled={busy}
          />
        </label>
      </div>
      <button className="secondary-action" type="button" onClick={onLoad} disabled={busy || !value.spreadsheetUrl.trim()}>
        평가 결과 불러오기
      </button>
      {summary ? (
        <p className="muted">응답자 {summary.respondentCount.toLocaleString()}명, 평균값 {summary.averages.filter(Boolean).length}/11개 확인</p>
      ) : null}
    </section>
  );
}

function SummaryAttendancePreview({ rows }: { rows: SummaryAttendanceRow[] }) {
  if (!rows.length) return null;
  const counts = countSummary(rows);
  return (
    <section className="sub-panel">
      <h3>3. 종합 출결자료</h3>
      <p className="muted">이수자 {counts.completionCount.toLocaleString()}명 / 미이수자 {counts.incompleteCount.toLocaleString()}명</p>
      <div className="table-wrap compact-preview">
        <table>
          <thead>
            <tr>
              <th>연번</th>
              <th>성명</th>
              <th>필수 결과</th>
              <th>선택 결과</th>
              <th>이수 여부</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${row.sequence}-${row.name}`}>
                <td>{row.sequence}</td>
                <td>{row.name}</td>
                <td>{row.result1}</td>
                <td>{row.result2}</td>
                <td>{row.result3}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function EvaluationPreview({ summary }: { summary: EvaluationSummary | null }) {
  if (!summary) return null;
  return (
    <section className="sub-panel">
      <h3>5. 종합평가서 미리보기</h3>
      <p className="muted">응답자 {summary.respondentCount.toLocaleString()}명</p>
      <div className="table-wrap compact-preview">
        <table>
          <thead>
            <tr>
              {summary.averages.map((_, index) => (
                <th key={`average-heading-${index}`}>응답_{index + 1}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr>
              {summary.averages.map((value, index) => (
                <td key={`average-value-${index}`}>{value}</td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
      <div className="opinion-preview">
        <div>
          <strong>종합의견_1</strong>
          <pre>{summary.opinion1 || "내용 없음"}</pre>
        </div>
        <div>
          <strong>종합의견_2</strong>
          <pre>{summary.opinion2 || "내용 없음"}</pre>
        </div>
      </div>
    </section>
  );
}

function GeneratedDocumentLinks({
  documents,
  downloadedCount,
  onDownloaded,
}: {
  documents: GeneratedDocument[];
  downloadedCount: number;
  onDownloaded: (index: number) => void;
}) {
  if (!documents.length) return null;
  return (
    <div className="download-list">
      {documents.map((document, index) => {
        const available = index <= downloadedCount;
        return available ? (
          <button
            key={document.filename}
            className="download-action"
            type="button"
            onClick={() => saveDocumentAs(document).then(() => onDownloaded(index))}
          >
            <Download size={17} />
            {index + 1}. {document.label} 저장
          </button>
        ) : (
          <button key={document.filename} className="download-action locked" type="button" disabled>
            <Download size={17} />
            {index + 1}. 이전 문서를 저장한 뒤 가능
          </button>
        );
      })}
    </div>
  );
}

function Progress({ current, configured }: { current: WorkflowStep; configured: boolean }) {
  const items: Array<{ id: WorkflowStep; label: string }> = [
    { id: "settings", label: "기본 설정" },
    { id: "upload", label: "명단 불러오기" },
    { id: "review", label: "검토" },
    { id: "issue", label: "기록" },
  ];

  return (
    <nav className="step-progress" aria-label="작업 단계">
      {items.map((item) => (
        <button
          key={item.id}
          type="button"
          className={current === item.id ? "active" : ""}
          disabled={item.id !== "settings" && !configured}
        >
          {item.label}
        </button>
      ))}
    </nav>
  );
}

function GoogleUrlSettingsForm({
  value,
  busy,
  onChange,
}: {
  value: GoogleUrlForm;
  busy: boolean;
  onChange: (value: GoogleUrlForm) => void;
}) {
  return (
    <div className="settings-form">
      <label>
        <span>전체명단 주소</span>
        <input
          value={value.spreadsheetUrl}
          onChange={(event) => onChange({ ...value, spreadsheetUrl: event.target.value })}
          placeholder="구글 스프레드시트 주소를 붙여넣으세요"
          disabled={busy}
        />
      </label>
      <label>
        <span>영수증 저장 폴더 주소</span>
        <input
          value={value.driveFolderUrl}
          onChange={(event) => onChange({ ...value, driveFolderUrl: event.target.value })}
          placeholder="구글 드라이브 폴더 주소를 붙여넣으세요"
          disabled={busy}
        />
      </label>
    </div>
  );
}

function CompletionSheetUrlForm({
  value,
  busy,
  onChange,
}: {
  value: CompletionSheetForm;
  busy: boolean;
  onChange: (value: CompletionSheetForm) => void;
}) {
  return (
    <div className="settings-form">
      <label>
        <span>연수 신청자 명단 주소</span>
        <input
          value={value.spreadsheetUrl}
          onChange={(event) => onChange({ spreadsheetUrl: event.target.value })}
          placeholder="구글 스프레드시트 주소를 붙여넣으세요"
          disabled={busy}
        />
      </label>
      <p className="muted">주소에 특정 탭 정보가 있으면 해당 탭을, 없으면 첫 번째 탭을 읽습니다.</p>
    </div>
  );
}

function StepCard({
  icon,
  title,
  description,
  children,
}: {
  icon: React.ReactNode;
  title: string;
  description: string;
  children: React.ReactNode;
}) {
  return (
    <section className="step-card">
      <div className="step-heading">
        <span className="step-icon">{icon}</span>
        <div>
          <h2>{title}</h2>
          <p>{description}</p>
        </div>
      </div>
      <div className="step-body">{children}</div>
    </section>
  );
}

function FileStatus({ title, state }: { title: string; state: LoadedFileState | null }) {
  if (!state) {
    return <p className="muted">{title}을 아직 불러오지 않았습니다.</p>;
  }

  return (
    <div className={state.missingHeaders.length ? "file-status blocked" : "file-status ok"}>
      <strong>{state.name}</strong>
      <span>{state.rowCount.toLocaleString()}건</span>
      {state.missingHeaders.length ? (
        <span>누락된 필수 항목: {state.missingHeaders.join(", ")}</span>
      ) : (
        <span>필수 항목 확인 완료</span>
      )}
    </div>
  );
}

function SummaryGrid({ summary }: { summary: { completed: number; eligible: number; manual: number; excluded: number } }) {
  return (
    <div className="summary-grid">
      <Metric label="이수자 명단 인원" value={summary.completed} />
      <Metric label="발급 가능 인원" value={summary.eligible} />
      <Metric label="수동 확인" value={summary.manual} />
      <Metric label="제외" value={summary.excluded} />
    </div>
  );
}

function Metric({ label, value }: { label: string; value: number }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value.toLocaleString()}</strong>
    </div>
  );
}

function IssueList({
  issues,
  emptyMessage = "전체명단의 발급 기록에 문제가 없습니다.",
}: {
  issues: ValidationIssue[];
  emptyMessage?: string;
}) {
  if (!issues.length) {
    return <p className="muted">{emptyMessage}</p>;
  }

  return (
    <div className="issue-list">
      {issues.slice(0, 5).map((issue, index) => (
        <div key={`${issue.row}-${index}`} className="issue-row">
          <AlertCircle size={16} />
          <span>{issue.row ? `${issue.row}행: ` : ""}{issue.message}</span>
        </div>
      ))}
      {issues.length > 5 ? <p className="muted">외 {issues.length - 5}건</p> : null}
    </div>
  );
}

function SamplePreview({ samplePdfUrl, firstName }: { samplePdfUrl: string | null; firstName?: string }) {
  return (
    <div className="sample-preview">
      <div>
        <strong>샘플 대상자</strong>
        <span>{firstName || "샘플 생성 중"}</span>
      </div>
      {samplePdfUrl ? (
        <object data={samplePdfUrl} type="application/pdf" aria-label="첫 번째 대상자 영수증 샘플" />
      ) : (
        <div className="sample-placeholder">샘플 영수증 생성 중</div>
      )}
    </div>
  );
}

function IssueProgress({ progress }: { progress: IssueProgressState }) {
  const percent = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;
  return (
    <div className="issue-progress">
      <div>
        <strong>{progress.current}/{progress.total}</strong>
        <span>{progress.label}</span>
      </div>
      <progress max={progress.total} value={progress.current} />
      <span>{percent}%</span>
    </div>
  );
}

function IssueCompletionPanel({
  completion,
  onReset,
}: {
  completion: IssueCompletionState;
  onReset: () => void;
}) {
  return (
    <div className="completion-panel">
      <div className="completion-heading">
        <CheckCircle2 size={22} />
        <div>
          <h3>저장이 완료되었습니다</h3>
          <p>{completion.count.toLocaleString()}건의 영수증 저장과 전체명단 기록을 완료했습니다.</p>
        </div>
      </div>
      <div className="completion-links">
        <a href={completion.rosterUrl} target="_blank" rel="noreferrer">
          전체명단 열기
        </a>
        {completion.folderLinks.map((folder) => (
          <a key={folder.url} href={folder.url} target="_blank" rel="noreferrer">
            저장 폴더 열기: {folder.name}
          </a>
        ))}
      </div>
      <button className="primary-action" type="button" onClick={onReset}>
        명단 불러오기로 돌아가기
      </button>
    </div>
  );
}

function ResultTable({ results }: { results: ReturnType<typeof matchRecipients> }) {
  if (!results.length) {
    return <p className="muted">이수자 명단을 불러오면 발급 대상자 목록이 표시됩니다.</p>;
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>상태</th>
            <th>성명</th>
            <th>학교명</th>
            <th>연수과정명</th>
            <th>영수증 파일명 또는 제외 사유</th>
          </tr>
        </thead>
        <tbody>
          {results.map((result, index) => (
            <tr key={`${result.completion.rowNumber}-${index}`} data-status={result.status}>
              <td>{statusLabel(result.status)}</td>
              <td>{result.completion.name}</td>
              <td>{result.completion.school}</td>
              <td>{result.completion.trainingName}</td>
              <td>{result.status === "eligible" ? result.receiptFilename : result.reason}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function buildCompletionRecordUpdates(
  completionRows: NormalizedCompletion[],
  rosterRows: NormalizedRoster[],
  rosterHeaders: string[],
): Array<{ range: string; values: string[][] }> {
  if (!completionRows.length || !rosterRows.length) return [];

  const completedByTraining = new Map<string, Set<string>>();
  completionRows.forEach((row) => {
    const trainingName = row.trainingName.trim();
    if (!trainingName) return;
    const completed = completedByTraining.get(trainingName) ?? new Set<string>();
    completed.add(personKey(row.name, row.phone));
    completedByTraining.set(trainingName, completed);
  });

  const trainingNames = Array.from(completedByTraining.keys());
  if (!trainingNames.length) return [];

  const firstRecordColumnIndex = 11; // L, zero-based
  const nextHeaders = [...rosterHeaders];
  const assignedColumns = new Map<string, number>();

  trainingNames.forEach((trainingName) => {
    const existingIndex = nextHeaders.findIndex(
      (header, index) => index >= firstRecordColumnIndex && header.trim() === trainingName,
    );
    if (existingIndex >= 0) {
      assignedColumns.set(trainingName, existingIndex);
      return;
    }

    let nextIndex = firstRecordColumnIndex;
    while ((nextHeaders[nextIndex] ?? "").trim()) nextIndex += 1;
    nextHeaders[nextIndex] = trainingName;
    assignedColumns.set(trainingName, nextIndex);
  });

  return trainingNames.flatMap((trainingName) => {
    const columnIndex = assignedColumns.get(trainingName);
    const completed = completedByTraining.get(trainingName);
    if (columnIndex == null || !completed) return [];

    const column = columnLetter(columnIndex + 1);
    return [
      {
        range: `${column}1`,
        values: [[trainingName]],
      },
      {
        range: `${column}2:${column}${rosterRows.length + 1}`,
        values: rosterRows.map((row) => [completed.has(personKey(row.name, row.phone)) ? "O" : "X"]),
      },
    ];
  });
}

function columnLetter(columnNumber: number): string {
  let value = columnNumber;
  let label = "";
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function extractSpreadsheetId(value: string, label: string): string {
  const trimmed = value.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match?.[1]) return match[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) return trimmed;
  throw new Error(`${label} 주소를 확인하지 못했습니다. 구글 스프레드시트 주소를 다시 확인하세요.`);
}

function extractGoogleSheetGid(value: string): number | undefined {
  const trimmed = value.trim();
  const hashMatch = trimmed.match(/[#&?]gid=(\d+)/);
  if (!hashMatch?.[1]) return undefined;
  return Number.parseInt(hashMatch[1], 10);
}

function isValidTimeText(value: string): boolean {
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(value.trim());
}

function extractDriveFolderId(value: string): string {
  const trimmed = value.trim();
  const folderMatch = trimmed.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (folderMatch?.[1]) return folderMatch[1];
  const queryMatch = trimmed.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (queryMatch?.[1]) return queryMatch[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) return trimmed;
  throw new Error("영수증 저장 폴더 주소를 확인하지 못했습니다. 구글 드라이브 폴더 주소를 다시 확인하세요.");
}

function buildGoogleSheetUrl(spreadsheetId: string): string {
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
}

function buildDriveFolderUrl(folderId: string): string {
  return `https://drive.google.com/drive/folders/${folderId}`;
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(reader.error ?? new Error("이미지를 읽지 못했습니다."));
    reader.readAsDataURL(file);
  });
}

async function resizeImageFileForHwpx(file: File, kind: "camera" | "chat"): Promise<string> {
  const dataUrl = await fileToDataUrl(file);
  const image = await loadHtmlImage(dataUrl);
  const maxWidth = kind === "camera" ? 1600 : 520;
  const maxHeight = kind === "camera" ? 700 : 700;
  const scale = Math.min(1, maxWidth / image.naturalWidth, maxHeight / image.naturalHeight);
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return dataUrl;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);
  return canvas.toDataURL("image/jpeg", 0.88);
}

function loadHtmlImage(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("이미지 크기를 계산하지 못했습니다."));
    image.src = dataUrl;
  });
}

function statusLabel(status: "eligible" | "manual-review" | "excluded") {
  if (status === "eligible") return "발급 가능";
  if (status === "manual-review") return "수동 확인";
  return "제외";
}

async function writeJobLog({
  runId,
  trainingId,
  personKey,
  sheetRow,
  uploaded,
  sheetUpdated,
  status,
  errorMessage,
}: {
  runId: string;
  trainingId: string;
  personKey: string;
  sheetRow: number;
  uploaded: Awaited<ReturnType<typeof uploadPdfToDrive>> | null;
  sheetUpdated: boolean;
  status: "completed" | "failed";
  errorMessage?: string;
}) {
  try {
    await appendJobLog({
      run_id: runId,
      training_id: trainingId,
      person_key: personKey,
      sheet_row: sheetRow,
      drive_file_id: uploaded?.id,
      drive_file_link: uploaded?.web_view_link,
      sheet_updated: sheetUpdated,
      status,
      error_message: errorMessage,
      updated_at: new Date().toISOString(),
    });
  } catch (logError) {
    console.warn("작업 로그를 기록하지 못했습니다.", logError);
  }
}
