import { invoke } from "@tauri-apps/api/core";
import { NormalizedCompletion, NormalizedRoster } from "./workflow";

export type ReceiptPdfInput = {
  completion: NormalizedCompletion;
  roster: NormalizedRoster;
  issuedDate: string;
  trainingPeriod: string;
};

export async function generateReceiptPdf(input: ReceiptPdfInput): Promise<Uint8Array> {
  void input.roster;
  const bytes = await invoke<number[]>("generate_receipt_pdf_from_pptx", {
    request: {
      name: input.completion.name,
      training_name: input.completion.trainingName,
      training_period: input.trainingPeriod,
    },
  });
  return new Uint8Array(bytes);
}

export function todayLocalDate(): string {
  const date = new Date();
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}
