import fontkit from "@pdf-lib/fontkit";
import { PDFDocument, PDFFont, PDFPage, RGB, rgb } from "pdf-lib";
import notoSansGothicUrl from "../assets/NotoSansGothic-Regular.ttf";
import { NormalizedCompletion, NormalizedRoster } from "./workflow";

export type ReceiptPdfInput = {
  completion: NormalizedCompletion;
  roster: NormalizedRoster;
  issuedDate: string;
};

export async function generateReceiptPdf(input: ReceiptPdfInput): Promise<Uint8Array> {
  const [templateBytes, fontBytes] = await Promise.all([
    fetchBinaryAsset("/receipt_sample.pdf", "영수증 양식"),
    fetchBinaryAsset(notoSansGothicUrl, "영수증 글꼴"),
  ]);
  const pdf = await PDFDocument.load(templateBytes);
  pdf.registerFontkit(fontkit);
  const font = await pdf.embedFont(fontBytes, { subset: false });
  const page = pdf.getPage(0);
  const color = rgb(0.08, 0.08, 0.08);

  drawCenteredText(page, font, input.completion.name, {
    x: 128,
    y: 586,
    width: 420,
    height: 28,
    maxSize: 11,
    color,
  });
  drawCenteredText(page, font, input.completion.trainingName, {
    x: 128,
    y: 533,
    width: 202,
    height: 54,
    maxSize: 12,
    color,
  });

  return pdf.save();
}

export function todayLocalDate(): string {
  const date = new Date();
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function draw(
  page: PDFPage,
  font: PDFFont,
  text: string,
  x: number,
  y: number,
  size: number,
  color: RGB,
) {
  page.drawText(text, {
    x,
    y,
    size,
    font,
    color,
  });
}

function drawCenteredText(
  page: PDFPage,
  font: PDFFont,
  text: string,
  rect: { x: number; y: number; width: number; height: number; maxSize: number; color: RGB },
) {
  const cleanText = text.trim();
  if (!cleanText) return;

  const size = fitFontSize(font, cleanText, rect.width - 12, rect.maxSize);
  const textWidth = font.widthOfTextAtSize(cleanText, size);
  const textHeight = font.heightAtSize(size, { descender: false });
  page.drawText(cleanText, {
    x: rect.x + (rect.width - textWidth) / 2,
    y: rect.y + (rect.height - textHeight) / 2,
    size,
    font,
    color: rect.color,
  });
}

function fitFontSize(font: PDFFont, text: string, maxWidth: number, maxSize: number): number {
  let size = maxSize;
  while (size > 7 && font.widthOfTextAtSize(text, size) > maxWidth) {
    size -= 0.5;
  }
  return size;
}

async function fetchBinaryAsset(url: string, label: string): Promise<ArrayBuffer> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`${label}을 불러오지 못했습니다. 다시 시도해 주세요.`);
  }
  return response.arrayBuffer();
}
