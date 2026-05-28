import { invoke } from "@tauri-apps/api/core";

export type GoogleConfigInput = {
  spreadsheetId: string;
  sheetName: string;
  driveParentFolderId: string;
};

export type GoogleConfigStatus = {
  configured: boolean;
  authenticated: boolean;
  client_id: string;
  spreadsheet_id: string;
  sheet_name: string;
  drive_parent_folder_id: string;
};

export type AppStatus = {
  app_data_dir: string;
  run_id: string;
};

export type JobLogEntry = {
  run_id: string;
  training_id: string;
  person_key: string;
  sheet_row?: number;
  receipt_path?: string;
  drive_file_id?: string;
  drive_file_link?: string;
  sheet_updated: boolean;
  status: string;
  error_message?: string;
  updated_at: string;
};

export async function getAppStatus(): Promise<AppStatus> {
  return invoke("app_status");
}

export async function appendJobLog(entry: JobLogEntry): Promise<void> {
  await invoke("append_job_log", { entry });
}

export async function saveGoogleConfig(config: GoogleConfigInput): Promise<void> {
  await invoke("save_google_config", {
    request: {
      client_id: "",
      client_secret: "",
      spreadsheet_id: config.spreadsheetId,
      sheet_name: config.sheetName,
      drive_parent_folder_id: config.driveParentFolderId,
    },
  });
}

export async function getGoogleConfigStatus(): Promise<GoogleConfigStatus> {
  return invoke("google_config_status");
}

export async function startGoogleOAuth(): Promise<void> {
  await invoke("start_google_oauth");
}

export async function readGoogleRosterValues(): Promise<unknown[][]> {
  return invoke("google_read_roster_values");
}

export async function readGoogleSheetValues(
  spreadsheetId: string,
  gid?: number,
  sheetName?: string,
): Promise<unknown[][]> {
  return invoke("google_read_sheet_values", {
    request: {
      spreadsheet_id: spreadsheetId,
      gid,
      sheet_name: sheetName,
    },
  });
}

export async function resolveGoogleSheetTitle(spreadsheetId: string, gid?: number): Promise<string> {
  return invoke("google_resolve_sheet_title", {
    request: {
      spreadsheet_id: spreadsheetId,
      gid,
    },
  });
}

export async function createDriveTrainingFolder(trainingName: string, issuedDate: string): Promise<{
  id: string;
  name: string;
}> {
  return invoke("drive_create_training_folder", {
    request: {
      training_name: trainingName,
      issued_date: issuedDate,
    },
  });
}

export async function uploadPdfToDrive(
  folderId: string,
  filename: string,
  pdfBytes: number[],
): Promise<{
  id: string;
  name: string;
  web_view_link: string;
}> {
  return invoke("drive_upload_pdf", {
    request: {
      folder_id: folderId,
      filename,
      pdf_bytes: pdfBytes,
    },
  });
}

export async function batchUpdateSheet(updates: Array<{ range: string; values: string[][] }>): Promise<void> {
  await invoke("google_batch_update_sheet", {
    request: {
      updates,
    },
  });
}

export async function batchUpdateGoogleSheet(
  spreadsheetId: string,
  sheetName: string,
  updates: Array<{ range: string; values: string[][] }>,
): Promise<void> {
  await invoke("google_batch_update_any_sheet", {
    request: {
      spreadsheet_id: spreadsheetId,
      sheet_name: sheetName,
      updates,
    },
  });
}
