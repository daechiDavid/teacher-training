use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use reqwest::blocking::{multipart, Client};
use reqwest::header::AUTHORIZATION;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::{
    env,
    fs,
    io::{Read, Write},
    net::TcpListener,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};
use tauri::{AppHandle, Manager};
use url::Url;
use uuid::Uuid;

const GOOGLE_AUTH_URL: &str = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL: &str = "https://oauth2.googleapis.com/token";
const GOOGLE_DRIVE_FILES_URL: &str = "https://www.googleapis.com/drive/v3/files";
const GOOGLE_DRIVE_UPLOAD_URL: &str = "https://www.googleapis.com/upload/drive/v3/files";
const GOOGLE_SHEETS_URL: &str = "https://sheets.googleapis.com/v4/spreadsheets";
const GOOGLE_SCOPES: &str =
    "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

#[derive(Debug, Serialize)]
struct AppStatus {
    app_data_dir: String,
    run_id: String,
}

#[derive(Debug, Deserialize, Serialize)]
struct JobLogEntry {
    run_id: String,
    training_id: String,
    person_key: String,
    sheet_row: Option<u32>,
    receipt_path: Option<String>,
    drive_file_id: Option<String>,
    drive_file_link: Option<String>,
    sheet_updated: bool,
    status: String,
    error_message: Option<String>,
    updated_at: String,
}

#[derive(Debug, Deserialize, Serialize, Default, Clone)]
struct GoogleConfig {
    client_id: String,
    client_secret: String,
    spreadsheet_id: String,
    sheet_name: String,
    drive_parent_folder_id: String,
}

#[derive(Debug, Deserialize)]
struct PreconfiguredOAuth {
    client_id: String,
    client_secret: String,
}

#[derive(Debug, Serialize)]
struct GoogleConfigStatus {
    configured: bool,
    authenticated: bool,
    client_id: String,
    spreadsheet_id: String,
    sheet_name: String,
    drive_parent_folder_id: String,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct GoogleToken {
    access_token: String,
    refresh_token: Option<String>,
    expires_at: u64,
    token_type: Option<String>,
    scope: Option<String>,
}

#[derive(Debug, Deserialize)]
struct TokenResponse {
    access_token: String,
    expires_in: Option<u64>,
    refresh_token: Option<String>,
    token_type: Option<String>,
    scope: Option<String>,
}

#[derive(Debug, Deserialize)]
struct TokenErrorResponse {
    error: String,
    error_description: Option<String>,
}

#[derive(Debug, Deserialize)]
struct SaveGoogleConfigRequest {
    client_id: String,
    client_secret: String,
    spreadsheet_id: String,
    sheet_name: String,
    drive_parent_folder_id: String,
}

#[derive(Debug, Deserialize)]
struct CreateDriveFolderRequest {
    training_name: String,
    issued_date: String,
}

#[derive(Debug, Serialize)]
struct DriveFolderResult {
    id: String,
    name: String,
}

#[derive(Debug, Deserialize)]
struct UploadPdfRequest {
    folder_id: String,
    filename: String,
    pdf_bytes: Vec<u8>,
}

#[derive(Debug, Deserialize)]
struct ReadSheetValuesRequest {
    spreadsheet_id: String,
    sheet_name: Option<String>,
    gid: Option<i64>,
}

#[derive(Debug, Deserialize)]
struct ResolveSheetTitleRequest {
    spreadsheet_id: String,
    gid: Option<i64>,
}

#[derive(Debug, Serialize)]
struct UploadedFileResult {
    id: String,
    name: String,
    web_view_link: String,
}

#[derive(Debug, Deserialize)]
struct SheetValueUpdate {
    range: String,
    values: Vec<Vec<String>>,
}

#[derive(Debug, Deserialize)]
struct BatchUpdateSheetRequest {
    updates: Vec<SheetValueUpdate>,
}

#[derive(Debug, Deserialize)]
struct BatchUpdateAnySheetRequest {
    spreadsheet_id: String,
    sheet_name: String,
    updates: Vec<SheetValueUpdate>,
}

#[tauri::command]
fn app_status(app: AppHandle) -> Result<AppStatus, String> {
    let dir = app_data_dir(&app)?;
    fs::create_dir_all(&dir).map_err(|error| error.to_string())?;
    Ok(AppStatus {
        app_data_dir: dir.to_string_lossy().to_string(),
        run_id: Uuid::new_v4().to_string(),
    })
}

#[tauri::command]
fn append_job_log(app: AppHandle, entry: JobLogEntry) -> Result<(), String> {
    let dir = app_data_dir(&app)?;
    fs::create_dir_all(&dir).map_err(|error| error.to_string())?;
    let path = dir.join("job-log.jsonl");
    let line = serde_json::to_string(&entry).map_err(|error| error.to_string())?;
    append_line(&path, &line)
}

#[tauri::command]
fn save_google_config(app: AppHandle, request: SaveGoogleConfigRequest) -> Result<(), String> {
    let existing = read_google_config(&app).unwrap_or_default();
    let preconfigured = read_preconfigured_oauth(&app).ok();
    let resolved_client_id = first_non_empty(&[
        request.client_id.trim(),
        existing.client_id.as_str(),
        preconfigured.as_ref().map(|value| value.client_id.as_str()).unwrap_or(""),
    ]);
    let resolved_client_secret = first_non_empty(&[
        request.client_secret.trim(),
        existing.client_secret.as_str(),
        preconfigured.as_ref().map(|value| value.client_secret.as_str()).unwrap_or(""),
    ]);
    let config = GoogleConfig {
        client_id: first_non_empty(&[request.client_id.trim(), existing.client_id.as_str()]),
        client_secret: first_non_empty(&[request.client_secret.trim(), existing.client_secret.as_str()]),
        spreadsheet_id: request.spreadsheet_id.trim().to_string(),
        sheet_name: if request.sheet_name.trim().is_empty() {
            "명단".to_string()
        } else {
            request.sheet_name.trim().to_string()
        },
        drive_parent_folder_id: request.drive_parent_folder_id.trim().to_string(),
    };

    if resolved_client_id.is_empty()
        || resolved_client_secret.is_empty()
        || config.spreadsheet_id.is_empty()
        || config.drive_parent_folder_id.is_empty()
    {
        return Err("관리자 OAuth 사전 설정, Spreadsheet ID, Drive 상위 폴더 ID는 필수입니다.".to_string());
    }

    write_json_secure(&google_config_path(&app)?, &config)
}

#[tauri::command]
fn google_logout(app: AppHandle) -> Result<(), String> {
    let path = google_token_path(&app)?;
    if path.exists() {
        fs::remove_file(&path).map_err(|error| error.to_string())?;
    }
    Ok(())
}

#[tauri::command]
fn google_config_status(app: AppHandle) -> Result<GoogleConfigStatus, String> {
    let config = read_google_config(&app).unwrap_or_default();
    let oauth = resolve_google_oauth_config(&app).unwrap_or_default();
    let authenticated = google_token_path(&app)?.exists();
    Ok(GoogleConfigStatus {
        configured: !oauth.client_id.is_empty()
            && !oauth.client_secret.is_empty()
            && !config.spreadsheet_id.is_empty()
            && !config.drive_parent_folder_id.is_empty(),
        authenticated,
        client_id: mask(&oauth.client_id),
        spreadsheet_id: config.spreadsheet_id,
        sheet_name: if config.sheet_name.is_empty() {
            "명단".to_string()
        } else {
            config.sheet_name
        },
        drive_parent_folder_id: config.drive_parent_folder_id,
    })
}

#[tauri::command]
fn start_google_oauth(app: AppHandle) -> Result<(), String> {
    let config = resolve_google_oauth_config(&app)?;
    let listener = TcpListener::bind("127.0.0.1:0").map_err(|error| error.to_string())?;
    let port = listener.local_addr().map_err(|error| error.to_string())?.port();
    let redirect_uri = format!("http://127.0.0.1:{port}/oauth2/callback");
    let state = Uuid::new_v4().to_string();
    let code_verifier = format!(
        "{}{}{}{}",
        Uuid::new_v4(),
        Uuid::new_v4(),
        Uuid::new_v4(),
        Uuid::new_v4()
    );
    let code_challenge = pkce_challenge(&code_verifier);

    let mut auth_url = Url::parse(GOOGLE_AUTH_URL).map_err(|error| error.to_string())?;
    auth_url.query_pairs_mut()
        .append_pair("client_id", &config.client_id)
        .append_pair("redirect_uri", &redirect_uri)
        .append_pair("response_type", "code")
        .append_pair("scope", GOOGLE_SCOPES)
        .append_pair("access_type", "offline")
        .append_pair("prompt", "consent")
        .append_pair("state", &state)
        .append_pair("code_challenge", &code_challenge)
        .append_pair("code_challenge_method", "S256");

    open_url(auth_url.as_str())?;

    let code = wait_for_oauth_code(listener, &state)?;
    let token = exchange_code_for_token(&config, &redirect_uri, &code, &code_verifier)?;
    write_json_secure(&google_token_path(&app)?, &token)
}

#[tauri::command]
fn google_read_roster_values(app: AppHandle) -> Result<Vec<Vec<String>>, String> {
    let config = resolve_google_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    read_sheet_values_with_token(&token, &config.spreadsheet_id, &config.sheet_name)
}

#[tauri::command]
fn google_read_sheet_values(app: AppHandle, request: ReadSheetValuesRequest) -> Result<Vec<Vec<String>>, String> {
    let config = resolve_google_oauth_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    let spreadsheet_id = request.spreadsheet_id.trim();
    if spreadsheet_id.is_empty() {
        return Err("Google Sheet ID가 비어 있습니다.".to_string());
    }

    let sheet_name = if let Some(sheet_name) = request.sheet_name {
        sheet_name.trim().to_string()
    } else {
        String::new()
    };
    let resolved_sheet_name = if sheet_name.is_empty() {
        resolve_sheet_title(&token, spreadsheet_id, request.gid)?
    } else {
        sheet_name
    };
    read_sheet_values_with_token(&token, spreadsheet_id, &resolved_sheet_name)
}

#[tauri::command]
fn google_resolve_sheet_title(app: AppHandle, request: ResolveSheetTitleRequest) -> Result<String, String> {
    let config = resolve_google_oauth_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    let spreadsheet_id = request.spreadsheet_id.trim();
    if spreadsheet_id.is_empty() {
        return Err("Google Sheet ID가 비어 있습니다.".to_string());
    }
    resolve_sheet_title(&token, spreadsheet_id, request.gid)
}

fn read_sheet_values_with_token(token: &str, spreadsheet_id: &str, sheet_name: &str) -> Result<Vec<Vec<String>>, String> {
    let range = encode_a1_range(&format!("{}!A:Z", quote_sheet_name(sheet_name)));
    let url = format!("{GOOGLE_SHEETS_URL}/{spreadsheet_id}/values/{range}");
    let response = Client::new()
        .get(url)
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .send()
        .map_err(|error| error.to_string())?;

    if !response.status().is_success() {
        return Err(format_google_error(response));
    }

    let body: serde_json::Value = response.json().map_err(|error| error.to_string())?;
    let rows = body
        .get("values")
        .and_then(|value| value.as_array())
        .map(|rows| {
            rows.iter()
                .map(|row| {
                    if let Some(cells) = row.as_array() {
                        cells
                            .iter()
                            .map(google_cell_to_string)
                            .collect::<Vec<_>>()
                    } else {
                        Vec::new()
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Ok(rows)
}

#[tauri::command]
fn google_batch_update_sheet(app: AppHandle, request: BatchUpdateSheetRequest) -> Result<(), String> {
    let config = resolve_google_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    batch_update_sheet_values(&token, &config.spreadsheet_id, &config.sheet_name, request.updates)
}

#[tauri::command]
fn google_batch_update_any_sheet(app: AppHandle, request: BatchUpdateAnySheetRequest) -> Result<(), String> {
    let config = resolve_google_oauth_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    let spreadsheet_id = request.spreadsheet_id.trim();
    let sheet_name = request.sheet_name.trim();
    if spreadsheet_id.is_empty() || sheet_name.is_empty() {
        return Err("업데이트할 Google Sheet ID와 탭 이름은 필수입니다.".to_string());
    }
    batch_update_sheet_values(&token, spreadsheet_id, sheet_name, request.updates)
}

fn batch_update_sheet_values(
    token: &str,
    spreadsheet_id: &str,
    sheet_name: &str,
    updates: Vec<SheetValueUpdate>,
) -> Result<(), String> {
    let data = updates
        .into_iter()
        .map(|update| {
            serde_json::json!({
                "range": format!("{}!{}", quote_sheet_name(sheet_name), update.range),
                "values": update.values,
            })
        })
        .collect::<Vec<_>>();

    let url = format!("{GOOGLE_SHEETS_URL}/{spreadsheet_id}/values:batchUpdate");
    let response = Client::new()
        .post(url)
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .json(&serde_json::json!({
            "valueInputOption": "USER_ENTERED",
            "data": data,
        }))
        .send()
        .map_err(|error| error.to_string())?;

    if response.status().is_success() {
        Ok(())
    } else {
        Err(format_google_error(response))
    }
}

#[tauri::command]
fn drive_create_training_folder(
    app: AppHandle,
    request: CreateDriveFolderRequest,
) -> Result<DriveFolderResult, String> {
    let config = resolve_google_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    let folder_name = safe_drive_name(&format!("{}_{}", request.training_name, request.issued_date));
    let response = Client::new()
        .post(GOOGLE_DRIVE_FILES_URL)
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .query(&[("fields", "id,name")])
        .json(&serde_json::json!({
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [config.drive_parent_folder_id],
        }))
        .send()
        .map_err(|error| error.to_string())?;

    if !response.status().is_success() {
        return Err(format_google_error(response));
    }

    let body: serde_json::Value = response.json().map_err(|error| error.to_string())?;
    let id = body.get("id").and_then(|value| value.as_str()).unwrap_or("").to_string();
    if id.trim().is_empty() {
        return Err("Google Drive 폴더 생성 응답에 폴더 ID가 없습니다.".to_string());
    }
    Ok(DriveFolderResult {
        id,
        name: body.get("name").and_then(|value| value.as_str()).unwrap_or("").to_string(),
    })
}

#[tauri::command]
fn drive_upload_pdf(app: AppHandle, request: UploadPdfRequest) -> Result<UploadedFileResult, String> {
    let config = resolve_google_config(&app)?;
    let token = valid_access_token(&app, &config)?;
    let filename = safe_drive_name(&request.filename);
    let metadata = serde_json::json!({
        "name": filename,
        "parents": [request.folder_id],
        "mimeType": "application/pdf",
    })
    .to_string();
    let form = multipart::Form::new()
        .part(
            "metadata",
            multipart::Part::text(metadata).mime_str("application/json; charset=UTF-8").map_err(|error| error.to_string())?,
        )
        .part(
            "file",
            multipart::Part::bytes(request.pdf_bytes)
                .file_name(filename)
                .mime_str("application/pdf")
                .map_err(|error| error.to_string())?,
        );

    let response = Client::new()
        .post(GOOGLE_DRIVE_UPLOAD_URL)
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .query(&[("uploadType", "multipart"), ("fields", "id,name,webViewLink")])
        .multipart(form)
        .send()
        .map_err(|error| error.to_string())?;

    if !response.status().is_success() {
        return Err(format_google_error(response));
    }

    let body: serde_json::Value = response.json().map_err(|error| error.to_string())?;
    let file_id = body.get("id").and_then(|value| value.as_str()).unwrap_or("").to_string();
    if file_id.trim().is_empty() {
        return Err("Google Drive 업로드 응답에 파일 ID가 없습니다.".to_string());
    }
    share_drive_file(&token, &file_id)?;
    let web_view_link = body
        .get("webViewLink")
        .and_then(|value| value.as_str())
        .unwrap_or("")
        .to_string();
    if web_view_link.trim().is_empty() {
        return Err("Google Drive 업로드 응답에 공유 링크가 없습니다.".to_string());
    }
    Ok(UploadedFileResult {
        id: file_id,
        name: body.get("name").and_then(|value| value.as_str()).unwrap_or("").to_string(),
        web_view_link,
    })
}

pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            app_status,
            append_job_log,
            save_google_config,
            google_logout,
            google_config_status,
            start_google_oauth,
            google_read_roster_values,
            google_read_sheet_values,
            google_resolve_sheet_title,
            google_batch_update_sheet,
            google_batch_update_any_sheet,
            drive_create_training_folder,
            drive_upload_pdf
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

fn app_data_dir(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_data_dir()
        .map_err(|error| format!("앱 데이터 폴더를 확인하지 못했습니다: {error}"))
}

fn append_line(path: &PathBuf, line: &str) -> Result<(), String> {
    use std::io::Write;

    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .map_err(|error| error.to_string())?;
    writeln!(file, "{line}").map_err(|error| error.to_string())
}

fn google_config_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app_data_dir(app)?.join("google-config.json"))
}

fn google_token_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app_data_dir(app)?.join("google-token.json"))
}

fn read_google_config(app: &AppHandle) -> Result<GoogleConfig, String> {
    read_json(&google_config_path(app)?)
}

fn read_google_token(app: &AppHandle) -> Result<GoogleToken, String> {
    read_json(&google_token_path(app)?)
}

fn resolve_google_oauth_config(app: &AppHandle) -> Result<GoogleConfig, String> {
    let mut config = read_google_config(app).unwrap_or_default();
    if config.client_id.trim().is_empty() || config.client_secret.trim().is_empty() {
        let preconfigured = read_preconfigured_oauth(app)?;
        if config.client_id.trim().is_empty() {
            config.client_id = preconfigured.client_id;
        }
        if config.client_secret.trim().is_empty() {
            config.client_secret = preconfigured.client_secret;
        }
    }

    if config.client_id.trim().is_empty() || config.client_secret.trim().is_empty() {
        return Err("관리자 OAuth 사전 설정이 없습니다.".to_string());
    }

    Ok(config)
}

fn resolve_google_config(app: &AppHandle) -> Result<GoogleConfig, String> {
    let mut config = resolve_google_oauth_config(app)?;
    let saved = read_google_config(app)?;
    config.spreadsheet_id = saved.spreadsheet_id;
    config.sheet_name = saved.sheet_name;
    config.drive_parent_folder_id = saved.drive_parent_folder_id;
    if config.sheet_name.trim().is_empty() {
        config.sheet_name = "명단".to_string();
    }

    if config.client_id.trim().is_empty()
        || config.client_secret.trim().is_empty()
        || config.spreadsheet_id.trim().is_empty()
        || config.drive_parent_folder_id.trim().is_empty()
    {
        return Err("Google 설정이 완료되지 않았습니다. 관리자 OAuth 사전 설정과 Sheet/Drive 설정을 확인하세요.".to_string());
    }

    Ok(config)
}

fn read_preconfigured_oauth(app: &AppHandle) -> Result<PreconfiguredOAuth, String> {
    let env_client_id = env::var("GOOGLE_OAUTH_CLIENT_ID").unwrap_or_default();
    let env_client_secret = env::var("GOOGLE_OAUTH_CLIENT_SECRET").unwrap_or_default();
    if !env_client_id.trim().is_empty() && !env_client_secret.trim().is_empty() {
        return Ok(PreconfiguredOAuth {
            client_id: env_client_id.trim().to_string(),
            client_secret: env_client_secret.trim().to_string(),
        });
    }

    let mut candidates = Vec::new();
    if let Ok(current_dir) = env::current_dir() {
        candidates.push(current_dir.join("google-oauth.local.json"));
        candidates.push(current_dir.join("src-tauri").join("google-oauth.local.json"));
        candidates.push(current_dir.join("google-oauth.json"));
        candidates.push(current_dir.join("src-tauri").join("google-oauth.example.json"));
    }
    if let Ok(app_dir) = app_data_dir(app) {
        candidates.push(app_dir.join("google-oauth.json"));
        candidates.push(app_dir.join("google-oauth.example.json"));
    }
    if let Ok(exe_path) = env::current_exe() {
        if let Some(exe_dir) = exe_path.parent() {
            candidates.push(exe_dir.join("google-oauth.json"));
            candidates.push(exe_dir.join("google-oauth.example.json"));
        }
    }
    if let Ok(resource_dir) = app.path().resource_dir() {
        candidates.push(resource_dir.join("google-oauth.json"));
        candidates.push(resource_dir.join("google-oauth.example.json"));
    }

    for path in candidates {
        if !path.exists() {
            continue;
        }
        let oauth = read_oauth_file(&path)?;
        if !oauth.client_id.trim().is_empty() && !oauth.client_secret.trim().is_empty() {
            return Ok(PreconfiguredOAuth {
                client_id: oauth.client_id.trim().to_string(),
                client_secret: oauth.client_secret.trim().to_string(),
            });
        }
    }

    Err("관리자가 Google OAuth 설정 파일을 아직 준비하지 않았습니다.".to_string())
}

fn read_oauth_file(path: &Path) -> Result<PreconfiguredOAuth, String> {
    let content = fs::read_to_string(path)
        .map_err(|error| format!("OAuth 설정 파일을 읽지 못했습니다: {} ({error})", path.display()))?;
    serde_json::from_str(&content)
        .map_err(|error| format!("OAuth 설정 파일 형식이 올바르지 않습니다: {} ({error})", path.display()))
}

fn first_non_empty(values: &[&str]) -> String {
    values
        .iter()
        .map(|value| value.trim())
        .find(|value| !value.is_empty())
        .unwrap_or("")
        .to_string()
}

fn read_json<T: for<'de> Deserialize<'de>>(path: &PathBuf) -> Result<T, String> {
    let filename = path.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    let content = fs::read_to_string(path).map_err(|_| format!("파일을 찾을 수 없습니다: {filename} (경로: {})", path.display()))?;
    serde_json::from_str(&content).map_err(|error| format!("{filename} 파싱 실패: {error}"))
}

fn write_json<T: Serialize>(path: &PathBuf, value: &T) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let content = serde_json::to_string_pretty(value).map_err(|error| error.to_string())?;
    fs::write(path, content).map_err(|error| error.to_string())
}

fn write_json_secure<T: Serialize>(path: &PathBuf, value: &T) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
        set_owner_only_dir(parent)?;
    }
    write_json(path, value)?;
    set_owner_only_file(path)
}

#[cfg(unix)]
fn set_owner_only_dir(path: &std::path::Path) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    let mut permissions = fs::metadata(path).map_err(|error| error.to_string())?.permissions();
    permissions.set_mode(0o700);
    fs::set_permissions(path, permissions).map_err(|error| error.to_string())
}

#[cfg(not(unix))]
fn set_owner_only_dir(_path: &std::path::Path) -> Result<(), String> {
    Ok(())
}

#[cfg(unix)]
fn set_owner_only_file(path: &std::path::Path) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    let mut permissions = fs::metadata(path).map_err(|error| error.to_string())?.permissions();
    permissions.set_mode(0o600);
    fs::set_permissions(path, permissions).map_err(|error| error.to_string())
}

#[cfg(not(unix))]
fn set_owner_only_file(_path: &std::path::Path) -> Result<(), String> {
    Ok(())
}

fn now_epoch_seconds() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
}

fn pkce_challenge(verifier: &str) -> String {
    let digest = Sha256::digest(verifier.as_bytes());
    URL_SAFE_NO_PAD.encode(digest)
}

fn open_url(url: &str) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    let mut command = {
        let mut command = Command::new("open");
        command.arg(url);
        command
    };

    #[cfg(target_os = "windows")]
    let mut command = {
        let mut command = Command::new("rundll32");
        command.args(["url.dll,FileProtocolHandler", url]);
        command
    };

    #[cfg(target_os = "linux")]
    let mut command = {
        let mut command = Command::new("xdg-open");
        command.arg(url);
        command
    };

    command.status().map_err(|error| error.to_string())?;
    Ok(())
}

fn wait_for_oauth_code(listener: TcpListener, expected_state: &str) -> Result<String, String> {
    let (mut stream, _) = listener.accept().map_err(|error| error.to_string())?;
    let mut buffer = [0_u8; 4096];
    let size = stream.read(&mut buffer).map_err(|error| error.to_string())?;
    let request = String::from_utf8_lossy(&buffer[..size]);
    let first_line = request.lines().next().unwrap_or("");
    let path = first_line.split_whitespace().nth(1).unwrap_or("");
    let url = Url::parse(&format!("http://127.0.0.1{path}")).map_err(|error| error.to_string())?;
    let code = url
        .query_pairs()
        .find(|(key, _)| key == "code")
        .map(|(_, value)| value.to_string());
    let state = url
        .query_pairs()
        .find(|(key, _)| key == "state")
        .map(|(_, value)| value.to_string());
    let error = url
        .query_pairs()
        .find(|(key, _)| key == "error")
        .map(|(_, value)| value.to_string());

    let body = if error.is_some() {
        "Google 인증이 취소되었거나 실패했습니다. 앱으로 돌아가 다시 시도하세요."
    } else {
        "Google 인증이 완료되었습니다. 이 창을 닫고 앱으로 돌아가세요."
    };
    let response = format!(
        "HTTP/1.1 200 OK\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {}\r\n\r\n{}",
        body.as_bytes().len(),
        body
    );
    let _ = stream.write_all(response.as_bytes());

    if let Some(error) = error {
        return Err(format!("Google 인증 실패: {error}"));
    }
    if state.as_deref() != Some(expected_state) {
        return Err("OAuth state 값이 일치하지 않습니다.".to_string());
    }
    code.ok_or_else(|| "OAuth 인증 코드가 응답에 없습니다.".to_string())
}

fn exchange_code_for_token(
    config: &GoogleConfig,
    redirect_uri: &str,
    code: &str,
    code_verifier: &str,
) -> Result<GoogleToken, String> {
    let response = Client::new()
        .post(GOOGLE_TOKEN_URL)
        .form(&[
            ("client_id", config.client_id.as_str()),
            ("client_secret", config.client_secret.as_str()),
            ("code", code),
            ("code_verifier", code_verifier),
            ("grant_type", "authorization_code"),
            ("redirect_uri", redirect_uri),
        ])
        .send()
        .map_err(|error| error.to_string())?;
    token_from_response(response)
}

fn refresh_access_token(config: &GoogleConfig, refresh_token: &str) -> Result<GoogleToken, String> {
    let response = Client::new()
        .post(GOOGLE_TOKEN_URL)
        .form(&[
            ("client_id", config.client_id.as_str()),
            ("client_secret", config.client_secret.as_str()),
            ("refresh_token", refresh_token),
            ("grant_type", "refresh_token"),
        ])
        .send()
        .map_err(|error| error.to_string())?;
    token_from_response(response)
}

fn token_from_response(response: reqwest::blocking::Response) -> Result<GoogleToken, String> {
    let status = response.status();
    let text = response.text().map_err(|error| error.to_string())?;
    if !status.is_success() {
        if let Ok(error) = serde_json::from_str::<TokenErrorResponse>(&text) {
            return Err(format!(
                "Google 토큰 요청 실패: {} {}",
                error.error,
                error.error_description.unwrap_or_default()
            ));
        }
        return Err(format!("Google 토큰 요청 실패: HTTP {status} {text}"));
    }
    let body: TokenResponse = serde_json::from_str(&text).map_err(|error| error.to_string())?;
    Ok(GoogleToken {
        access_token: body.access_token,
        refresh_token: body.refresh_token,
        expires_at: now_epoch_seconds() + body.expires_in.unwrap_or(3600).saturating_sub(60),
        token_type: body.token_type,
        scope: body.scope,
    })
}

fn valid_access_token(app: &AppHandle, config: &GoogleConfig) -> Result<String, String> {
    let mut token = read_google_token(app)?;
    if token.expires_at > now_epoch_seconds() + 30 {
        return Ok(token.access_token);
    }
    let refresh_token = token
        .refresh_token
        .clone()
        .ok_or_else(|| "Refresh token이 없습니다. Google 로그인을 다시 진행하세요.".to_string())?;
    let mut refreshed = refresh_access_token(config, &refresh_token)?;
    refreshed.refresh_token = Some(refresh_token);
    token = refreshed;
    write_json_secure(&google_token_path(app)?, &token)?;
    Ok(token.access_token)
}

fn encode_a1_range(range: &str) -> String {
    range
        .as_bytes()
        .iter()
        .map(|byte| match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'.' | b'_' | b'~' => {
                (*byte as char).to_string()
            }
            _ => format!("%{byte:02X}"),
        })
        .collect::<String>()
}

fn quote_sheet_name(sheet_name: &str) -> String {
    format!("'{}'", sheet_name.replace('\'', "''"))
}

fn resolve_sheet_title(token: &str, spreadsheet_id: &str, gid: Option<i64>) -> Result<String, String> {
    let response = Client::new()
        .get(format!("{GOOGLE_SHEETS_URL}/{spreadsheet_id}"))
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .query(&[("fields", "sheets(properties(sheetId,title,index))")])
        .send()
        .map_err(|error| error.to_string())?;

    if !response.status().is_success() {
        return Err(format_google_error(response));
    }

    let body: serde_json::Value = response.json().map_err(|error| error.to_string())?;
    let sheets = body
        .get("sheets")
        .and_then(|value| value.as_array())
        .ok_or_else(|| "Google Sheet 탭 정보를 찾지 못했습니다.".to_string())?;

    let mut first_title: Option<String> = None;
    for sheet in sheets {
        let properties = sheet.get("properties").unwrap_or(&serde_json::Value::Null);
        let title = properties
            .get("title")
            .and_then(|value| value.as_str())
            .unwrap_or("")
            .to_string();
        if title.is_empty() {
            continue;
        }
        if first_title.is_none() {
            first_title = Some(title.clone());
        }
        if let Some(expected_gid) = gid {
            if properties.get("sheetId").and_then(|value| value.as_i64()) == Some(expected_gid) {
                return Ok(title);
            }
        }
    }

    if gid.is_some() {
        return Err("URL의 gid와 일치하는 Google Sheet 탭을 찾지 못했습니다.".to_string());
    }

    first_title.ok_or_else(|| "Google Sheet에 읽을 수 있는 탭이 없습니다.".to_string())
}

fn google_cell_to_string(value: &serde_json::Value) -> String {
    if let Some(text) = value.as_str() {
        return text.to_string();
    }
    if let Some(number) = value.as_f64() {
        if number.fract() == 0.0 {
            return format!("{}", number as i64);
        }
        return number.to_string();
    }
    if let Some(boolean) = value.as_bool() {
        return boolean.to_string();
    }
    String::new()
}

fn share_drive_file(token: &str, file_id: &str) -> Result<(), String> {
    let url = format!("{GOOGLE_DRIVE_FILES_URL}/{file_id}/permissions");
    let response = Client::new()
        .post(url)
        .header(AUTHORIZATION, format!("Bearer {token}"))
        .query(&[("sendNotificationEmail", "false")])
        .json(&serde_json::json!({
            "type": "anyone",
            "role": "reader",
        }))
        .send()
        .map_err(|error| error.to_string())?;

    if response.status().is_success() {
        Ok(())
    } else {
        Err(format_google_error(response))
    }
}

fn format_google_error(response: reqwest::blocking::Response) -> String {
    let status = response.status();
    let text = response.text().unwrap_or_default();
    format!("Google API 요청 실패: HTTP {status} {text}")
}

fn mask(value: &str) -> String {
    if value.len() <= 10 {
        return value.to_string();
    }
    format!("{}...{}", &value[..6], &value[value.len() - 4..])
}

fn safe_drive_name(value: &str) -> String {
    let sanitized = value
        .chars()
        .map(|character| match character {
            '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|' | '#' | '%' | '{' | '}'
            | '[' | ']' | '^' | '~' | '`' => ' ',
            _ => character,
        })
        .collect::<String>();
    sanitized.split_whitespace().collect::<Vec<_>>().join(" ")
}
