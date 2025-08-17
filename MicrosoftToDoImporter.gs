// GoogleスプレッドシートからMicrosoft To Doへタスクを連携するGASスクリプト

// シート名に関する定数
const SHEET_NAME_AUTH = "Auth";
const SHEET_NAME_TASKS = "Tasks";

// 認証情報などを格納するセルアドレスの定数
const CELL_CLIENT_ID = "A1";
const CELL_CLIENT_SECRET = "A2";
const CELL_AUTH_CODE = "A3";
const CELL_ACCESS_TOKEN = "A4";
const CELL_REFRESH_TOKEN = "A5";
const CELL_AUTH_URL = "A6";
const CELL_TOKEN_EXPIRY = "A7";

// Microsoft認証・APIアクセスに必要な各種定数
const MS_AUTH_ENDPOINT = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const MS_TOKEN_ENDPOINT = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = "offline_access Tasks.ReadWrite";
const MS_TODO_LISTS_ENDPOINT = "https://graph.microsoft.com/v1.0/me/todo/lists";
const MS_TODO_TASKS_ENDPOINT = "https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks";

// 列名定数
const COL_NAME_RESULT = "result";


// メッセージ定数（ユーザー向け・エラー・結果・バリデーション）
const MSG_SHEET_NOT_FOUND = "{sheetName}シートが存在しません";
const MSG_TOKEN_NOT_FOUND = "Authシートにトークン情報がありません。初回認証が必要です。";
const MSG_RESULT_COL_NOT_FOUND = "Tasksシートに'result'列がありません。'result'列を追加してください。";
const MSG_TASK_REGISTERED = "タスク登録処理が完了しました！";
const MSG_AUTH_URL_GENERATED = "認証URLを生成しました。\nセルA6をクリックしてブラウザで開いてください。";
const MSG_INPUT_AUTH_CODE = "A3セルにAuthorization Codeを入力してください。";
const MSG_TOKEN_ACQUIRED = "アクセストークンとリフレッシュトークンを取得しました。";
const MSG_LIST_NOT_FOUND = "指定リストが見つかりません: ";
const MSG_TOKEN_REQUEST_FAILED = "トークン取得リクエストに失敗しました: {msg}";
const MSG_TITLE_LISTNAME_MISSING = "title/list_name missing";
const MSG_INVALID_DUE_DATE = "due日付が不正です";
const MSG_INVALID_REMINDER_DATE = "reminder日付が不正です";
const TASK_RESULT_SUCCESS = "Success";
const TASK_RESULT_ERROR = "Error: {msg}";
const REGEX_REMOVE_MILLISECONDS = /\.\d{3}Z$/;

/**
 * 数値変換し、NaNならデフォルト値を返す
 * @param {any} val - 変換対象
 * @param {number} def - デフォルト値
 * @returns {number}
 */
function parseNumberOrDefault(val, def) {
    const num = Number(val);
    return isNaN(num) ? def : num;
}

/**
 * 指定したシート名のシートを取得し、存在しない場合はエラーを投げる。
 * @param {string} sheetName - 取得するシート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクト
 * @throws {Error} シートが存在しない場合
 */
function getSheetOrThrow(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(MSG_SHEET_NOT_FOUND.replace("{sheetName}", sheetName));
    }
    return sheet;
}

/**
 * Authシートから認証情報を取得する。
 * @returns {{clientId: string, clientSecret: string, accessToken: string, refreshToken: string, tokenExpiry: number}} 認証情報
 */
function getAuthProps() {
    const sheet = getSheetOrThrow(SHEET_NAME_AUTH);
    return {
        clientId: sheet.getRange(CELL_CLIENT_ID).getValue(),
        clientSecret: sheet.getRange(CELL_CLIENT_SECRET).getValue(),
        accessToken: sheet.getRange(CELL_ACCESS_TOKEN).getValue(),
        refreshToken: sheet.getRange(CELL_REFRESH_TOKEN).getValue(),
        tokenExpiry: parseNumberOrDefault(sheet.getRange(CELL_TOKEN_EXPIRY).getValue(), 0)
    };
}

/**
 * アクセストークンを取得（有効期限の30秒前を過ぎている場合はリフレッシュする）。
 * @returns {string} アクセストークン
 * @throws {Error} トークンが未取得の場合
 */
function getAccessToken() {
    const sheet = getSheetOrThrow(SHEET_NAME_AUTH);
    const auth = getAuthProps();
    // トークンが未取得の場合はエラー
    if (!auth.accessToken || !auth.refreshToken) {
        throw new Error(MSG_TOKEN_NOT_FOUND);
    }

    // 有効期限の30秒前を過ぎている場合はリフレッシュ
    if (Date.now() > auth.tokenExpiry - 30000) {
        const payload = {
            client_id: auth.clientId,
            scope: SCOPES,
            refresh_token: auth.refreshToken,
            grant_type: "refresh_token",
            redirect_uri: REDIRECT_URI,
            client_secret: auth.clientSecret
        };
        const postOptions = { method: "post", payload: payload };
        const postResponse = UrlFetchApp.fetch(MS_TOKEN_ENDPOINT, postOptions);
        const result = JSON.parse(postResponse.getContentText());

        // 新しいトークン情報をシートに保存
        sheet.getRange(CELL_ACCESS_TOKEN).setValue(result.access_token);
        sheet.getRange(CELL_REFRESH_TOKEN).setValue(result.refresh_token || auth.refreshToken);
        sheet.getRange(CELL_TOKEN_EXPIRY).setValue(Date.now() + result.expires_in * 1000);

        return result.access_token;
    }

    return auth.accessToken;
}


/**
 * 指定リスト名からMicrosoft To DoリストIDを取得する。
 * @param {string} listName - リスト名
 * @param {string} accessToken - アクセストークン
 * @returns {string} リストID
 * @throws {Error} リストが見つからない場合
 */
function getTodoListId(listName, accessToken) {
    // Microsoft To Doのリスト一覧を取得
    const getOptions = { method: "get", headers: { Authorization: "Bearer " + accessToken } };
    const getResponse = UrlFetchApp.fetch(MS_TODO_LISTS_ENDPOINT, getOptions);
    const lists = JSON.parse(getResponse.getContentText()).value;

    // 指定名のリストを検索
    const list = lists.find(l => l.displayName === listName);
    if (!list) {
        // 見つからなければエラー
        throw new Error(MSG_LIST_NOT_FOUND + listName);
    }
    return list.id;
}


/**
 * result列へ処理結果を書き込む（2行目以降がデータ）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} rowIndex - データ行インデックス（0始まり）
 * @param {number} resultColIndex - result列のインデックス（0始まり）
 * @param {string} value - 書き込む値
 */
function setResultToSheet(sheet, rowIndex, resultColIndex, value) {
    sheet.getRange(rowIndex + 2, resultColIndex + 1).setValue(value);
}


/**
 * タスクデータの必須項目（title, list_name）をチェックする。
 * @param {Object} task - タスクデータ
 * @returns {string|null} エラー時はエラーメッセージ、正常時はnull
 */
function validateTaskRow(task) {
    if (!task.title || !task.list_name) {
        return MSG_TITLE_LISTNAME_MISSING;
    }
    return null;
}

/**
 * タスクデータをMicrosoft To Do API用のリクエスト形式に変換する。
 * @param {Object} task - タスクデータ
 * @returns {Object} APIリクエスト用ペイロード
 */
function buildTaskPayload(task) {
    const payload = {
        title: task.title,
        status: task.status || "notStarted"
    };

    // 本文があれば追加
    if (task.body) {
        payload.body = { content: task.body, contentType: "text" };
    }

    // 期限があれば必ず23:59:00（ローカルタイムゾーン）を補完しUTC変換
    if (task.due) {
        const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
        let dueDate = new Date(task.due);
        if (isNaN(dueDate.getTime())) {
            throw new Error(MSG_INVALID_DUE_DATE);
        }
        // ローカルタイムゾーンで23:59:00を補完
        const dueLocalDateTimeStr = Utilities.formatDate(dueDate, tz, "yyyy-MM-dd") + " 23:59:00";
        // ローカルタイムゾーンの文字列をDateとしてUTCに変換
        const dueUtcDate = Utilities.parseDate(dueLocalDateTimeStr, tz);
        // ISO8601（ミリ秒除去）
        const dueIso = dueUtcDate.toISOString().replace(REGEX_REMOVE_MILLISECONDS, "Z");
        payload.dueDateTime = { dateTime: dueIso, timeZone: "UTC" };
    }

    // リマインダーがあれば追加
    if (task.reminder) {
        let d = new Date(task.reminder);
        if (isNaN(d.getTime())) {
            throw new Error(MSG_INVALID_REMINDER_DATE);
        }
        // toISOString()はミリ秒付き（.000Z）になるため、replaceでミリ秒を除去しISO8601形式（秒まで）に整形
        // 例: 2025-08-17T00:00:00.000Z → 2025-08-17T00:00:00Z
        const remIso = d.toISOString().replace(REGEX_REMOVE_MILLISECONDS, "Z");
        payload.reminderDateTime = { dateTime: remIso, timeZone: "UTC" };
    }

    // 繰り返し設定があれば追加
    if (task.recurrence_type && task.recurrence_start) {
        payload.recurrence = {
            pattern: {
                type: task.recurrence_type.toLowerCase(),
                interval: parseNumberOrDefault(task.recurrence_interval, 1)
            },
            range: {
                type: task.recurrence_end ? "endDate" : "noEnd",
                startDate: task.recurrence_start,
                endDate: task.recurrence_end || undefined
            }
        };
    }

    return payload;
}

/**
 * 1件のタスクをMicrosoft To Doへ登録するAPI呼び出し。
 * @param {Object} task - タスクデータ
 * @param {string} accessToken - アクセストークン
 */
function registerTaskToMicrosoftToDo(task, accessToken) {
    const listId = getTodoListId(task.list_name, accessToken);
    const url = MS_TODO_TASKS_ENDPOINT.replace("${listId}", listId);
    const payload = buildTaskPayload(task);
    const options = {
        method: "post",
        headers: { Authorization: "Bearer " + accessToken },
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    UrlFetchApp.fetch(url, options);
}

/**
 * シートからタスクデータを配列で取得（1行=1タスク）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {Object[]} タスクデータ配列
 */
function getTasksFromSheet(sheet) {
    const rows = sheet.getDataRange().getValues();
    const headers = rows.shift();
    return rows.map(row => {
        const task = {};
        headers.forEach((h, i) => task[h] = row[i]);
        return task;
    });
}

/**
 * シートの全タスクをMicrosoft To Doへ登録するメイン処理。
 */
function addTasksFromSheet() {
    const ACCESS_TOKEN = getAccessToken();
    const tasksSheet = getSheetOrThrow(SHEET_NAME_TASKS);
    const rows = tasksSheet.getDataRange().getValues();
    const headers = rows.shift();

    // result列のインデックスを取得
    let resultColIndex = headers.indexOf(COL_NAME_RESULT);
    if (resultColIndex === -1) {
        throw new Error(MSG_RESULT_COL_NOT_FOUND);
    }

    // 各行ごとにタスク登録処理
    rows.forEach((row, rowIndex) => {
        // タスクデータをオブジェクトに変換
        const task = {};
        headers.forEach((h, i) => task[h] = row[i]);
        
        // 必須項目チェック
        const validationError = validateTaskRow(task);
        if (validationError) {
            setResultToSheet(tasksSheet, rowIndex, resultColIndex, validationError);
            return;
        }

        // タスク登録API呼び出し
        try {
            registerTaskToMicrosoftToDo(task, ACCESS_TOKEN);
            setResultToSheet(tasksSheet, rowIndex, resultColIndex, TASK_RESULT_SUCCESS);
        } catch (e) {
            setResultToSheet(tasksSheet, rowIndex, resultColIndex, TASK_RESULT_ERROR.replace("{msg}", e.message));
        }
    });

    SpreadsheetApp.getUi().alert(MSG_TASK_REGISTERED);
}

/**
 * 認証URLを生成しAuthシートに出力する。
 */
function generateAuthUrl() {
    const authSheet = getSheetOrThrow(SHEET_NAME_AUTH);
    const clientId = authSheet.getRange(CELL_CLIENT_ID).getValue();
    const params = [
        ["client_id", clientId],
        ["response_type", "code"],
        ["redirect_uri", REDIRECT_URI],
        ["scope", SCOPES]
    ];
    const queryString = params.map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join("&");
    const url = `${MS_AUTH_ENDPOINT}?${queryString}`;
    authSheet.getRange(CELL_AUTH_URL).setValue(url);
    SpreadsheetApp.getUi().alert(MSG_AUTH_URL_GENERATED);
}

/**
 * 認証コードからアクセストークン・リフレッシュトークンを取得しAuthシートに保存する。
 */
function exchangeCodeForTokenFromSheet() {
    const authSheet = getSheetOrThrow(SHEET_NAME_AUTH);
    const clientId = authSheet.getRange(CELL_CLIENT_ID).getValue();
    const clientSecret = authSheet.getRange(CELL_CLIENT_SECRET).getValue();
    const authCode = authSheet.getRange(CELL_AUTH_CODE).getValue();
    if (!authCode) {
        SpreadsheetApp.getUi().alert(MSG_INPUT_AUTH_CODE);
        return;
    }

    const payload = {
        client_id: clientId,
        scope: SCOPES,
        code: authCode,
        redirect_uri: REDIRECT_URI,
        grant_type: "authorization_code",
        client_secret: clientSecret
    };
    const postOptions = { method: "post", payload: payload };
    try {
        const postResponse = UrlFetchApp.fetch(MS_TOKEN_ENDPOINT, postOptions);
        const result = JSON.parse(postResponse.getContentText());

        // トークン情報をシートに保存
        authSheet.getRange(CELL_ACCESS_TOKEN).setValue(result.access_token);
        authSheet.getRange(CELL_REFRESH_TOKEN).setValue(result.refresh_token);
        authSheet.getRange(CELL_TOKEN_EXPIRY).setValue(Date.now() + result.expires_in * 1000);

        SpreadsheetApp.getUi().alert(MSG_TOKEN_ACQUIRED);
    } catch (e) {
        SpreadsheetApp.getUi().alert(MSG_TOKEN_REQUEST_FAILED.replace("{msg}", e.message || e));
    }
}

/**
 * Googleスプレッドシートのメニューにカスタム項目を追加する（onOpenトリガー）。
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Microsoft To Do")
        .addItem("認証URL生成", "generateAuthUrl")
        .addItem("トークン取得", "exchangeCodeForTokenFromSheet")
        .addSeparator()
        .addItem("TasksシートからTo Doに登録", "addTasksFromSheet")
        .addToUi();
}
