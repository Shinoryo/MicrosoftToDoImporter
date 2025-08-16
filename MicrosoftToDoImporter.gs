const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = "offline_access Tasks.ReadWrite";

// -------------------------
// アクセストークン管理
// -------------------------
function getAuthProps() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auth");
  if (!sheet) {
    throw new Error("Authシートが存在しません");
  }

  return {
    clientId: sheet.getRange("A1").getValue(),
    clientSecret: sheet.getRange("A2").getValue(),
    accessToken: sheet.getRange("A4").getValue(),
    refreshToken: sheet.getRange("A5").getValue(),
    tokenExpiry: (() => {
      const val = sheet.getRange("A7").getValue();
      if (!val || isNaN(val)) {
        throw new Error("AuthシートのA7セル（tokenExpiry）が未設定です。認証処理を実行してください。");
      }
      return parseInt(val, 10);
    })()
  };
}

function getAccessToken() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auth");
  const auth = getAuthProps();

  if (!auth.accessToken || !auth.refreshToken) {
    throw new Error("Authシートにトークン情報がありません。初回認証が必要です。");
  }

  if (Date.now() > auth.tokenExpiry - 30000) {
    const payload = {
      client_id: auth.clientId,
      scope: SCOPES,
      refresh_token: auth.refreshToken,
      grant_type: "refresh_token",
      redirect_uri: REDIRECT_URI,
      client_secret: auth.clientSecret
    };
    const options = { method: "post", payload: payload };
    const response = UrlFetchApp.fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", options);
    const result = JSON.parse(response.getContentText());

    sheet.getRange("A4").setValue(result.access_token);
    sheet.getRange("A5").setValue(result.refresh_token || auth.refreshToken);
    sheet.getRange("A7").setValue(Date.now() + result.expires_in * 1000);

    return result.access_token;
  }

  return auth.accessToken;
}

// -------------------------
// Microsoft To Do 操作
// -------------------------
function getTodoListId(listName, accessToken) {
  const url = "https://graph.microsoft.com/v1.0/me/todo/lists";
  const options = { method: "get", headers: { Authorization: "Bearer " + accessToken } };
  const response = UrlFetchApp.fetch(url, options);
  const lists = JSON.parse(response.getContentText()).value;

  const list = lists.find(l => l.displayName === listName);
  if (!list) throw new Error("指定リストが見つかりません: " + listName);
  return list.id;
}

function addTasksFromSheet() {
  const ACCESS_TOKEN = getAccessToken();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  if (!sheet) throw new Error("Tasksシートが存在しません");

  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  let resultColIndex = headers.indexOf("result");
  if (resultColIndex === -1) {
    resultColIndex = headers.length;
    sheet.getRange(1, resultColIndex + 1).setValue("result");
  }

  rows.forEach((row, rowIndex) => {
    const task = {};
    headers.forEach((h, i) => task[h] = row[i]);

    if (!task.title || !task.list_name) {
      sheet.getRange(rowIndex + 2, resultColIndex + 1).setValue("title/list_name missing");
      return;
    }

    const payload = {
      title: task.title,
      status: task.status || "notStarted"
    };

    if (task.body) payload.body = { content: task.body, contentType: "text" };

    if (task.due) {
      // 日付のみの場合はISO形式に変換
      const dueDate = task.due.length === 10 ? task.due + "T23:59:00Z" : task.due;
      payload.dueDateTime = { dateTime: dueDate, timeZone: "UTC" };
    }

    // リマインダー設定（reminder列の有無）
    if (task.reminder) {
      const remDate = task.reminder.length === 10 ? task.reminder + "T09:00:00Z" : task.reminder;
      payload.reminderDateTime = { dateTime: remDate, timeZone: "UTC" };
    }

    // 繰り返しタスク
    if (task.recurrence_type && task.recurrence_start) {
      payload.recurrence = {
        pattern: {
          type: task.recurrence_type.toLowerCase(),
          interval: parseInt(task.recurrence_interval || 1, 10)
        },
        range: {
          type: task.recurrence_end ? "endDate" : "noEnd",
          startDate: task.recurrence_start,
          endDate: task.recurrence_end || undefined
        }
      };
    }

    try {
      const listId = getTodoListId(task.list_name, ACCESS_TOKEN);
      const url = `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks`;

      const options = {
        method: "post",
        headers: { Authorization: "Bearer " + ACCESS_TOKEN },
        contentType: "application/json",
        payload: JSON.stringify(payload)
      };

      UrlFetchApp.fetch(url, options);
      sheet.getRange(rowIndex + 2, resultColIndex + 1).setValue("Success");
    } catch (e) {
      sheet.getRange(rowIndex + 2, resultColIndex + 1).setValue("Error: " + e.message);
      Logger.log("タスク登録エラー: " + e);
    }
  });

  SpreadsheetApp.getUi().alert("タスク登録処理が完了しました！");
}

// -------------------------
// 認証用ボタン関数
// -------------------------
function generateAuthUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auth");
  const clientId = sheet.getRange("A1").getValue();

  const url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
              "?client_id=" + encodeURIComponent(clientId) +
              "&response_type=code" +
              "&redirect_uri=" + encodeURIComponent(REDIRECT_URI) +
              "&scope=" + encodeURIComponent(SCOPES);

  sheet.getRange("A6").setValue(url);
  SpreadsheetApp.getUi().alert("認証URLを生成しました。\nセルA6をクリックしてブラウザで開いてください。");
}

function exchangeCodeForTokenFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auth");
  const clientId = sheet.getRange("A1").getValue();
  const clientSecret = sheet.getRange("A2").getValue();
  const authCode = sheet.getRange("A3").getValue();

  if (!authCode) {
    SpreadsheetApp.getUi().alert("A3セルにAuthorization Codeを入力してください。");
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

  const options = { method: "post", payload: payload };
  const response = UrlFetchApp.fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", options);
  const result = JSON.parse(response.getContentText());

  sheet.getRange("A4").setValue(result.access_token);
  sheet.getRange("A5").setValue(result.refresh_token);
  sheet.getRange("A7").setValue(Date.now() + result.expires_in * 1000);

  SpreadsheetApp.getUi().alert("アクセストークンとリフレッシュトークンを取得しました。");
}

// -------------------------
// メニュー
// -------------------------
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Microsoft To Do")
    .addItem("認証URL生成", "generateAuthUrl")
    .addItem("トークン取得", "exchangeCodeForTokenFromSheet")
    .addSeparator()
    .addItem("TasksシートからTo Doに登録", "addTasksFromSheet")
    .addToUi();
}
