/**
 * 就活メール自動ラベル仕分けシステム
 *
 * このスクリプトは、指定されたGmailアカウントから未読の就活メールを検索し、
 * Gemini APIを使用して締め切り日を抽出して、曜日付きの期限ラベルを自動付与します。
 *
 * 動作には以下の設定が必要です：
 * 1. プロジェクトのプロパティに「GEMINI_API_KEY」を設定
 */

// グローバル定数とプロパティの取得
const PROPERTIES = PropertiesService.getScriptProperties();
const GEMINI_API_KEY = PROPERTIES.getProperty('GEMINI_API_KEY');
const CALENDAR_NAME = PROPERTIES.getProperty('CALENDAR_NAME') || '就活';

// 検索クエリ
const SEARCH_QUERY = 'is:unread -label:処理済み (subject:("締め切り" OR "締切" OR "〆切" OR "期限" OR "提出" OR "予約" OR "受験" OR "受検" OR "テスト") OR "エントリーシート" OR "ES") ("インターン" OR "選考" ) -("メルマガ" OR "ニュースレター" OR "コラム" OR "マガジン")';

/**
 * 1時間に1回実行するメイン処理
 */
function main() {
  if (!GEMINI_API_KEY) {
    console.error('スクリプトプロパティ GEMINI_API_KEY が設定されていません。');
    return;
  }

  // 1. 対象スレッドの検索
  const threads = GmailApp.search(SEARCH_QUERY);
  if (threads.length === 0) {
    console.log('処理対象の未読メールはありません。');
    return;
  }
  
  console.log(`${threads.length} 件のスレッドを処理します。`);
  
  // 2. 「処理済み」ラベルの取得または作成
  const processedLabel = getOrCreateLabel('処理済み');

  // 3. 各スレッドのループ処理
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    console.log(`[${i + 1}/${threads.length}] スレッド ID: ${thread.getId()} を処理中...`);
    
    try {
      // スレッドのメッセージ詳細を取得し、最新のメッセージを解析対象とする
      const messages = thread.getMessages();
      if (messages.length === 0) {
        continue;
      }
      
      // 最新のメッセージを選択
      const latestMessage = messages[messages.length - 1];
      const subject = latestMessage.getSubject();
      
      // 本文の取得とトリミング（HTMLタグ除去や文字数制限）
      let body = latestMessage.getPlainBody();
      body = body.replace(/<[^>]*>/g, ' '); // 万が一HTMLが含まれていた場合の簡易タグ除去
      body = body.replace(/\s+/g, ' '); // 連続する余白を整形
      body = body.substring(0, 5000); // 最大5000文字
      
      console.log(`件名: ${subject}`);
      
      // Gemini APIで締め切り日を抽出
      const deadlineInfo = askGeminiForDeadline(GEMINI_API_KEY, subject, body);
      console.log(`Gemini抽出結果: ${JSON.stringify(deadlineInfo)}`);
      
      // 締め切り日が存在する場合、ラベルを作成して付与
      if (deadlineInfo && deadlineInfo.deadline && deadlineInfo.deadline !== 'NONE') {
        const labelName = formatLabelName(deadlineInfo.deadline);
        if (labelName) {
          // 日付文字列から曜日を取得
          const parts = deadlineInfo.deadline.split('-');
          const date = new Date(parts[0], parts[1] - 1, parts[2]);
          const dayOfWeekList = ['日', '月', '火', '水', '木', '金', '土'];
          const dayOfWeek = dayOfWeekList[date.getDay()];

          const dateLabel = getOrCreateLabel(labelName, dayOfWeek);
          thread.addLabel(dateLabel);
          console.log(`期限ラベル「${labelName}」を追加しました。`);
        }
        
        // カレンダーに登録
        try {
          createCalendarEvent(deadlineInfo.deadline, thread, deadlineInfo);
        } catch (calendarError) {
          console.error(`カレンダー登録中にエラーが発生しました: ${calendarError.message}`);
        }
      } else {
        console.log('締め切り日は検出されませんでした。');
      }
      
      // 処理済みラベルを付与
      thread.addLabel(processedLabel);
      console.log(`スレッド ${thread.getId()} の処理が完了しました。`);
      
    } catch (e) {
      console.error(`スレッド ${thread.getId()} の処理中にエラーが発生しました: ${e.message}`);
    }
    
    // Gemini APIの無料枠制限(5RPM)対策として、ループの最後に13秒（13000ms）待機
    if (i < threads.length - 1) {
      console.log('APIレート制限回避のため13秒間待機します...');
      Utilities.sleep(13000);
    }
  }
}

/**
 * ラベルを取得または新規作成する。新規作成した場合は色も設定する。
 */
function getOrCreateLabel(labelName, dayOfWeek) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    console.log(`ラベル「${labelName}」が存在しないため、新規作成します。`);
    label = GmailApp.createLabel(labelName);
    
    // 新規作成されたラベルに曜日に応じたカラーを設定
    if (dayOfWeek) {
      try {
        setLabelColor(label.getId(), dayOfWeek);
      } catch (colorError) {
        console.error(`ラベル「${labelName}」の色設定中にエラーが発生しました: ${colorError.message}`);
      }
    }
  }
  return label;
}

/**
 * YYYY-MM-DD から「MM/DD(曜日) まで」のラベル名を生成する
 * タイムゾーンのズレを防止するため、日付文字列から直接年・月・日を抽出します
 */
function formatLabelName(dateStr) {
  const parts = dateStr.split('-');
  if (parts.length !== 3) return null;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  
  const date = new Date(year, month, day); // 実行環境（JST）のタイムゾーンで生成
  if (isNaN(date.getTime())) {
    return null;
  }
  const mm = ('0' + (date.getMonth() + 1)).slice(-2);
  const dd = ('0' + date.getDate()).slice(-2);
  const dayOfWeekList = ['日', '月', '火', '水', '木', '金', '土'];
  const dayOfWeek = dayOfWeekList[date.getDay()];
  return `${mm}/${dd}(${dayOfWeek}) まで`;
}

/**
 * Gemini APIを呼び出して締め切り日を抽出
 */
function askGeminiForDeadline(apiKey, subject, body) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  
  // 現在の日付（曜日の算出や、メールに「本日」などの表現がある場合の解決用）
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  
  const systemInstruction = 'あなたは就職活動のメールから「アクションが必要な最終期限（締め切り日）」と「カレンダー用タスク名」を抽出する専門のアシスタントです。\n' +
    'メールの件名と本文から、エントリーシート(ES)の提出期限、WEBテストの受検期限、面接の予約期限など、ユーザーがアクションを起こさなければならない「最終締め切り日」を最優先で特定してください。\n' +
    'また、カレンダーの予定タイトルとしてふさわしい簡潔なタスク概要（例：「◯◯インターン 応募締め切り」、「◯◯会社 WEBテスト期限」）を生成してください。\n' +
    '就活に関係のないメールや、明確な期限（締め切り）が記載されていない場合は、期限の日付およびサマリーの代わりに "NONE" を返してください。';

  const prompt = `メールの受信基準日時(本日): ${todayStr}\n\n` +
    `【件名】\n${subject}\n\n` +
    `【本文】\n${body}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    systemInstruction: {
      parts: [{ text: systemInstruction }]
    },
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          deadline: {
            type: 'STRING',
            description: 'メール内の最終期限日付（YYYY-MM-DD形式）。期限がない、あるいは関係ない場合は "NONE" を返す。'
          },
          summary: {
            type: 'STRING',
            description: 'カレンダーに登録するための簡潔なタスク概要（例：「◯◯インターン 応募締め切り」、「◯◯会社 WEBテスト期限」）。期限がない場合は "NONE" を返す。'
          }
        },
        required: ['deadline', 'summary']
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Gemini API エラー (ステータス: ${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const json = JSON.parse(response.getContentText());
  
  try {
    const textResult = json.candidates[0].content.parts[0].text;
    return JSON.parse(textResult);
  } catch (e) {
    throw new Error(`Gemini API のレスポンス解析に失敗しました: ${e.message}. レスポンス: ${response.getContentText()}`);
  }
}

/**
 * 締め切り日を指定された名前のカレンダーに登録する。
 * カレンダーが存在しない場合は自動作成します。
 * @param {string} dateStr YYYY-MM-DD
 * @param {GmailThread} thread
 * @param {Object} deadlineInfo
 */
function createCalendarEvent(dateStr, thread, deadlineInfo) {
  const parts = dateStr.split('-');
  if (parts.length !== 3) return;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  const eventDate = new Date(year, month, day);

  const threadId = thread.getId();
  const threadUrl = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
  
  // カレンダーの取得または作成
  let calendar = null;
  const calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (calendars.length > 0) {
    calendar = calendars[0];
  } else {
    console.log(`カレンダー「${CALENDAR_NAME}」が存在しないため、新規作成します。`);
    calendar = CalendarApp.createCalendar(CALENDAR_NAME);
  }
  
  // 重複チェック: 対象日のイベントを取得
  const events = calendar.getEventsForDay(eventDate);
  const isDuplicate = events.some(event => {
    const desc = event.getDescription();
    return desc && desc.indexOf(`ThreadID:${threadId}`) !== -1;
  });
  
  if (isDuplicate) {
    console.log(`既にカレンダー「${CALENDAR_NAME}」に登録済みのスレッドです（Thread ID: ${threadId}）。スキップします。`);
    return;
  }
  
  // イベントタイトルの決定 (Geminiから得られた要約、またはメール件名)
  const taskSummary = (deadlineInfo && deadlineInfo.summary && deadlineInfo.summary !== 'NONE')
    ? deadlineInfo.summary
    : thread.getFirstMessageSubject();
  
  const eventTitle = `【就活期限】${taskSummary}`;
  const description = `メール件名: ${thread.getFirstMessageSubject()}\n` +
                      `詳細・メールリンク: ${threadUrl}\n\n` +
                      `[システム管理用メタデータ]\nThreadID:${threadId}`;
                      
  calendar.createAllDayEvent(eventTitle, eventDate, { description: description });
  console.log(`カレンダー「${CALENDAR_NAME}」に終日イベント「${eventTitle}」を追加しました (${dateStr})。`);
}

/**
 * 曜日に応じた背景色・文字色をGmail APIで設定する
 * @param {string} labelId ラベルのID
 * @param {string} dayOfWeek 曜日（'月'〜'日'）
 */
function setLabelColor(labelId, dayOfWeek) {
  // Gmail APIがサポートしているカラーパレットから設定
  const colors = {
    '月': { backgroundColor: '#b896e4', textColor: '#ffffff' }, // 紫
    '火': { backgroundColor: '#e68285', textColor: '#ffffff' }, // 赤
    '水': { backgroundColor: '#a2c2e8', textColor: '#ffffff' }, // 青
    '木': { backgroundColor: '#b3e1b3', textColor: '#ffffff' }, // 緑
    '金': { backgroundColor: '#fad165', textColor: '#ffffff' }, // 黄色
    '土': { backgroundColor: '#cccccc', textColor: '#ffffff' }, // グレー
    '日': { backgroundColor: '#cccccc', textColor: '#ffffff' }  // グレー
  };

  const color = colors[dayOfWeek];
  if (!color) return;

  const url = `https://gmail.googleapis.com/gmail/v1/users/me/labels/${labelId}`;
  const token = ScriptApp.getOAuthToken();
  const payload = {
    color: color
  };
  const options = {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    console.log(`ラベル「${labelId}」の色を「${dayOfWeek}」の色（背景色: ${color.backgroundColor}）に設定しました。`);
  } else {
    console.error(`ラベル「${labelId}」の色設定に失敗しました: ${response.getContentText()}`);
  }
}

