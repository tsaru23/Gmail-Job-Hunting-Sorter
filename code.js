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

// 検索クエリ
const SEARCH_QUERY = 'is:unread -label:処理済み (subject:("締め切り" OR "締切" OR "期限" OR "提出" OR "予約" OR "受験" OR "受検" OR "テスト") OR "エントリーシート" OR "ES") ("インターン" OR "選考" ) -("メルマガ" OR "ニュースレター" OR "コラム" OR "マガジン")';

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
          const dateLabel = getOrCreateLabel(labelName);
          thread.addLabel(dateLabel);
          console.log(`期限ラベル「${labelName}」を追加しました。`);
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
 * ラベルを取得または新規作成する
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    console.log(`ラベル「${labelName}」が存在しないため、新規作成します。`);
    label = GmailApp.createLabel(labelName);
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
  
  const systemInstruction = 'あなたは就職活動のメールから「アクションが必要な最終期限（締め切り日）」を抽出する専門のアシスタントです。\n' +
    'メールの件名と本文から、エントリーシート(ES)の提出期限、WEBテストの受検期限、面接の予約期限など、ユーザーがアクションを起こさなければならない「最終締め切り日」を最優先で1つ特定してください。\n' +
    '就活に関係のないメールや、明確な期限（締め切り）が記載されていない場合は、期限の日付の代わりに "NONE" を返してください。';

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
          }
        },
        required: ['deadline']
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
