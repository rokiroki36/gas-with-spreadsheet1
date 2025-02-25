/**
* =====================================================
* LINE遅延応答Bot - AutoResponseProcessor
*
* 主な機能：
* - キーワードに基づく自動応答
* - 即時応答と遅延応答の分岐処理
* - メッセージキュー処理
* - 定期メンテナンス
* - 応答済みフラグ管理
* - プレースホルダー置換
* =====================================================
*/

const AutoResponseProcessor = {
  // システム全体の設定
  CONFIG: {
    SPREADSHEET_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1C6DMs',
    SHEET_NAMES: {
      AUTO_RESPONSE: '自動応答シート',  // 自動応答の設定用シート
      UNIFIED: '統一シート'            // 顧客情報管理シート
    },
    COLUMN_NAMES: {
      KEYWORD: 'B',      // キーワード列（自動応答シート）
      RESPONSE1: 'C',    // 送信内容1
      ALT_TEXT_1:'D',     // 送信内容1の代替テキスト
      RESPONSE2: 'E',    // 送信内容2
      ALT_TEXT_2: 'F',      // 送信内容2の代替テキスト
      
      DELAY: 'G',        // 遅延時間（分）列（自動応答シート）
      FLAG_COLUMN: 'H',  // フラグを立てる列を指定する列（自動応答シート）（E列内の値は、変更しても毎回再デプロイしなくて良い）

      USER_ID: 'DS',      // LINEユーザーID列（統一シート）
      // 統一シートの、姓名とメアドの列
      SEI: 'F',
      MEI: 'G',
      MAIL: 'D'

    }
  },

/**
* スプレッドシートから応答設定を取得
* キーワードに対して最大2通の応答内容を取得
* プレースホルダーが含まれる場合は置換する
* @param {string} userId - LINEのユーザーID
* @returns {Promise<Object>} - キーワードとメッセージのペア
*/
getMessagePairs_: async function(userId) {
  GeneralUtils.logDebug_('getMessagePairs_ 開始 - 引数確認', { userId: userId });
  const ss = SpreadsheetApp.openById(this.CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(this.CONFIG.SHEET_NAMES.AUTO_RESPONSE);
  const data = sheet.getDataRange().getValues();

  // スプレッドシートの変更を強制的に反映
  SpreadsheetApp.flush();

  // 統一シートから、ターゲットにすべき顧客のsei,meiを取得
  const placeholders = await this.getUserInfo_(userId);
  

  const messagePairs = {};

  // 1行目はヘッダーなのでスキップ
  for (let i = 1; i < data.length; i++) {
    
    const cols = {
      keyword: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.KEYWORD) - 1,
      response1: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.RESPONSE1) - 1,
      response2: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.RESPONSE2) - 1,
      delay: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.DELAY) - 1,
      flag: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.FLAG_COLUMN) - 1,
      altText1: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.ALT_TEXT_1) - 1,
      altText2: GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.ALT_TEXT_2) - 1
    };

    const keyword = data[i][cols.keyword];
    const response1 = data[i][cols.response1];
    const response2 = data[i][cols.response2];
    const delayMinutes = data[i][cols.delay];
    const flagColumn = data[i][cols.flag];

    if (keyword && response1) {
      messagePairs[keyword] = {
        messages: [
          {
            content: response1.startsWith('https://docs.google.com/')
              ? await this.getDocumentContent_(response1, placeholders)
              : this.replacePlaceholders_(response1, placeholders),
            isDocument: response1.startsWith('https://docs.google.com/'),
            altText: this.replacePlaceholders_(data[i][cols.altText1] || '画像メッセージ', placeholders)
          }
        ],
        delayMinutes: delayMinutes || 0,
        flagColumn: flagColumn
      };


      // 2通目の応答内容が存在する場合は追加
      if (response2) {
        messagePairs[keyword].messages.push({
          content: response2.startsWith('https://docs.google.com/')
            ? await this.getDocumentContent_(response2, placeholders)
            : this.replacePlaceholders_(response2, placeholders),
          isDocument: response2.startsWith('https://docs.google.com/'),
          altText: this.replacePlaceholders_(data[i][cols.altText2] || '画像メッセージ', placeholders)
        });
        GeneralUtils.logDebug_('getMessagePairs_ - 2通目の応答内容を追加', {
          keyword: keyword,
          messages: messagePairs[keyword].messages
        });
      }
    } else {
      GeneralUtils.logDebug_('getMessagePairs_ - キーワードまたは応答内容1が空のためスキップ', { 行番号: i });
    }
  }

  GeneralUtils.logDebug_('getMessagePairs_ 完了', {
    取得キーワード数: Object.keys(messagePairs).length,
    キーワード一覧: Object.keys(messagePairs)
  });

  return messagePairs;
},

/**
* 統一シートからユーザー情報を取得する
* @param {string} userId - LINEのユーザーID
* @return {Promise<Object>} - プレースホルダーに対応する情報を含むオブジェクト
*/
getUserInfo_: async function(userId) {
  const ss = SpreadsheetApp.openById(this.CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(this.CONFIG.SHEET_NAMES.UNIFIED);
  const userIdColumn = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.USER_ID);
  const seiColumn = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.SEI);
  const meiColumn = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.MEI);
  const mailColumn = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.MAIL);

  // ユーザーIDから該当する行を検索
  const userIds = sheet.getRange(1, userIdColumn, sheet.getLastRow()).getValues().flat();
  const rowIndex = userIds.indexOf(userId);

  if (rowIndex === -1) {
    // ユーザーが見つからない場合は空のオブジェクトを返す
    return { sei: '', mei: '', mail: '' };
  }

  // ユーザー情報（姓、名、メールアドレス）を取得
  const sei = sheet.getRange(rowIndex + 1, seiColumn).getValue();
  const mei = sheet.getRange(rowIndex + 1, meiColumn).getValue();
  const mail = sheet.getRange(rowIndex + 1, mailColumn).getValue();

  GeneralUtils.logDebug_('ユーザー情報取得', { userId, sei, mei, mail });
  return { sei, mei, mail };
},



/**
* Googleドキュメントから本文を取得し、プレースホルダーを置換する
* @param {string} documentUrl - ドキュメントのURL
* @param {Object} placeholders - プレースホルダーと値のペア
* @return {Promise<string>} - プレースホルダーが置換されたドキュメントの本文
*/
getDocumentContent_: async function(documentUrl, placeholders) {
  try {
    const doc = DocumentApp.openByUrl(documentUrl.trim());
    const docContent = doc.getBody().getText();
    return this.replacePlaceholders_(docContent, placeholders);
  } catch (error) {
    GeneralUtils.logError_('ドキュメント処理エラー', error);
    throw error;
  }
},

/**
* 文字列内のプレースホルダーを置換する
* @param {string} text - プレースホルダーを含む文字列
* @param {Object} placeholders - プレースホルダーと値のペア
* @return {string} - プレースホルダーが置換された文字列
*/
replacePlaceholders_: function(text, placeholders) {
  if (!text) return '';

  let replacedText = text;
  for (const key in placeholders) {
    const regex = new RegExp(`\\{${key}\\}`, 'g');
    replacedText = replacedText.replace(regex, placeholders[key]);
  }
  return replacedText;
},


 /**
 * メッセージ処理のメインロジック
 * キーワードマッチ後の応答処理全体を制御
 * 
 * @param {Object} userInfo ユーザー情報（必須項目：userId, matchedKeyword）
 * - PhoneNumberProcessor.gsからの呼び出し時：KEY_FOR_FRESHがmatchedKeywordとして渡される
 * - 通常の自動応答時：ユーザー入力に対するキーワードマッチング結果が渡される
 * @returns {Promise<Object>} 処理結果
 */
process_: async function(userInfo) {
  GeneralUtils.logDebug_('AutoResponseProcessor開始', userInfo);

  try {
    if (!userInfo.userId || !userInfo.matchedKeyword) {
      return {
        handled: false,
        debugMessage: 'キーワードマッチ情報がありません'
      };
    }

    // すでに応答済みかチェック
    try {
      const messagePairs = await this.getMessagePairs_(userInfo.userId);
      const responseData = messagePairs[userInfo.matchedKeyword];
      
      if (!responseData || !responseData.flagColumn) {
        GeneralUtils.logDebug_('フラグ列指定なし、処理継続', userInfo.matchedKeyword);
      } else {
        // 統一シートで該当ユーザーのフラグ確認
        const ss = SpreadsheetApp.openById(this.CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(this.CONFIG.SHEET_NAMES.UNIFIED);
        const userIdCol = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.USER_ID);
        const flagCol = GeneralUtils.columnToNumber_(responseData.flagColumn);
        
        // ユーザーIDで該当行を検索
        const userIds = sheet.getRange(1, userIdCol, sheet.getLastRow()).getValues();
        const rowIndex = userIds.findIndex(row => row[0] === userInfo.userId);

        if (rowIndex !== -1) {
          const flagValue = sheet.getRange(rowIndex + 1, flagCol).getValue();
          if (flagValue === 1) {
            GeneralUtils.logDebug_('既に応答済みのユーザー', {
              userId: userInfo.userId,
              keyword: userInfo.matchedKeyword
            });
            return {
              handled: true,
              debugMessage: '既に応答済みのため、処理をスキップしました',
              alreadyResponded: true
            };
          }
        }
      }
    } catch (error) {
      GeneralUtils.logError_('応答済みチェック時にエラー', error);
      // エラー時は安全のため処理を継続
    }

    // PhoneNumberProcessorから呼ばれた場合も、通常の自動応答と同様に
    // matchedKeyword（この場合はKEY_FOR_FRESH）を使って
    // 自動応答シートから対応するメッセージを取得
    const messagePairs = await this.getMessagePairs_(userInfo.userId);
    const responseData = messagePairs[userInfo.matchedKeyword];

    if (!responseData) {
      GeneralUtils.logDebug_('キーワードに対応するメッセージが未設定', userInfo.matchedKeyword);
      return {
        handled: false,
        debugMessage: 'メッセージ設定が見つかりません'
      };
    }

    const messages = responseData.messages;
    const delayMinutes = responseData.delayMinutes || 0;
    const flagColumn = responseData.flagColumn;

    if (delayMinutes === 0) {
      try {
        // 即時応答の場合
        const accessToken = LineUtils.getAccessToken_();

        for (const msg of messages) {
          if (msg.content) {
            await LineUtils.sendMessage_(
              userInfo.userId, 
              { 
                content: msg.content, 
                altText: msg.altText 
              }, 
              accessToken
            );
          } else {
            GeneralUtils.logDebug_('メッセージ内容が undefined です', msg);
          }
        }

        await this.updateResponseFlag_(userInfo.userId, flagColumn, true);
        await this.logAutoResponse_(userInfo, responseData, true);

        return {
          handled: true,
          debugMessage: `キーワード「${userInfo.matchedKeyword}」に対して即時応答しました。`,
          keyword: userInfo.matchedKeyword,
          immediate: true
        };

      } catch (error) {
        GeneralUtils.logError_('即時メッセージ送信エラー', error);
        throw error;
      }

    } else { 
      // 遅延応答の場合
      const scheduledTime = new Date();
      scheduledTime.setMinutes(scheduledTime.getMinutes() + delayMinutes);

      for (let i = 0; i < messages.length; i++) {
        if (messages[i].content) {
          await this.queueMessage_(
            userInfo.userId,
            { 
              content: messages[i].content, 
              altText: messages[i].altText 
            },
            scheduledTime,
            i === messages.length - 1 ? flagColumn : null,  // 最後のメッセージにのみフラグ列を設定
            i * 10  // メッセージごとに10ミリ秒ずらす
          );
        } else {
          GeneralUtils.logDebug_('メッセージ内容が undefined です', messages[i]);
        }
      }

      await this.logAutoResponse_(userInfo, responseData, false);

      return {
        handled: true,
        debugMessage: `キーワード「${userInfo.matchedKeyword}」に対する応答をキューに追加しました。`,
        keyword: userInfo.matchedKeyword,
        scheduledDelay: delayMinutes
      };
    }

  } catch (error) {
    GeneralUtils.logError_('AutoResponseProcessorエラー', error);
    return {
      handled: false,
      debugMessage: 'エラーが発生しました: ' + error.message
    };
  }
},


/**
 * メッセージをキューに追加
 * 遅延メッセージを一時保管し、後続の処理で送信
 * @param {string} userId ユーザーID
 * @param {Object} messageObj メッセージオブジェクト (content, altText を含む)
 * @param {Date} scheduledTime 送信予定時刻 (Date オブジェクト)
 * @param {string} flagColumn フラグを立てる列
 * @param {number} offsetTime createdAt をずらす時間 (ミリ秒)
 * @private
 */
queueMessage_: async function(userId, messageObj, scheduledTime, flagColumn, offsetTime = 0) {
  // キューに保存するメッセージデータを作成
  const messageData = {
    userId: userId,
    message: messageObj, // メッセージオブジェクトを保存 (content, altText を含む)
    scheduledTime: scheduledTime.getTime(),
    createdAt: new Date().getTime() + offsetTime, // offsetTime を加算
    flagColumn: flagColumn
  };

  // PropertiesServiceを使用してキューを取得・更新
  const properties = PropertiesService.getScriptProperties();
  const currentQueue = JSON.parse(properties.getProperty('MESSAGE_QUEUE') || '[]');
  currentQueue.push(messageData);
  properties.setProperty('MESSAGE_QUEUE', JSON.stringify(currentQueue));

  GeneralUtils.logDebug_('メッセージをキューに追加', messageData);
},

/**
 * 送信予定メッセージの処理
 * キューに保存された遅延メッセージを定期的にチェックして送信
 * トリガーで定期的に実行される
 */
processQueue_: async function() {
  GeneralUtils.logDebug_('キュー処理開始');

  const properties = PropertiesService.getScriptProperties();
  const currentQueue = JSON.parse(properties.getProperty('MESSAGE_QUEUE') || '[]');
  const now = new Date().getTime();
  const remainingMessages = [];

  for (const msg of currentQueue) {
    if (msg.scheduledTime <= now) {
      try {
        const accessToken = LineUtils.getAccessToken_();
        await LineUtils.sendMessage_(msg.userId, msg.message, accessToken);
        
        if (msg.flagColumn) {
          await this.updateResponseFlag_(msg.userId, msg.flagColumn, false);
        }

        GeneralUtils.logDebug_('メッセージ送信成功', msg.userId);
      } catch (error) {
        GeneralUtils.logError_('メッセージ送信エラー', error);
        remainingMessages.push(msg);
      }
    } else {
      remainingMessages.push(msg);
    }
  }

  properties.setProperty('MESSAGE_QUEUE', JSON.stringify(remainingMessages));
  GeneralUtils.logDebug_('キュー処理完了', `残りのメッセージ数: ${remainingMessages.length}`);
},

/**
 * 応答後のフラグ更新処理
 * 統一シート内の指定された列に応答済みフラグを設定
 * @param {string} userId LINE ユーザーID
 * @param {string} flagColumn フラグを立てる列（アルファベット）
 * @param {boolean} immediate 即時応答かどうか
 * @private
 */
updateResponseFlag_: async function(userId, flagColumn, immediate) {
  try {
    // フラグ列が指定されていない場合は処理をスキップ
    if (!flagColumn) {
      GeneralUtils.logDebug_('フラグ列が指定されていません');
      return;
    }

    const ss = SpreadsheetApp.openById(this.CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(this.CONFIG.SHEET_NAMES.UNIFIED);

    // 列番号の計算（アルファベット → 数値）
    const userIdCol = GeneralUtils.columnToNumber_(this.CONFIG.COLUMN_NAMES.USER_ID);
    const flagCol = GeneralUtils.columnToNumber_(flagColumn);

    // ユーザーIDで該当行を検索
    const userIds = sheet.getRange(1, userIdCol, sheet.getLastRow()).getValues();
    const rowIndex = userIds.findIndex(row => row[0] === userId);

    if (rowIndex !== -1) {
      // 該当行が見つかった場合、フラグを設定（1を入力）
      const actualRow = rowIndex + 1;  // 0始まりのインデックスを1始まりの行番号に変換
      sheet.getRange(actualRow, flagCol).setValue(1);

      GeneralUtils.logDebug_('応答フラグを更新しました', {
        userId: userId,
        row: actualRow,
        flagColumn: flagColumn,
        immediate: immediate
      });
    } else {
      GeneralUtils.logDebug_('ユーザーIDに該当する行が見つかりません', userId);
    }
  } catch (error) {
    GeneralUtils.logError_('updateResponseFlag_', error);
  }
},

/**
* 自動応答のログを記録する
* @param {Object} userInfo - ユーザー情報
* @param {Object} responseData - 応答データ
* @param {boolean} isImmediate - 即時応答かどうか
* @private
*/
logAutoResponse_: async function(userInfo, responseData, isImmediate) {
  try {
    const ss = SpreadsheetApp.openById(this.CONFIG.SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('自動送信ログ');

    // 現在の日時を取得
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    // ユーザーのLINE名を取得
    const profile = await LineUtils.getUserName_(userInfo.userId);
    const userName = profile.success ? profile.displayName : "不明";

    // 送信内容1のログ
    let message1Log = responseData.messages[0].content;
    if (responseData.messages[0].content.startsWith('https://docs.google.com/')) {
      // Google ドキュメントの URL の場合、最初の50文字を表示
      message1Log = responseData.messages[0].content.replace(/\n/g, ' ').slice(0, 50);
    } else if (responseData.messages[0].content.match(/\.(jpeg|jpg|gif|png)$/i)) {
      // 画像の場合は URL をそのまま表示
      // message1Log はそのまま
    } else {
      // テキストメッセージの場合、最初の50文字を表示
      message1Log = responseData.messages[0].content.replace(/\n/g, ' ').slice(0, 50);
    }

    // 送信内容2のログ (送信内容2が存在する場合のみ)
    let message2Log = '';
    if (responseData.messages.length > 1) {
      message2Log = responseData.messages[1].content;
      if (responseData.messages[1].content.startsWith('https://docs.google.com/')) {
        // Google ドキュメントの URL の場合、最初の50文字を表示
        message2Log = responseData.messages[1].content.replace(/\n/g, ' ').slice(0, 50);
      } else if (responseData.messages[1].content.match(/\.(jpeg|jpg|gif|png)$/i)) {
        // 画像の場合は URL をそのまま表示
        // message2Log はそのまま
      } else {
        // テキストメッセージの場合、最初の50文字を表示
        message2Log = responseData.messages[1].content.replace(/\n/g, ' ').slice(0, 50);
      }
    }

    // ログデータの作成
    const logData = [
      timestamp,                    // A列: タイムスタンプ
      userInfo.userId,             // B列: ユーザーID（全文字）
      userName,                    // C列: ユーザーのLINE名
      userInfo.matchedKeyword,     // D列: マッチしたキーワード
      message1Log,                 // E列: 送信内容1
      message2Log,                 // F列: 送信内容2
      isImmediate ? 0 : responseData.delayMinutes  // G列: 遅延時間（分）
    ];

    // ログの追記
    logSheet.appendRow(logData);

    GeneralUtils.logDebug_('自動応答ログを記録しました', logData);
  } catch (error) {
    GeneralUtils.logError_('logAutoResponse_', error);
  }
},

  /**
  * 古いメッセージキューの削除をやる関数
  */
  dailyMaintenanceForAR: function () {
    GeneralUtils.logDebug_('dailyMaintenance 開始');

    try {
      const properties = PropertiesService.getScriptProperties();
      const now = new Date().getTime();
      const TWO_DAYS_MS = 48 * 60 * 60 * 1000;  // 2日間のミリ秒

      // メッセージキューの整理（2日以上経過したものを削除）
      const currentQueue = JSON.parse(properties.getProperty('MESSAGE_QUEUE') || '[]');
      const updatedQueue = currentQueue.filter(msg => {
        return (now - msg.createdAt) < TWO_DAYS_MS;
      });
      properties.setProperty('MESSAGE_QUEUE', JSON.stringify(updatedQueue));

      GeneralUtils.logDebug_('dailyMaintenanceForAR 正常終了');

    } catch (error) {
      GeneralUtils.logError_('dailyMaintenanceForAR エラー', error);
    }
  }
};

/**
* トリガーから見えるようにするため（ラッパー関数）
*/
function processQueueOfAR() {
  AutoResponseProcessor.processQueue_();
}

function dailyMaintenanceForAR() {
  AutoResponseProcessor.dailyMaintenanceForAR();
}

/**
* トリガー作るためのおまけ
* 時間間隔設定とか自動でやってくれる
*/
function setupTriggersForAutoResponse() {
  ScriptApp.newTrigger('processQueueOfAR')
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger('dailyMaintenanceForAR')
    .timeBased()
    .everyDays(1)
    .atHour(3) // 毎日午前3時に実行
    .create();
};

