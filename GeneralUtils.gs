//スプシ開くことをスタートとするトリガーで叩くために露出
function onOpen() {
    return GeneralUtils.onOpen();
}



/**
* =====================================================
* GeneralUtils namespace
* 共通で使用するユーティリティ関数をまとめたネームスペース
* =====================================================
*/
const GeneralUtils = {

/**
* =====================================================
* スプシ開いた時に自動実行される関数
* =====================================================
*/
  onOpen: function() {
      const ui = SpreadsheetApp.getUi();
      
      ui.createMenu('任意送信関連')
          .addItem('メール送信', 'manuGmailMain')  // メール送信時にアカウントチェックする関数からスタート
          .addItem('LINE送信', 'manuLINEMain')
          .addToUi();
  },
  
/**
* =====================================================
* Googleドライブ系
* =====================================================
*/
/**
 * GoogleドライブのURLからファイルIDを取得する関数
 * @param {string} url - GoogleドライブのURL
 * @return {string} ファイルID
 */
  getFileIdFromDriveUrl_: function(url) {
  // 共有リンクからファイルIDを抽出
  let fileId;
  if (url.includes('/file/d/')) {
    fileId = url.match(/\/file\/d\/([^\/]+)/)[1];
    fileId = fileId.split('/')[0]; // viewの前まで
  } else if (url.includes('/document/d/')) {
    fileId = url.match(/\/document\/d\/([^\/]+)/)[1];
    // クエリパラメータを除去
    fileId = fileId.split('?')[0];  
  } else if (url.includes('/spreadsheets/d/')) {
    fileId = url.match(/\/spreadsheets\/d\/([^\/]+)/)[1];
    fileId = fileId.split('?')[0];  
  } else if (url.includes('/presentation/d/')) {
    fileId = url.match(/\/presentation\/d\/([^\/]+)/)[1];
    fileId = fileId.split('?')[0];  
  } else if (url.includes('?id=')) {
    fileId = url.match(/\?id=([^\&]+)/)[1];
  } else {
    throw new Error('無効なGoogleドライブのURLです: ' + url);
  }
  return fileId;
},

/**
 * driveのファイルID投げたらdocsとか画像とか判定してくれる
 */
 getMimeType_ : function(fileId){
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    return mimeType //
 },


 /**
  * Google Driveの画像URLから画像のサイズを取得する
  * @param {string} url - Google Driveの画像URL
  * @return {object} - 画像の幅と高さ（width, height）を含むオブジェクト
  * @throws {Error} - URLが無効な場合、またはファイルIDを抽出できない場合
  *                   または画像のメタデータを取得できない場合
  */
 getImageDimensions_: function(url) {
   if (!url || typeof url !== 'string') {
     throw new Error("URL must be a non-empty string");
   }

   try {
     // URLからファイルIDを抽出
     const patterns = [
       /\/d\/([^\/\?]+)/, // /d/[fileId] パターン
       /id=([^&]+)/, // id=[fileId] パターン
     ];

     let fileId = null;
     for (const pattern of patterns) {
       const match = url.match(pattern);
       if (match && match[1]) {
         fileId = match[1];
         break;
       }
     }

     if (!fileId) {
       throw new Error("Could not extract file ID from Google Drive URL");
     }

     // Google Drive APIを使用して画像のメタデータを取得
     const file = Drive.Files.get(fileId, { fields: 'imageMediaMetadata' });

     if (file.imageMediaMetadata) {
       const dimensions = {
         width: file.imageMediaMetadata.width,
         height: file.imageMediaMetadata.height
       };
       Logger.log(`Image dimensions: Width=${dimensions.width}px, Height=${dimensions.height}px`);
       return dimensions;
     } else {
       throw new Error("Could not retrieve image dimensions from metadata");
     }

   } catch (error) {
     Logger.log(`Error: ${error.message}`);
     throw error;
   }
 },


 /**
  * 列名（アルファベット）から列番号を取得する関数
  * 例: A -> 1, B -> 2, Z -> 26, AA -> 27
  * @param {string} column - 列名（アルファベット）
  * @return {number} - 列番号
  * @private
  */
 columnToNumber_: function(column) {
   column = column.toUpperCase();
   let result = 0;
  
   for (let i = 0; i < column.length; i++) {
     result = result * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
   }
   return result;
 },


/**
* =====================================================
* ログ関連の設定
* =====================================================
*/
 LOGGING_CONFIG: {
   SPREADSHEET_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1CXXXX',
   SHEET_NAMES: {
     DEBUG: 'Debug',
     ALL_EVENTS: 'AllEventsLog'
   }
 },

 /**
  * キャッシュ用の変数
  */
 _debugSheet: null,

 /**
  * シートの参照を取得
  * @private
  */
 _getDebugSheet: function() {
   if (!this._debugSheet) {
     this._debugSheet = SpreadsheetApp.openById(this.LOGGING_CONFIG.SPREADSHEET_ID)
                                    .getSheetByName(this.LOGGING_CONFIG.SHEET_NAMES.DEBUG);
   }
   return this._debugSheet;
 },


 /**
  * デバッグログを記録
  * @param {string} message - ログメッセージ
  * @param {any} data - 追加のデータ（オプション）
  * @private
  */
 logDebug_: function(message, data = null) {
   try {
     const logData = [new Date(), message];
    
     if (data) {
       try {
         if (typeof data === 'object') {
           const stringified = JSON.stringify(data);
           logData.push(stringified.length > 1000 ?
             stringified.substring(0, 1000) + '...' :
             stringified);
         } else {
           logData.push(String(data));
         }
       } catch (e) {
         logData.push('[データの変換に失敗]');
       }
     }


     // 即時書き込み
     const sheet = this._getDebugSheet();
     sheet.appendRow(logData);
    
   } catch (error) {
     console.error('Log write failed:', error);
   }
 },


 /**
  * エラーログを記録
  * @param {string} location - エラーが発生した場所
  * @param {Error} error - エラーオブジェクト
  * @private
  */
 logError_: function(location, error) {
   try {
     const sheet = this._getDebugSheet();
    
     sheet.appendRow([
       new Date(),
       `Error in ${location}`,
       error.message,
       error.stack || 'スタックトレースなし'
     ]);
   } catch (e) {
     console.error('Error logging failed:', e);
   }
 },


 /**
  * イベントログをバッチ処理
  * @param {Array} events - イベントの配列
  * @param {Object} sheet - 書き込み先のシート
  */
processBatchEvents_: function(events, sheet) {
    if (!events || events.length === 0) return;

    try {
      const batchData = events.map(event => {
        const messageContent = this._formatMessageContent(event);
        return [
          this.formatDate_(new Date(event.timestamp), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
          event.source.userId || 'Unknown',
          event.message ? event.message.type : 'N/A', // event.message が存在するか確認
          event.message ? JSON.stringify(event.message) : 'N/A', // event.message が存在するか確認
          messageContent
        ];
      });

      if (batchData.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, batchData.length, batchData[0].length)
            .setValues(batchData);
      }
    } catch (error) {
      this.logError_('processBatchEvents', error);
    }
  },


 /**
  * メッセージ内容のフォーマット（内部ヘルパー）
  * @private
  */
 _formatMessageContent: function(event) {
   if (!event.message) return '';


   switch(event.message.type) {
     case 'text':
       return event.message.text;
     case 'sticker':
       return `[Sticker] PackageID:${event.message.packageId}, StickerID:${event.message.stickerId}`;
     case 'image':
       return '[Image]';
     default:
       return `[${event.message.type}]`;
   }
 },


 /**
  * 日時を指定したフォーマットで整形
  * @param {Date} date - 日時オブジェクト
  * @param {string} timeZone - タイムゾーン（例: 'Asia/Tokyo'）
  * @param {string} format - 日時フォーマット
  * @returns {string} フォーマットされた日時文字列
  * @private
  */
 formatDate_: function(date, timeZone, format) {
   return Utilities.formatDate(date, timeZone, format);
 }
};
