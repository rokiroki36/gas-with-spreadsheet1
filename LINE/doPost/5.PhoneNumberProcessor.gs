const PhoneNumberProcessor = {
 config: {
   SPREADSHEET_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1C6DMs',
   SHEET_NAME: '統一シート',
   PHONE_NUMBER_COLUMN: 'I',
   USER_ID_COLUMN: 'DS',
   DAY_OF_FIRST_COLUMN: 'BM',
   DAYS_THRESHOLD: 7,  // 初着金日が今日からこの日数以内なら新規客として、ユーザーID書き込み後メッセ自動遅延返信


   CACHE_DURATION: 10 * 60 * 1000,


   KEYWORD_FOR_FRESHERS: 'KEY_FOR_FRESH' // 新規入会者向け応答のキーワード（統一シートじゃなく自動応答シートにいれる）
 },


 _phoneMap: null,
 _lastCacheUpdate: 0,


 buildPhoneMap_: function() {
   try {


     const now = Date.now();
     if (this._phoneMap && now - this._lastCacheUpdate < this.config.CACHE_DURATION) {
       GeneralUtils.logDebug_('キャッシュ再利用');
       return;
     }


     const ss = SpreadsheetApp.openById(this.config.SPREADSHEET_ID);
     const sheet = ss.getSheetByName(this.config.SHEET_NAME);
     const data = sheet.getDataRange().getValues();


     this._phoneMap = {};
     for (let i = 0; i < data.length; i++) {
       try {
         const rawNumber = String(data[i][GeneralUtils.columnToNumber_(this.config.PHONE_NUMBER_COLUMN) - 1] || '');
         const cleanNumber = rawNumber.replace(/\D/g, '');


         if (cleanNumber.length >= 4) {
           const last4 = cleanNumber.slice(-4);
           if (!this._phoneMap[last4]) this._phoneMap[last4] = [];
           this._phoneMap[last4].push(i + 1); // 行番号は1から始まる
         }
       } catch (error) {
         GeneralUtils.logError_('行処理エラー', { row: i + 1, error: error.message });
       }
     }


     this._lastCacheUpdate = now;
     GeneralUtils.logDebug_('キャッシュ更新完了', {
       登録件数: Object.keys(this._phoneMap).length,
       最終更新: new Date(now).toISOString()
     });


   } catch (error) {
     GeneralUtils.logError_('buildPhoneMap_エラー', error);
     throw error;
   }
 },

 mainPhoneProcess_: async function(userInfo) {
   try {
     GeneralUtils.logDebug_('PhoneNumberProcessor.mainPhoneProcess_ 開始', userInfo);


     // 前処理済みテキストから数字抽出
     const processedText = userInfo.processedText || '';
     const last4 = processedText.slice(-4).replace(/\D/g, '');


     GeneralUtils.logDebug_('電話番号処理情報', {
       入力テキスト: userInfo.messageText,
       処理後テキスト: processedText,
       抽出4桁: last4
     });


     if (last4.length !== 4) {
       GeneralUtils.logDebug_('4桁数字不正');
       return {
         handled: false,
         code: 'INVALID_INPUT',
         debugMessage: "4桁の数字が検出されません"
       };
     }


     // キャッシュ構築
     this.buildPhoneMap_();


     const ss = SpreadsheetApp.openById(this.config.SPREADSHEET_ID);
     const sheet = ss.getSheetByName(this.config.SHEET_NAME);
     const userIdColumn = GeneralUtils.columnToNumber_(this.config.USER_ID_COLUMN);


     const matchingRows = this._phoneMap[last4] || [];
     GeneralUtils.logDebug_('マッチング検索結果', {
       検索キー: last4,
       候補数: matchingRows.length
     });


     const handleResult = await this.findMatchingCustomer_(sheet, matchingRows, userInfo.userId, userIdColumn);
      
     if (handleResult.code === 'ALREADY_REGISTERED') {
       return {
         handled: true,
         code: 'ALREADY_REGISTERED',
         debugMessage: "既に登録済みのユーザーです"
       };
     }


     return handleResult;


   } catch (error) {
     GeneralUtils.logError_('PhoneNumberProcessor全体エラー', {
       error: error.message,
       stack: error.stack,
       userInfo: userInfo
     });
     return {
       handled: false,
       code: 'FATAL_ERROR',
       debugMessage: `システムエラー: ${error.message}`
     };
   }
 },




 findMatchingCustomer_: async function(sheet, matchingRows, userId, userIdColumn) {
   try {


     if (matchingRows.length === 0) {
       GeneralUtils.logDebug_('照合結果: 該当なし');
       return { handled: true, code: 'NO_MATCH' };
     }


     if (matchingRows.length > 1) {
       GeneralUtils.logDebug_('複数候補あり', { 該当行: matchingRows });
       return { handled: true, code: 'MULTIPLE_MATCHES' };
     }


     const rowNumber = matchingRows[0];
     const existingUserId = sheet.getRange(rowNumber, userIdColumn).getValue();


     if (existingUserId) {


       // ★★★ 明示的に値を解決して返す ★★★
       return await Promise.resolve({ handled: true, code: 'ALREADY_REGISTERED' });
     }


     // ユーザーID登録
     sheet.getRange(rowNumber, userIdColumn).setValue(userId);
     GeneralUtils.logDebug_('ユーザーID登録完了', {
       行: rowNumber,
       ユーザーID: userId
     });


     // 新規客にファーストメッセージ送るため、ユーザーID入れた時と初着金日が7日以内かマッチングする関数にパス
     await this.isFresherCheck_(sheet, rowNumber, userId);


     return { handled: true, code: 'SUCCESS' };


   } catch (error) {
     GeneralUtils.logError_('findMatchingCustomer_エラー', error);
     throw error;
   }
 },




 // 新規会員判定とメッセージ送信処理
 // 日付判定のエラー防止と年またぎ判定
 isFresherCheck_: async function(sheet, rowNumber, userId) {
   GeneralUtils.logDebug_('isFresherCheck_呼び出し確認', {
     sheet: sheet,
     rowNumber: rowNumber,
     userId: userId
   });
   const firstDayCell = sheet.getRange(rowNumber, GeneralUtils.columnToNumber_(this.config.DAY_OF_FIRST_COLUMN));
   let firstDayValue = firstDayCell.getDisplayValue();


   // 空文字列対策: firstDayValueが空の場合は処理をスキップ
   if (!firstDayValue) {
     GeneralUtils.logDebug_('isFresherCheck_: firstDayValueが空のため処理スキップ', { rowNumber, userId });
     return;
   }


   GeneralUtils.logDebug_('isFresherCheck_内isWithinPastDays_呼び出し前', {
     firstDayValue: firstDayValue,
     DAYS_THRESHOLD: this.config.DAYS_THRESHOLD
   });
   if (this.isWithinPastDays_(firstDayValue, this.config.DAYS_THRESHOLD)) {
     const userInfo = {
       userId: userId,
       matchedKeyword: this.config.KEYWORD_FOR_FRESHERS,
       messageText: this.config.KEYWORD_FOR_FRESHERS
     };
     GeneralUtils.logDebug_('isFresherCheck_内AutoResponseProcessor.process_呼び出し確認', {
       userInfo: userInfo
     });
     await AutoResponseProcessor.process_(userInfo);
   }
 },




 // 日付判定ロジック
 isWithinPastDays_: function (dateString, days) {
   const [month, day] = dateString.split('/').map(Number);
   const today = new Date();
   const targetDate = new Date(today.getFullYear(), month - 1, day);


   // 年をまたぐ場合を考慮
   if (targetDate > today) {
     targetDate.setFullYear(today.getFullYear() - 1);
   }


   const diff = today.getTime() - targetDate.getTime();
   const diffDays = diff / (1000 * 60 * 60 * 24);


   return diffDays <= days;
 },




 //手動で電話番号マップ？のキャッシュ更新
 forceRefreshCache: function() {
   try {
     this._phoneMap = null;
     this.buildPhoneMap_();
     GeneralUtils.logDebug_('手動キャッシュ更新完了');
   } catch (error) {
     GeneralUtils.logError_('手動キャッシュ更新失敗', error);
   }
 }
};




function forceRefreshCache() {
 PhoneNumberProcessor.forceRefreshCache();
}

