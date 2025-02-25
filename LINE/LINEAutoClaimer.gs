/**
 * =====================================================
 * LINEで、自動で、振込日リマインドや滞納者督促を行う機能
 * =====================================================
 */

/**
* 最後の関数で設置されたトリガーからLINEメイン関数を叩けるように
* 名前空間より外に出す
*/
function lineClaimMain() {
  LineClaimSpace.lineClaimMain();
}

/**
* 名前空間開始
*/
const LineClaimSpace = (function() {  // 即時実行関数の開始

/**
* このファイル全体の設定
*/
const CONFIG_CLAIM_L = {

  // 基本設定
  DEBUG_MODE: false,                 // true: デバッグモード, false: 通常モード（本番！モード）
  TIME_ZONE: 'Asia/Tokyo',         // タイムゾーン設定
  DATE_FORMAT: 'yyyy/MM/dd',       // 日付の表示形式

  // LINE API関連の設定
  LINE: {
    API_ENDPOINT: 'https://api.line.me/v2/bot/message/push',
    TOKEN_PROPERTY_KEY: 'LINE_ACCESS_TOKEN'
  },

  // スプレッドシート関連の設定
  SPREADSHEET_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1C6DMs',  // 操作対象のスプレッドシートID
  SHEET_NAME: '統一シート',                                          // メインの操作対象シート名
  SETTEI_SHEET_NAME: '督促送信内容', // 各種設定があるシート名

  PATTERNS: {
    PATTERN1: '松尾',
    PATTERN2: '中村'
  },

 // 統一シートの列設定 フラグや客情報および、督促日付がある一番左の列
  TARGET_COLUMNS: {
    LINE_USER_ID: 'DS',   // LINEユーザーIDが記録されている列
    LAST_NAME: 'F',      // 姓が記録されている列
    FIRST_NAME: 'G',     // 名が記録されている列
    SENTDATE_OF_L: 'ER', // LINE送信日を記録する列
    CUSTOMER_PATTERN: 'EP' //顧客パターン列（今は松尾か中村か）
  },
  TOKUSOKUBI_COLUMN_START: 'EV',  // 督促対象日のデータがある一番左の列

  // 督促送信内容シートの列設定
  MESSAGE_COLUMNS: {
    TIMING: 'A',        // 督促タイミングの名前が入ってる（この名前の文字列をプログラムが認識して、各関数に受け渡す）
    OFFSET:'B',
    CONTENT_1: 'C',     // 送信内容1
    ALT_TEXT_1: 'D',     // 送信内容1の代替テキスト
    CONTENT_2: 'E',     // 送信内容2
    ALT_TEXT_2: 'F',     // 送信内容2の代替テキスト
    PATTERN: 'O' , //松尾か中村か
    IMAGE_HEIGHT_1: 'Q', //googledrive直URLを原稿に入れて、縦長リッチメッセ送りたいときは縦幅指定（1通目）
    IMAGE_HEIGHT_2: 'R', // 2通目縦幅
  }
};


/**
* LINEのアクセストークンを取得
* @returns {string} LINEチャネルアクセストークン
* @private
*/
function getLineAccessToken_() {
  const token = PropertiesService.getScriptProperties().getProperty(
    CONFIG_CLAIM_L.LINE.TOKEN_PROPERTY_KEY
  );
  if (!token) {
    throw new Error('LINEチャネルアクセストークンが設定されていません。');
  }
  return token;
}

// --- ここからメイン処理関連 ---

/**
* 督促処理のメイン関数(LINE)
*/
function lineClaimMain() {
  // スプレッドシートとシートを取得
  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_L.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_L.SHEET_NAME);

  // 今日が、処理を走らせる日かどうかをチェック
  let tokusokuDates = []; // 初期化
  try {
    tokusokuDates = shouldTakeActionToday_(sheet);
  } catch (e) {
    // isTodayShouldAction_ がエラーをスローした場合の処理
    Logger.log(e.message); // エラーメッセージを出力
    Logger.log("処理を中断します。");
    return; // 処理を中断
  }

  // 処理を走らせる日であれば、それぞれの督促日（の情報）について、findTargetCustomers_関数で、ターゲットにすべき顧客を特定して、LINEメッセージを作成・送信
  if (tokusokuDates.length > 0) {
    tokusokuDates.forEach(tokusokuDate => {
      let recordSentData = findTargetCustomers_(sheet, tokusokuDate);
      if(recordSentData.length > 0){

        //履歴入れる関数に、LINE送信したよって情報渡す
        recordSentDate(recordSentData, 'LINE');
      }
    });
  } else {
    // 今日が処理すべき日でない場合のログ出力（エラーの場合はここは実行されない）
    const today = new Date();
    const kyouFormatted = Utilities.formatDate(today, CONFIG_CLAIM_L.TIME_ZONE, CONFIG_CLAIM_L.DATE_FORMAT) + '(' + ['日', '月', '火', '水', '木', '金', '土'][today.getDay()] + ')';
    Logger.log(`今日${kyouFormatted}は、アクション日ではないです`);
  }
}



/**
* 今日、処理を走らせるべきかを判断する関数
* @param {SpreadsheetApp.Sheet} sheet - 督促対象のシート
* @return {Object[]} - 処理を走らせるべき日の情報を含む配列
*/
function shouldTakeActionToday_(sheet) {
  // 督促タイミングの設定を取得
  const offsetInfo = prepareWithSetteiSheet_();
  if (offsetInfo.length === 0) {
    Logger.log("有効な督促タイミング設定が見つかりません");
    return []; // 空配列を返すのは、タイミング設定が無い場合のみ
  }

  // 統一シートから日付データを取得
  const tokusokubiColumnStart = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TOKUSOKUBI_COLUMN_START);
  const values = sheet.getRange(2, tokusokubiColumnStart, 1, sheet.getLastColumn() - tokusokubiColumnStart + 1).getValues();

  // 督促日付エリアの最初が日付型の値になってるかどうかで、列ズレを検知
  try {
    if (!values[0][0] || isNaN(new Date(values[0][0]).getTime())) {
      // エラーの場合は、エラーオブジェクトを投げる
      throw new Error("エラー: 開始列が日付データではありません。列がズレた可能性");
    }
  } catch (e) {
    // エラーの場合は、元のエラーオブジェクトのメッセージを再利用して新しいエラーオブジェクトを投げる
    throw new Error(e.message);
  }

  // 今日の日付を取得（時刻部分は0に設定）
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 有効な日付のみを抽出してソート
  const sortedDates = values[0]
    .map(value => {
      try {
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
      } catch (e) {
        return null;
      }
    })
    .filter(date => date !== null)
    .sort((a, b) => a - b);

  // 処理対象の日付をログ出力
  Logger.log("処理対象の日付一覧: " + sortedDates.map(date =>
    Utilities.formatDate(date, CONFIG_CLAIM_L.TIME_ZONE, CONFIG_CLAIM_L.DATE_FORMAT)
  ));

  // 処理を走らせるべき日を特定
  const tokusokuDates = [];

  // 各日付について、オフセット値と照合
  sortedDates.forEach(date => {
    const diff = calculateDateDiff_(today, date);

    // オフセット値と一致するものを探す
    offsetInfo.forEach(info => {
      if (diff === info.offset) {
        tokusokuDates.push({
          date: date,
          timing: info.name
        });

        // 処理を走らせるべき日として特定された日付をログ出力
        const formattedDate = Utilities.formatDate(date, CONFIG_CLAIM_L.TIME_ZONE, CONFIG_CLAIM_L.DATE_FORMAT);
        Logger.log(`${formattedDate} は ${info.name} の日です（今日から${diff}日）`);
      }
    });
  });
  return tokusokuDates;
}

/**
* ターゲットにすべき顧客を特定する関数
*
* @param {SpreadsheetApp.Sheet} sheet - 督促対象のシート
* @param {Object} tokusokuDate - 処理を走らせるべき日とタイミングの情報を含むオブジェクト
* @return {Object[]} recordSentData - 送信成功した顧客の行番号と送信日時の配列
*/
function findTargetCustomers_(sheet, tokusokuDate) {
  // 送信成功した顧客の行番号と送信日時を記録する配列
  let recordSentData = [];

  // いま関数が走ってる日が何列目にあるか取得
  const tokusokubiColumnStart = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TOKUSOKUBI_COLUMN_START);
  const dateColumnIndex = sheet.getRange(2, tokusokubiColumnStart, 1, sheet.getLastColumn() - tokusokubiColumnStart + 1)
    .getValues()[0]
    .map((date, index) => {
      try {
        const parsedDate = new Date(date);
        return isNaN(parsedDate.getTime()) ? null : parsedDate;
      } catch (e) {
        return null;
      }
    })
    .findIndex(date => date && date.getTime() === tokusokuDate.date.getTime())
    + tokusokubiColumnStart;

  // dateColumnIndex が見つからなかった場合の処理
  if (dateColumnIndex === tokusokubiColumnStart - 1) {
    Logger.log('エラー: 通知送信日が見つかりません。');
    return [];
  }

  // スプレッドシートの最終行を取得
  const lastRow = sheet.getLastRow();

  // 列名（アルファベット）から列番号を取得しておく
  const lineUserIdColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.LINE_USER_ID);
  const lastNameColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.LAST_NAME);
  const firstNameColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.FIRST_NAME);
  const customerPatternColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.CUSTOMER_PATTERN);

  // いま関数が走ってる対象日がある列を縦にガーッと走査
  for (let row = 3; row <= lastRow; row++) {
    const cell = sheet.getRange(row, dateColumnIndex);
    let cellValue = cell.getValue(); // あとで3桁カンマ区切りしなおして、それを新cellValueとするから、letにしとく

    // セルの値が数値型で、かつ NaN でなく、空欄でないことを確認
    if (typeof cellValue === 'number' && !isNaN(cellValue) && cellValue !== "") {
      // 着金済みの背景黄色セルの客はスキップ
      if (cell.getBackground() === '#ffff00') {
        Logger.log(`${row}行目は背景色が黄色のため、処理をスキップします。`);
        continue;
      }

      // 合致するセルがあれば、LINEユーザーID、姓、名を拾う
      const lineUserId = sheet.getRange(row, lineUserIdColumn).getValue();
      const lastName = sheet.getRange(row, lastNameColumn).getValue();
      const firstName = sheet.getRange(row, firstNameColumn).getValue();
      const customerPattern = sheet.getRange(row, customerPatternColumn).getValue();
      Logger.log(`行 ${row} の顧客パターン: ${customerPattern}`); 

      // LINEユーザーIDが空欄の場合はスキップ
      if (!lineUserId) {
        Logger.log(`${row}行目: ${lastName} ${firstName} は、LINEユーザーIDが空欄のためスキップされました。`);
        continue;
      }

      // 2行目の日付をそのまま原稿の材料にできるように曜日とかつける
      const shiharaiKigen = Utilities.formatDate(tokusokuDate.date, CONFIG_CLAIM_L.TIME_ZONE, CONFIG_CLAIM_L.DATE_FORMAT) +
        '(' + ['日', '月', '火', '水', '木', '金', '土'][tokusokuDate.date.getDay()] + ')';

      // 拾った各情報をログに出力
      Logger.log(`ターゲットの顧客を特定: 行: ${row}, 値: ${cellValue}, LINEユーザーID: ${lineUserId}, 姓: ${lastName}, 名: ${firstName}, 期限: ${shiharaiKigen}`);

      // 各顧客の金額を数値として扱ったので、プレースホルダー置換時に3桁区切りカンマ出るように、戻す
      cellValue = Utilities.formatString('%d', cellValue).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

      // 拾ったデータを変数に格納して各関数で使えるように
      const customerData = {
        lineUserId,
        kingaku: cellValue,
        timing: tokusokuDate.timing,
        lastName,
        firstName,
        kigen: shiharaiKigen,
        customerPattern
      };

      // createLineMessageに顧客データを渡し、送信結果を確認
      if (createLineMessage_(customerData)) {
        // 送信に成功したら、送信日時を記録
        recordSentData.push({
          row: row,
          sentDate: new Date()
        });
      }
    }
  }

  return recordSentData;
}


/**
* 設定シートから顧客パターンと日付、送信内容が入ってるか、を見て、今回必要なセルを見つける
*/
function locateTargetCells_(timing, customerPattern) {
  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_L.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_L.SETTEI_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  const colNums = {};
  for (const key in CONFIG_CLAIM_L.MESSAGE_COLUMNS) {
    colNums[key] = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.MESSAGE_COLUMNS[key]) - 1;
  }

  for (let i = 1; i < values.length; i++) {

    if (values[i][colNums.TIMING] === timing && values[i][colNums.PATTERN] === customerPattern) {
      // 該当する設定が見つかった場合、セルの位置情報と基本的な値を返す
      const row = i + 1;  // 実際のシートの行番号
      Logger.log(`該当の設定 - タイミング: ${timing}, パターン: ${customerPattern}, 行番号: ${row}`);

      // content1は必須項目なのでチェック
      if (!values[i][colNums.CONTENT_1]) {
        Logger.log(`エラー: ${timing}, パターン${customerPattern}のcontent1が空です`);
        throw new Error(`${timing}, パターン${customerPattern}のcontent1が空です`);
      }

      return {
        // セルの位置情報を返す
        content1Cell: sheet.getRange(row, colNums.CONTENT_1 + 1),
        content2Cell: values[i][colNums.CONTENT_2] ? sheet.getRange(row, colNums.CONTENT_2 + 1) : null,
        altText1: values[i][colNums.ALT_TEXT_1],
        altText2: values[i][colNums.ALT_TEXT_2],
        imageHeight1: values[i][colNums.IMAGE_HEIGHT_1],
        imageHeight2: values[i][colNums.IMAGE_HEIGHT_2]
      };
    }
  }

  // 該当する設定が見つからなかった場合
  Logger.log(`エラー: ${timing}、パターン${customerPattern}に対応するメッセージが見つかりません。`);
  throw new Error(`${timing}、パターン${customerPattern}に対応するメッセージが見つかりません。`);
}

// --- ここからメッセージ送信関連 ---

/**
* findTarget~~からもらったデータを実際にメッセ作る関数に投げて
* 戻ってきたら、実送信関数に投げる、中継ぎ兼 現場職の司令塔的な関数
*/
function createLineMessage_(customerData) {
  try {
    // メッセージ設定のセル位置をlocateTargetCells_から取得しとく
    const targetCells = locateTargetCells_(customerData.timing, customerData.customerPattern);

    // findTargetからまとめてもらったcustomerDataを解凍   
    const placeholders = {
      userId: customerData.lineUserId,
      kingaku: customerData.kingaku,
      sei: customerData.lastName,
      mei: customerData.firstName,
      kigen: customerData.kigen
    };

    const accessToken = getLineAccessToken_();

    // Content1の処理 とにかくcreateMessageContentに投げて、画像判定とかもそっちでやらせる。
    // ここが空欄だとエラー（1通目は絶対存在する前提）
    let message1 = imageOrDocsAndPrepTextMessage_(
      targetCells.content1Cell,
      placeholders,
      targetCells.altText1,
      targetCells.imageHeight1
    );
    if (!message1) return false;

    let sendResult1 = sendMessageByTypeForLine_(
      customerData.lineUserId,
      message1,
      accessToken
    );
    if (!sendResult1) return false;


    // Content2の処理（存在する場合のみ）
    if (targetCells.content2Cell) {
      let message2 = imageOrDocsAndPrepTextMessage_(
        targetCells.content2Cell,
        placeholders,
        targetCells.altText2,
        targetCells.imageHeight2
      );
      if (!message2) return false;

      let sendResult2 = sendMessageByTypeForLine_(
        customerData.lineUserId,
        message2,
        accessToken
      );
      if (!sendResult2) return false;
    }

    return true;

  } catch (e) {
    Logger.log(`LINE通知メッセージ作成に失敗しました。: ${e.message}`);
    return false;
  }
}

/**
 * セルの内容を見て、画像・document判定をやって、テキスト系ならここで下ごしらえを完了させる
 * 画像の下ごしらえは後続関数に分離
 */
function imageOrDocsAndPrepTextMessage_(cell, placeholders, altText, manualImageHeight) {
    const richTextValue = cell.getRichTextValue();
    let content = cell.getValue().toString().trim();

    try {
        let targetUrl = '';
        let displayText = '';

        // リッチテキストからURLを取得、なければ通常の値を使用
        if (richTextValue && richTextValue.getRuns().some(run => run.getLinkUrl())) {
            const runs = richTextValue.getRuns();
            const linkRun = runs.find(run => run.getLinkUrl());
            targetUrl = linkRun.getLinkUrl();
            displayText = linkRun.getText();
            content = targetUrl; // 後続の処理のため contentを更新
        } else {
            targetUrl = content;
        }

        // Googleドライブからのリンクか判定
        if (targetUrl.includes('google.com/')) {
            let fileId;
            try {
                fileId = GeneralUtils.getFileIdFromDriveUrl_(targetUrl); // ヘルパー関数でID取得
            } catch (e) {
                Logger.log(`GoogleドライブURL形式エラー: ${e.message}`);
                return {
                    type: 'text',
                    text: displayText ? 
                        `エラー：GoogleドライブのURLからファイルIDを取得できませんでした リンクテキスト：${displayText}` :
                        `エラー：GoogleドライブのURLからファイルIDを取得できませんでした：${targetUrl}`
                };
            }
            
            const file = DriveApp.getFileById(fileId);
            const mimeType = file.getMimeType();

            // 画像かdocsか判定
            if (mimeType.includes('image')) {
                Logger.log(`MIMEで画像と判定`);
                // 画像と判定された場合はこの関数離脱
                return createImageMessage_(targetUrl, altText, placeholders, manualImageHeight);
            } else if (mimeType === 'application/vnd.google-apps.document') {

                // Googleドキュメントと判定された場合はこの関数の後続に飛ばす
                const doc = DocumentApp.openById(fileId);
                return {
                    type: 'text',
                    text: replacePlaceholdersForLINEClaim_(doc.getBody().getText(), placeholders)
                };
                
            } else {  // ドキュメントでも画像でもない場合
                Logger.log(`非対応ファイル形式: ${file.getName()}, MIMEタイプ: ${mimeType}`);
                return {
                    type: 'text',
                    text: displayText ? 
                        `エラー：ドキュメントでも画像でもないファイルです リンクテキスト：${displayText} ファイル形式：${mimeType}` :
                        `エラー：ドキュメントでも画像でもないファイルです\n${targetUrl}\nファイル形式：${mimeType}`
                };
           }
       }

        // driveじゃないアプロダ系画像判定
        if (targetUrl.match(/\.(jpe?g|png|gif)$/i) ||
            targetUrl.match(/\/(\d+)[_:](\d+)[_.](jpe?g|png|gif)$/i)) {
            //画像ならここで、この関数からは離脱
            return createImageMessage_(targetUrl, altText, placeholders, manualImageHeight);
        }

        // 通常のテキストメッセージ
        return {
            type: 'text',
            text: replacePlaceholdersForLINEClaim_(displayText || targetUrl, placeholders)
        };

    } catch (e) {
        Logger.log(`メッセージ作成エラー: ${e.message}`);
        return {
            type: 'text',
            text: replacePlaceholdersForLINEClaim_(cell.getDisplayValue(), placeholders)
        };
    }
}


/**
* メッセージの実送信処理！（LINE）
* @param {string} userId - LINE ユーザーID
* @param {Object} messageObj - 送信するメッセージオブジェクト
* @param {string} accessToken - LINE チャネルアクセストークン
* @return {boolean} 送信成功時はtrue、失敗時はfalse
*/
function sendMessageByTypeForLine_(userId, messageObj, accessToken) {
  try {
    if (CONFIG_CLAIM_L.DEBUG_MODE) {
      Logger.log(JSON.stringify(messageObj, null, 2));
      Logger.log('デバッグモードのため、メッセージは送信されません。');
      return true;
    }

    const options = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': `Bearer ${accessToken}`
      },
      'payload': JSON.stringify({
        'to': userId,
        'messages': [messageObj]
      }),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(CONFIG_CLAIM_L.LINE.API_ENDPOINT, options);
    
    if (response.getResponseCode() === 200) {
      Logger.log(`メッセージ送信成功: ${userId}`);
      return true;
    } else {
      Logger.log(`メッセージ送信エラー: ${userId}, HTTPステータスコード: ${response.getResponseCode()}, エラーメッセージ: ${response.getContentText()}`);
      return false;
    }

  } catch (e) {
    Logger.log(`送信処理中にエラーが発生: ${e.message}`);
    return false;
  }
}

// --- ここからヘルパー関数 ---

/**
 * 画像メッセージ作成の場合は、ﾘｯﾁﾒｯｾのアス比処理が複雑なので分離させた
 */
function createImageMessage_(content, altText, placeholders, manualImageHeight) {

  // 30日以内に過去送ったことある画像なら、ｷｬｯｼｭ使って処理時短
  const CACHE_DURATION = 30 * 24 * 60 * 60;

  let imageUrl = content;
  let aspectRatio = "1:1";

  try {
    // アプロダでURLにサイズ盛り込んでるか確認
    const bitlySizeMatch = content.match(/\/(\d+)_(\d+)_(jpe?g|png)$/i);
    const normalSizeMatch = content.match(/\/(\d+)[_:](\d+)[_.](jpe?g|png)$/i);

    // いずれかのパターンからサイズを取得
    if (bitlySizeMatch || normalSizeMatch) {
      const match = bitlySizeMatch || normalSizeMatch;
      const width = parseInt(match[1]);
      const height = parseInt(match[2]);

      // アスペクト比を計算する内部関数
      function calculateAspectRatio_(width, height) {
        const ratio = height / width;
        // LINEの制限に合わせて3:1を超えないように調整
        return ratio > 3 ? `${width}:${width * 3}` : `${width}:${height}`;
      }

      aspectRatio = calculateAspectRatio_(width, height);
      Logger.log(`URLから取得した画像サイズ: ${width}x${height}, アスペクト比: ${aspectRatio}`);
    }

    // ---Googleドライブの画像URLが直接貼られてるかチェック---
    if (content.includes('drive.google.com') || content.includes('docs.google.com/file')) {
      const fileId = GeneralUtils.getFileIdFromDriveUrl_(content);

      // キャッシュあるか確認
      const cache = CacheService.getScriptCache();
      const urlCacheKey = `image_url_${fileId}`;
      const dimensionsCacheKey = `image_dimensions_${fileId}`;
      let cachedUrl = cache.get(urlCacheKey);
      let cachedDimensions = cache.get(dimensionsCacheKey);

      if (cachedUrl && cachedDimensions) {
        imageUrl = cachedUrl;
        const dimensions = JSON.parse(cachedDimensions);
        aspectRatio = `${dimensions.width}:${dimensions.height}`;
        Logger.log(`キャッシュヒット！ fileId: ${fileId}`);

      } else {
        Logger.log(`キャッシュまだ無い。fileId: ${fileId} の画像を新規処理します`);
        const file = DriveApp.getFileById(fileId);
        file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
        
        // 公開URLの生成
        imageUrl = `https://drive.google.com/uc?id=${fileId}&export=download`;

        // ユーティリティ関数でサイズ取得を試みる
        let dimensionsObtained = false;
        try {
          const dimensions = GeneralUtils.getImageDimensions_(content);
          Logger.log(`Drive画像サイズ取得成功 - 幅: ${dimensions.width}px, 高さ: ${dimensions.height}px`);
          aspectRatio = `${dimensions.width}:${dimensions.height}`;
          dimensionsObtained = true;

          // サイズ情報もキャッシュに保存
          cache.put(dimensionsCacheKey, JSON.stringify(dimensions), CACHE_DURATION);
          Logger.log(`画像サイズをキャッシュに保存: ${JSON.stringify(dimensions)}`);
        } catch (dimensionError) {
          Logger.log(`Drive画像サイズ取得失敗: ${dimensionError.message}`);
        }

        // サイズ取得に失敗した場合のみ、IMAGE_HEIGHT列の値をの処理を使用
        if (!dimensionsObtained) {
          const customHeight = (manualImageHeight === null || manualImageHeight === undefined || isNaN(manualImageHeight)) ? 1040 : manualImageHeight;
          aspectRatio = (manualImageHeight === null || manualImageHeight === undefined || isNaN(manualImageHeight)) ? '1:1' : `1040:${customHeight}`;
          Logger.log(`サイズ取得失敗時の処理 - アスペクト比: ${aspectRatio}, カスタム高さ: ${customHeight}`);
        }

        cache.put(urlCacheKey, imageUrl, CACHE_DURATION);
        Logger.log(`新規URL生成。30日間キャッシュに保存: ${imageUrl}`);
      }
    }

    // Flexメッセージの作成
    const messageObj = {
      type: 'flex',
      altText: altText ? replacePlaceholdersForLINEClaim_(altText, placeholders) : '画像メッセージ', 
      contents: {
        type: 'bubble',
        size: 'giga',
        hero: {
          type: 'image',
          url: imageUrl,
          size: 'full',
          aspectMode: 'cover',
          aspectRatio: aspectRatio,
        }
      }
    };

    return messageObj;

  } catch (e) {
    Logger.log(`画像メッセージの作成中にエラーが発生: ${e.message}`);
    throw e; // エラーを上位で捕捉できるように再スロー
  }
}


/**
* 下準備として、関数同士の受け渡しに使う督促名称とか、今日との差分日付として重要になってくる「オフセット値」とかを、先に拾っておく関数
*
* @returns {Object[]} - 以下の形式のオブジェクトの配列
* {
*   name: string,      // 督促名称（例：「3日前」「1週間後」など）
*   offset: number,    // 日付オフセット値（正：未来、負：過去）
* }
*/
function prepareWithSetteiSheet_() {
  // B列の日付オフセットが顧客パターンが異なっても同じ場合が多いので
  // 縦に列走査したとき、被ってる日付を重複して取得せず、片方顧客パターンのみ取得するために
  // 既に取得した日付オフセット値を関数に意識させる
  const addedOffsets = new Set();

  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_L.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_L.SETTEI_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  // 督促送信内容シートの列設定から、A,B列を取得
  const timingColumnNumber = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.MESSAGE_COLUMNS.TIMING);
  const offsetColumnNumber = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.MESSAGE_COLUMNS.OFFSET);

  // オフセット情報を格納する配列
  const offsetInfo = [];

  // ヘッダー行をスキップして処理（i=1から開始）
  for (let i = 1; i < values.length; i++) {
    const timingName = values[i][timingColumnNumber - 1];  // 督促名称
    const offsetDays = values[i][offsetColumnNumber - 1];  // 日付オフセット（プラスが未来、マイナスが過去）

    // 空行または不正な値の行はスキップ
    if (!timingName || offsetDays === undefined || offsetDays === "") continue;
    if (typeof offsetDays !== 'number') continue;  // 数値以外はスキップ

    // このオフセット値がまだ追加されていない場合のみ追加
    if (!addedOffsets.has(offsetDays)) {
      addedOffsets.add(offsetDays);
      offsetInfo.push({
        name: timingName,
        offset: offsetDays,
      });
    }
  }
  return offsetInfo;
}

/**
* 2つの日付間の日数差を計算する補助関数
* 時刻の部分は無視して純粋な日付の差分のみを計算する
*
* @param {Date} date1 - 基準日
* @param {Date} date2 - 比較対象日
* @returns {number} 日数差（date2 - date1）
*/
function calculateDateDiff_(date1, date2) {
  // 時刻部分を除去して純粋な日付として扱う
  const d1 = new Date(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const d2 = new Date(date2.getFullYear(), date2.getMonth(), date2.getDate());
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}


/**
* プレースホルダーの置換処理
*/
function replacePlaceholdersForLINEClaim_(text, placeholders) {

  let replacedText = text;

  // {sei} のチェックと置換
  if (text.includes('{sei}')) {
    if (placeholders.sei === undefined || placeholders.sei === null || placeholders.sei === '') {
      throw new Error("エラー: {sei} が原稿内にありますが、値が空欄または、変な引数・変数入ってます。");
    }
    replacedText = replacedText.replace('{sei}', placeholders.sei);
  }

  // {kigen} のチェックと置換
  if (text.includes('{kigen}')) {
    if (placeholders.kigen === undefined || placeholders.kigen === null || placeholders.kigen === '') {
      throw new Error("エラー: {kigen} が原稿内にありますが、値が空欄または、変な引数・変数入ってます。");
    }
    replacedText = replacedText.replace('{kigen}', placeholders.kigen);
  }

  // {kingaku} のチェックと置換
  if (text.includes('{kingaku}')) {
    if (placeholders.kingaku === undefined || placeholders.kingaku === null || placeholders.kingaku === '') {
      throw new Error("エラー: {kingaku} が原稿内にありますが、値が空欄または、変な引数・変数入ってます。");
    }
    replacedText = replacedText.replace('{kingaku}', placeholders.kingaku);
  }

  // {mei} は空欄でもいい
  replacedText = replacedText.replace('{mei}', placeholders.mei || ''); // 値がない場合は空文字列に置換

  return replacedText;
}


/**
* 送った履歴をスプシに入れる関数
*/
function recordSentDate(recordSentData) {

  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_L.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_L.SHEET_NAME);
  const sentDateColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.SENTDATE_OF_L);

  try {
    // 処理対象の行を含む範囲全体を一度に取得
    const minRow = Math.min(...recordSentData.map(data => data.row));
    const maxRow = Math.max(...recordSentData.map(data => data.row));
    
    // 既存の値を一括で取得
    const range = sheet.getRange(minRow, sentDateColumn, maxRow - minRow + 1, 1);
    const existingValues = range.getValues();
    
    // 更新用の2次元配列を作成
    const updatedValues = existingValues.map((row, index) => {
      const currentRow = minRow + index;
      const recordData = recordSentData.find(data => data.row === currentRow);
      
      if (!recordData) {
        return row; // 更新対象でない行は既存の値をそのまま返す
      }

      const formattedDate = Utilities.formatDate(
        new Date(recordData.sentDate),
        CONFIG_CLAIM_L.TIME_ZONE,
        'MM/dd'
      );

      const existingValue = row[0];
      const newValue = existingValue
        ? String(existingValue) + ', ' + formattedDate 
        : formattedDate;

      return [newValue];
    });

    // 一括で書き込み
    range.setNumberFormat('@');
    range.setValues(updatedValues);
    
    //ログ準備
    const targetNames = recordSentData.map(data => {
    const lastName = sheet.getRange(data.row, GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.LAST_NAME)).getValue();
    const firstName = sheet.getRange(data.row, GeneralUtils.columnToNumber_(CONFIG_CLAIM_L.TARGET_COLUMNS.FIRST_NAME)).getValue();
    return `${lastName}${firstName}`;
    }).join('、');

    Logger.log(`${recordSentData.length}件（${targetNames}）の送信完了日を一括記録しました。`);

  } catch (e) {
    Logger.log(`一括書き込みに失敗したため、1行ずつの処理にフォールバックします: ${e.message}`);
    
    recordSentData.forEach(data => {
      try {
        const range = sheet.getRange(data.row, sentDateColumn);
        const existingValue = range.getValue();
        const formattedDate = Utilities.formatDate(
          new Date(data.sentDate),
          CONFIG_CLAIM_L.TIME_ZONE,
          'MM/dd'
        );
        
        const newValue = existingValue
          ? String(existingValue) + ', ' + formattedDate
          : formattedDate;

        range.setNumberFormat('@').setValue(newValue);
        Logger.log(`${data.row}行目のメール送信完了日に ${formattedDate} を個別に追加記録しました。`);
      } catch (e) {
        Logger.log(`エラー発生: ${data.row}行目の書き込みに失敗しました。エラーメッセージ: ${e.message}`);
      }
    });
  }
}

  // パブリックAPIとして公開する関数
  return {
    lineClaimMain: lineClaimMain
  };

}) //即時実行関数の終わり

(); //即時実行関数を、実行するためのカッコ


/**
* トリガーのためのトリガーから叩かれる関数
* 名前空間より外に置いとく必要ある
*/
function setLineClaimTrigger() {

  // 既存のトリガーをすべて削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "lineClaimMain") { //この名前のトリガーを探して消す
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const time = new Date();
  
  time.setHours(7);
  time.setMinutes(12);
  
  // 指定した時分が現在より過去の場合は、翌日のその時刻に設定
  if (time < new Date()) {
    time.setDate(time.getDate() + 1);
  }
  
  ScriptApp.newTrigger('lineClaimMain').timeBased().at(time).create();
}

