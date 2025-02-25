/**
 * =====================================================
 * メールで、自動で、振込日リマインドや滞納者督促を行う機能
 * =====================================================
 */


/**
* 最後の関数で設置されたトリガーからメイン関数を叩けるように
* 名前空間より外に出す
*/
function mailClaimMain() {
  MailClaimSpace.mailClaimMain();
}


/**
* 名前空間開始
*/
const MailClaimSpace = (function() {

/**
* このファイル全体の設定
*/
const CONFIG_CLAIM_Ma = {
  // 基本設定
  DEBUG_MODE: false,                 // true: デバッグモード, false: 通常モード（本番！モード）実装できてない！
  TIME_ZONE: 'Asia/Tokyo',         // タイムゾーン設定
  DATE_FORMAT: 'yyyy/MM/dd',       // 日付の表示形式

  // Gmail関連の設定
  GMAIL: {
    SENDER_NAME: '', // 送信者名
    SENDER_EMAIL_ADDRESS: '' // 送信者メールアドレス
  },

  // スプレッドシート関連の設定
  SPREADSHEET_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1CXXXX',  // 操作対象のスプレッドシートID
  SHEET_NAME: '統一シート',                                          // メインの操作対象シート名
  SETTEI_SHEET_NAME: '督促送信内容', // 各種設定があるシート名

  PATTERNS: {
  PATTERN1: '松尾',
  PATTERN2: '中村'
  },

  // 統一シートの列設定
  TARGET_COLUMNS: {
    MAIL: 'D',   // メアドが記録されている列
    LAST_NAME: 'F',      // 姓が記録されている列
    FIRST_NAME: 'G',     // 名が記録されている列
    SENTDATE_OF_M: 'ES', // メール送信日を記録する列
    CUSTOMER_PATTERN: 'EP' //顧客パターン列（今は松尾か中村か） 
  },
  TOKUSOKUBI_COLUMN_START: 'EV',  // 督促対象日のデータが開始される列

  // 督促送信内容シートの列設定
  MESSAGE_COLUMNS: {
    TIMING: 'A',        // 督促タイミングの名前が入ってる（この名前の文字列をプログラムが認識して、各関数に受け渡す）
    OFFSET:'B',        // 統一シートに入力されてる日付と今日との差分（合致すればアクションすべき日）
    KENMEI: 'G',     // メール件名（空欄の場合は本文1行目を件名にする）
    HONBUN: 'H',     // メール本文
    GAZOU_1: 'I',     // 添付画像1のURL（空欄の場合もある）
    GAZOU_1_NAME: 'J',     // 添付画像1の題名（空欄の場合もある）
    GAZOU_2: 'K',     // 添付画像2のURL（空欄の場合もある）
    GAZOU_2_NAME: 'L',     // 添付画像2の題名（空欄の場合もある）
    PATTERN: 'O' , //松尾か中村か
  }
};



/**
* 督促処理のメイン関数(Mail)
*/
function mailClaimMain() {
  // スプレッドシートとシートを取得
  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_Ma.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_Ma.SHEET_NAME);

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

  // 処理を走らせる日であれば、それぞれの督促日（の情報）について、findTargetCustomers_関数で、ターゲットにすべき顧客を特定して、メールメッセージを作成・送信
  if (tokusokuDates.length > 0) {
    tokusokuDates.forEach(tokusokuDate => {
      //修正開始
      let recordSentData = findTargetCustomers_(sheet, tokusokuDate); 
      //修正終了
      if(recordSentData.length > 0){

        //履歴入れる関数に、送信したよって情報渡す
        recordSentDate(recordSentData);
      }
    });
  } else {
    // 今日が処理すべき日でない場合のログ出力（エラーの場合はここは実行されない）
    const today = new Date();
    const kyouFormatted = Utilities.formatDate(today, CONFIG_CLAIM_Ma.TIME_ZONE, CONFIG_CLAIM_Ma.DATE_FORMAT) + '(' + ['日', '月', '火', '水', '木', '金', '土'][today.getDay()] + ')';
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
  const tokusokubiColumnStart = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TOKUSOKUBI_COLUMN_START);
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
    Utilities.formatDate(date, CONFIG_CLAIM_Ma.TIME_ZONE, CONFIG_CLAIM_Ma.DATE_FORMAT)
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
        const formattedDate = Utilities.formatDate(date, CONFIG_CLAIM_Ma.TIME_ZONE, CONFIG_CLAIM_Ma.DATE_FORMAT);
        Logger.log(`${formattedDate} は ${info.name} の日です（今日から${diff}日）`);
      }
    });
  });
  return tokusokuDates;
}

/**
* ターゲットにすべき顧客を特定する関数(メール版)
*
* @param {SpreadsheetApp.Sheet} sheet - 督促対象のシート
* @param {Object} tokusokuDate - 処理を走らせるべき日とタイミングの情報を含むオブジェクト
* @return {Object[]} recordSentData - 送信成功した顧客の行番号と送信日時の配列
*/
function findTargetCustomers_(sheet, tokusokuDate) {

  // 送信成功した顧客の行番号と送信日時を記録する配列
  let recordSentData = [];

  // いま関数が走ってる日が何列目にあるか取得
  const tokusokubiColumnStart = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TOKUSOKUBI_COLUMN_START);
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
  const mailAddressColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.MAIL);
  const lastNameColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.LAST_NAME);
  const firstNameColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.FIRST_NAME);
  const customerPatternColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.CUSTOMER_PATTERN); // 顧客パターンを取得

  // いま関数が走ってる対象日がある列を縦にガーッと走査
  for (let row = 3; row <= lastRow; row++) {
    const cell = sheet.getRange(row, dateColumnIndex);
    const cellValue = cell.getValue();

    // セルの値が数値型で、かつ NaN でなく、空欄でないことを確認
    if (typeof cellValue === 'number' && !isNaN(cellValue) && cellValue !== "") {

      //着金済みの背景黄色セルの客はスキップ
      if (cell.getBackground() === '#ffff00') {
        Logger.log(`${row}行目は背景色が黄色のため、処理をスキップします。`);
        continue;
      }

      // 合致するセルがあれば、メールアドレス、姓、名、顧客パターンを拾って督促文の材料にする
      const mailAddress = sheet.getRange(row, mailAddressColumn).getValue();
      const lastName = sheet.getRange(row, lastNameColumn).getValue();
      const firstName = sheet.getRange(row, firstNameColumn).getValue();
      const customerPattern = sheet.getRange(row, customerPatternColumn).getValue();

      // メールアドレスが空欄の場合はスキップ
      if (!mailAddress) {
        Logger.log(`${row}行目: ${lastName} ${firstName} は、メールアドレスが空欄のためスキップされました。`);
        continue;
      }

      // 2行目の日付をそのまま原稿の材料にできるように曜日とかつける
      const furikomiKiGen = Utilities.formatDate(tokusokuDate.date, CONFIG_CLAIM_Ma.TIME_ZONE, CONFIG_CLAIM_Ma.DATE_FORMAT) +
        '(' + ['日', '月', '火', '水', '木', '金', '土'][tokusokuDate.date.getDay()] + ')';

      // 拾った各情報をログに出力
      Logger.log(`ターゲットの顧客を特定: 行: ${row}, 値: ${cellValue}, メールアドレス: ${mailAddress}, 姓: ${lastName}, 名: ${firstName}, 期限: ${furikomiKiGen}, パターン: ${customerPattern}`);

      //createMailMessage_関数にまとめて渡す。向こうでのまとまった引数名は
      let sendResult = createMailMessage_({
        row: row, 
        mailAddress: mailAddress,
        cellValue: cellValue,
        timing: tokusokuDate.timing,
        lastName: lastName,
        firstName: firstName,
        furikomiKiGen: furikomiKiGen,
        customerPattern: customerPattern 
      });

      // 送信に成功したら、送信日時を記録
      if (sendResult){
        recordSentData.push({
          row: row,
          sentDate: new Date()
        });
      }
    }
  }
  return recordSentData;
}


// --- ここからメッセージ送信関連 ---

/**
* メールメッセージの下ごしらえを行う関数
*
* @param {Object} dataByFindTarget - 顧客データを含むオブジェクト
* @return {boolean} メッセージ送信が成功したかどうか
*/
function createMailMessage_(dataByFindTarget) {

  try {
    // メッセージ設定のセル位置をlocateTargetCells_から取得しとく
    const targetCells = locateTargetCells_(dataByFindTarget.timing, dataByFindTarget.customerPattern);

    let emailSubjectCell = targetCells.kenmeiCell.getValue();
    let emailBodyCell = targetCells.honbunCell;  // Rangeオブジェクトをそのまま保持
    let emailBodyValue = emailBodyCell.getValue(); 

    // 本文がGoogleドキュメントのURLなら、本文を取得
    if (emailBodyValue && emailBodyValue.toString().startsWith('https://docs.google.com/document/')) {
      emailBodyCell = getDocumentContent_(emailBodyCell);

      //件名が空欄 かつ 本文が取得できている場合、本文の1行目を件名にし、本文からは削除
      if (!emailSubjectCell && emailBodyCell) {
        let lines = emailBodyCell.split('\n');
        emailSubjectCell = lines[0];
        emailBodyCell = lines.slice(1).join('\n');
      }

    } else {
       const richTextValue = emailBodyCell.getRichTextValue();
       const runs = richTextValue.getRuns();
        for (let run of runs) {
          const linkUrl = run.getLinkUrl();
          if (linkUrl && linkUrl.includes('google.com/')) {
            let fileId;
            try {
              fileId = GeneralUtils.getFileIdFromDriveUrl_(linkUrl);
               Logger.log(`ファイルID抽出成功: ${fileId}`);
            } catch (e) {
              Logger.log(`GoogleドライブURL形式エラー: ${e.message}`);
              // エラー時は処理を停止
              throw new Error(`GoogleドライブのURLからファイルIDを取得できませんでした: ${linkUrl}`);
            }

            const mimeType = GeneralUtils.getMimeType_(fileId);

            if (mimeType === 'application/vnd.google-apps.document') {
              Logger.log(`リッチテキストでMIMEでdocs判定`);
              docUrl = linkUrl;

              // URLからドキュメントIDを抽出
              const docIdMatch = docUrl.match(/\/document\/d\/([-\w]{25,})/);
              if (!docIdMatch) {
                throw new Error('有効なドキュメントIDが見つかりません');
              }
              const docId = docIdMatch[1];
              
              // ドキュメントから本文を取得
              const docrich = DocumentApp.openById(docId);
              const docTextFromRich = docrich.getBody().getText();
              Logger.log(`ドキュメントから本文取得成功: ${docTextFromRich.substring(0, 50)}...`);
              
              emailBodyCell = docTextFromRich;

              //件名が空欄 かつ 本文が取得できている場合、本文の1行目を件名にし、本文からは削除
              if (!emailSubjectCell && emailBodyCell) {
                let lines = emailBodyCell.split('\n');
                emailSubjectCell = lines[0];
                emailBodyCell = lines.slice(1).join('\n');
              }
            
            } else {
              Logger.log(`非対応ファイル形式: ${file.getName()}, MIMEタイプ: ${mimeType}`);
              return {
              type: 'text',
              text: `エラー：ドキュメントじゃないリンクが置かれてるかも\n${content}\nファイル形式：${mimeType}`
              }
            }
          }
        }
      }

    // メール本文に埋め込む値を定義
    const placeholders = {
      mail: dataByFindTarget.mailAddress,
      kingaku: dataByFindTarget.cellValue,
      sei: dataByFindTarget.lastName,
      mei: dataByFindTarget.firstName,
      kigen: dataByFindTarget.furikomiKiGen
    };

    // プレースホルダーを実際の値に置換
    emailBodyCell = replaceAllPlaceholders_(emailBodyCell, placeholders);
    emailSubjectCell = replaceAllPlaceholders_(emailSubjectCell, placeholders);

    // HTML形式の本文を作成
    const htmlBody = emailBodyCell.replace(/\n/g, '<br>');

    //添付画像の取得処理
    const attachments = [];
    
    // 画像1の処理
    const image1Blob = getImageBlob_(targetCells.gazou_1Cell, targetCells.gazou_1_nameCell);
    if (image1Blob) {
      attachments.push(image1Blob);
    } else {
      Logger.log("画像1はない");
    }

    // 画像2の処理
    const image2Blob = getImageBlob_(targetCells.gazou_2Cell, targetCells.gazou_2_nameCell);
    if (image2Blob) {
      attachments.push(image2Blob);
    } else {
      Logger.log("画像2はない");
    }

    // メール送信処理
    if (sendMailMessage_(dataByFindTarget.mailAddress, emailSubjectCell, emailBodyCell, htmlBody, attachments)) {
      Logger.log(`${dataByFindTarget.mailAddress} への ${dataByFindTarget.timing} の督促メール作成完了。`);
      return true;
    } else {
      Logger.log(`${dataByFindTarget.mailAddress} への ${dataByFindTarget.timing} の督促メール作成失敗。`);
      return false;
    }
  } catch (e) {
    Logger.log(`メールメッセージ作成中にエラーが発生しました。: ${e.message}`);
    return false;
  }
}

/**
 * 画像のリンクについて、原稿シートに画像URL直打ちでも、ファイル名リンクにしても読み取ってくれる関数
 * @param {Range} imageCell - 画像URLまたはリンクが含まれるセル
 * @param {Range} nameCell - 画像のファイル名が含まれるセル
 * @return {Blob|null} - 取得した画像のBlobデータ、取得できない場合はnull
 */
function getImageBlob_(imageCell, nameCell) {
  if (!imageCell) {
    return null;
  }

  try {
    let blob;
    const url = imageCell.getValue().trim();

    let richTextValue = imageCell.getRichTextValue();
    if (richTextValue) {
      Logger.log("リッチテキストを検出しました。URL: " + url); // デバッグログ追加
      const runs = richTextValue.getRuns();
      for (let run of runs) {
        const linkUrl = run.getLinkUrl();
        if (linkUrl && linkUrl.includes('google.com/')) {
          try {
            const fileId = GeneralUtils.getFileIdFromDriveUrl_(linkUrl);
            Logger.log("抽出したファイルID: " + fileId); // デバッグログ追加
            
            try {
              const file = DriveApp.getFileById(fileId);
              const mimeType = file.getMimeType();
              Logger.log(`MIMEタイプ: ${mimeType}`);

              if (mimeType.includes('image')) {
                try {
                  blob = file.getBlob();
                } catch (blobError) {
                  Logger.log(`Blobの取得に失敗: ${blobError.message}`);
                  continue;
                }
                
                // nameCellが空欄の場合、リッチテキストからファイル名を取得
                if (!nameCell || !nameCell.getValue()) {
                  Logger.log("nameCellが空欄です。リッチテキストからファイル名を取得します。");
                  const richTextFileName = run.getText();
                  if(richTextFileName){
                    blob.setName(richTextFileName);
                    Logger.log(`ファイル名を ${richTextFileName} に設定しました。`);
                  } else {
                    blob.setName("名称不明.jpeg");
                    Logger.log(`ファイル名を 名称不明.jpeg に設定しました。`);
                  }
                }
                break;
              }
            } catch (fileError) {
              Logger.log(`ファイルへのアクセスエラー: ${fileError.message}`);
              Logger.log("このファイルに対する適切なアクセス権限があることを確認してください。");
              continue;
            }
          } catch (urlError) {
            Logger.log(`画像のGoogleドライブURL処理エラー: ${urlError.message}`);
            continue;
          }
        }
      }
    }

    // リッチテキストで画像が見つからなかった場合のフォールバック
    if (!blob) {
      Logger.log("リッチテキストで画像が見つからなかったため、従来のURL処理を試みます。");
      if (url.includes('drive.google.com') || url.includes('docs.google.com/file')) {
        // GoogleドライブのファイルIDを抽出
        let fileId;
        if (url.includes('/file/d/')) {
          fileId = url.match(/\/file\/d\/([^\/]*)/)[1];
        } else if (url.includes('id=')) {
          fileId = url.match(/id=([^&]*)/)[1];
        }
        Logger.log(`従来のURL処理で抽出されたファイルID: ${fileId}`);

        if (fileId) {
          const file = DriveApp.getFileById(fileId);
          blob = file.getBlob();
          Logger.log("従来のURL処理で画像を取得しました。");
           // nameCellが空欄の場合、URLからファイル名を取得
          if (!nameCell || !nameCell.getValue()) {
            const urlFileName = url.split('/').pop();
            blob.setName(urlFileName);
            Logger.log(`ファイル名を ${urlFileName} に設定しました。`);
          }
        }
      } else if (url.startsWith('http')) {
        const response = UrlFetchApp.fetch(url);
        blob = response.getBlob();
        Logger.log("HTTP URLから画像を取得しました。");
         // nameCellが空欄の場合、URLからファイル名を取得
        if (!nameCell || !nameCell.getValue()) {
          const urlFileName = url.split('/').pop();
          blob.setName(urlFileName);
          Logger.log(`ファイル名を ${urlFileName} に設定しました。`);
        }
      }
    }
    if (blob) {
      // nameCell が存在する場合は、その値で上書き
      if (nameCell && nameCell.getValue()) {
        blob.setName(nameCell.getValue());
        Logger.log(`画像名: ${nameCell.getValue()} を設定しました。`);
      }
      return blob;
    }
  } catch (e) {
    Logger.log(`画像の取得に失敗しました: ${e.message}`);
  }
  Logger.log("画像の取得に失敗しました。null を返します。");
  return null;
}


/**
* メール送信処理（メールの実送信を行う）
*
* @param {string} mailAddress - メールアドレス
* @param {string} subject - メールの件名
* @param {string} body - メールの本文（プレーンテキスト）
* @param {string} htmlBody - メールの本文（HTML形式）
* @param {Blob[]} attachments - 添付ファイルの配列
* @return {boolean} メッセージ送信が成功したかどうか
*/
function sendMailMessage_(mailAddress, subject, body, htmlBody, attachments) {

  try {
    if (CONFIG_CLAIM_Ma.DEBUG_MODE) {
      Logger.log(`デバッグモード: ${mailAddress} へのメール送信をシミュレート`);
      Logger.log(`送信メール情報: 宛先: ${mailAddress}, 件名: ${subject}`); 
      Logger.log(`  本文: ${body}`);

      // 添付ファイルの情報をログ出力
      if (attachments && attachments.length > 0) {
        attachments.forEach((attachment, index) => {
          Logger.log(`  添付ファイル${index + 1}:`);
          Logger.log(`  ファイル名: ${attachment.getName()}`);
        });
      } else {
        Logger.log(`  添付ファイル: なし`);
      }

      Logger.log('デバッグモードのため、メールは送信されません。');
      return true; 
    }

      // メール送信
      GmailApp.sendEmail(
        mailAddress,
        subject,
        body,
        {
          htmlBody: htmlBody,
          name: CONFIG_CLAIM_Ma.GMAIL.SENDER_NAME, // 送信者名
          from: CONFIG_CLAIM_Ma.GMAIL.SENDER_EMAIL_ADDRESS, // 送信者メールアドレス
          attachments: attachments
        }
      );

      // 送信したメールの情報（引数で受け取った内容）をログ出力
      Logger.log(`送信メール情報: 宛先: ${mailAddress}, 件名: ${subject}`); 
      Logger.log(`${body.substring(0, 100)}`);
      

      // 添付ファイルの情報をログ出力
      if (attachments && attachments.length > 0) {
        attachments.forEach((attachment, index) => {
        Logger.log(`  添付ファイル${index + 1}: ファイル名: ${attachment.getName()}`);
        });
      } else {
        Logger.log(`  添付ファイル: なし`);
      }
      Logger.log(` ■ ${mailAddress} へのメールを送信しました。`);

    return true;

  } catch (e) {
    Logger.log(`メール送信に失敗しました。: ${e.message}`);
    return false;
  }
}

// --- ここから共通関数・ヘルパー関数 ---

/**
* 設定シートから顧客パターンと日付、送信内容が入ってるか、を見て、今回必要なセルを見つける
*/
function locateTargetCells_(timing, customerPattern) {

  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_Ma.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_Ma.SETTEI_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  // 必要な列の列番号を、設定情報から先に全て取得してしまう
  const colNums = {};
  for (const key in CONFIG_CLAIM_Ma.MESSAGE_COLUMNS) {
    colNums[key] = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.MESSAGE_COLUMNS[key]) - 1;
  }

  // valuesの各行から、タイミングとカスタマーパターンが一致する行を探索
  for (let i = 1; i < values.length; i++) {
    // タイミングとカスタマーパターンが一致する行から、必要な情報を取得して返す
    if (values[i][colNums.TIMING] === timing && values[i][colNums.PATTERN] === customerPattern) {
      
      const row = i + 1;  // スプレッドシートの実際の行番号（0始まりのインデックスに1を足す）

      return {
        kenmeiCell: sheet.getRange(row, colNums.KENMEI + 1),
        honbunCell: sheet.getRange(row, colNums.HONBUN + 1),
        gazou_1Cell: values[i][colNums.GAZOU_1] ? sheet.getRange(row, colNums.GAZOU_1 + 1) : null,
        gazou_1_nameCell: values[i][colNums.GAZOU_1_NAME] ? sheet.getRange(row, colNums.GAZOU_1_NAME + 1) : null,
        gazou_2Cell: values[i][colNums.GAZOU_2] ? sheet.getRange(row, colNums.GAZOU_2 + 1) : null,
        gazou_2_nameCell: values[i][colNums.GAZOU_2_NAME] ? sheet.getRange(row, colNums.GAZOU_2_NAME + 1) : null
      };
    }
  }
  throw new Error(`${timing}、パターン${customerPattern}に対応するメッセージが見つかりません。`);
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

  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_Ma.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_Ma.SETTEI_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  // 督促送信内容シートの列設定から、必要な列の番号を取得
  const timingColumnNumber = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.MESSAGE_COLUMNS.TIMING);
  const offsetColumnNumber = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.MESSAGE_COLUMNS.OFFSET);

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
* Googleドキュメントから本文を取得
*/
function getDocumentContent_(documentUrl) {
  try {
    const docId = documentUrl.match(/\/d\/(.+)\//)[1];
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } catch (error) {
    Logger.log('ドキュメント処理エラー: ' + error.message);
    throw error;
  }
}

/**
* プレースホルダーの置換処理
*/
function replaceAllPlaceholders_(text, placeholders) {
  return text
    .replace('{userId}', placeholders.userId)
    .replace('{kingaku}', placeholders.kingaku)
    .replace('{sei}', placeholders.sei)
    .replace('{mei}', placeholders.mei)
    .replace('{kigen}', placeholders.kigen);
}


/**
* 送った履歴をスプシに入れる関数
*/
function recordSentDate(recordSentData) {

  const ss = SpreadsheetApp.openById(CONFIG_CLAIM_Ma.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_CLAIM_Ma.SHEET_NAME);
  const sentDateColumn = GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.SENTDATE_OF_M);

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
        CONFIG_CLAIM_Ma.TIME_ZONE,
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
    const lastName = sheet.getRange(data.row, GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.LAST_NAME)).getValue();
    const firstName = sheet.getRange(data.row, GeneralUtils.columnToNumber_(CONFIG_CLAIM_Ma.TARGET_COLUMNS.FIRST_NAME)).getValue();
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
          CONFIG_CLAIM_Ma.TIME_ZONE,
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



/**
* パブリックAPIとして公開するため
*/
  return {
    mailClaimMain: mailClaimMain
  };
})();


/**
* トリガーのためのトリガーから叩かれる関数
* 名前空間より外に置いとく必要ある
*/
function setMailClaimTrigger() {

  // 既存のトリガーをすべて削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "mailClaimMain") { //この名前のトリガーを探して消す
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const time = new Date();
  
  time.setHours(7);
  time.setMinutes(03);
  
  // 指定した時分が現在より過去の場合は、翌日のその時刻に設定
  if (time < new Date()) {
    time.setDate(time.getDate() + 1);
  }
  
  ScriptApp.newTrigger('mailClaimMain').timeBased().at(time).create();
}
