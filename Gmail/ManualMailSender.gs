/**
 * =====================================================
 * Mail任意（Manual）送信Bot  
 *
 * スプシの原稿用シートの情報を元に、各メアドに送る。
 * A列にnowってあれば即時送信
 * それ以外は予約送信、と条件分岐が自動でできる。
 * 送信ログシートは予約送信待ちのメッセのストレージも兼ねる。
 * 予約送信ならトリガー作って、指定時刻になったらトリガーが発火して、ストレージから情報引っ張って、メール実送信関数に渡す
 * =====================================================
 */
const ManualMailSender = (function() {

const MANU_CONFIG_MAIL = {

    DEBUG_MODE: false, // true: デバッグモード, false: 本番モード
    
    SS_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1C6DMs',
    MAIN_SH: '統一シート',
    GENKOU_SH: '任意送信原稿',
    LOG_SH: '予約ｽﾄﾚｰｼﾞ兼送信ﾛｸﾞ',
    
    ALLOWED_USER: 'worldsecret.bm@gmail.com', // このアカウントでログインしないと、送信者名変える時にエラー出る


    //任意送信原稿シート内の列の設定
    COL_BORDER: 'Z',
    COL_ADDRESS: 'P',
    COL_SENDER_NAME: 'Q',
    COL_SENDER_EMAIL_ADDRESS: 'R',
    COL_KENMEI: 'S',
    COL_HONBUN:'T',

    COL_ATTACHMENTS: {
        IMAGE_1: 'U',
        IMAGE_NAME_1: 'V',
        IMAGE_2: 'W',
        IMAGE_NAME_2: 'X',
        IMAGE_3: 'Y',
        IMAGE_NAME_3: 'Z',
    },


    // 統一シートの列設定
    MAIN_SH_COLUMNS: {
        ADDRESS: 'D', // メアド
        LAST_NAME: 'F', // 姓
        FIRST_NAME: 'G', // 名（空欄の場合もあり）
    },

    // 任意送信ログシートの列設定
    LOG_SH_COLUMNS: {
        SENDING_ID: 'A',
        KEY_ROW: 'B',
        GROUP_INDEX: 'C',
        TARGET_USER_COUNT: 'D',
        RESERVATION_TIME: 'E',
        PROCESSED: 'F',
        ADDRESS: 'G',
        SEI: 'H',
        MEI: 'I',

        TIMESTAMP: 'V',
        TRIGGER_ID: 'W',

        FILE_PATH: 'Y',
        KENMEI: 'Z',
        HONBUN_PREVIEW: 'AA',
        
        IMAGE_1: 'AB',
        IMAGE_NAME_1: 'AC',
        IMAGE_2: 'AD',
        IMAGE_NAME_2: 'AE',
        IMAGE_3: 'AF',
        IMAGE_NAME_3: 'AG',
    },
    DRIVE_FOLDER_ID: '1m8DnseISuaVGqlJofTfKAF-QBVxzMEhT'

};

/**
 * =====================================================
 * メインエントリーポイント（メール）
 * =====================================================
 */
function manuGmailMain() {
    try {
        // ユーザーチェック
        const currentUser = Session.getEffectiveUser().getEmail();
        if (currentUser !== MANU_CONFIG_MAIL.ALLOWED_USER) {
            const ui = SpreadsheetApp.getUi();
            ui.alert(
                '注意',
                `メール送信は ${MANU_CONFIG_MAIL.ALLOWED_USER} で実行する必要があります。\n` +
                `現在 ${currentUser} でログインしています。\n\n` +
                'アカウントを切り替えてから再度実行してください。',
                ui.ButtonSet.OK
            );
            return;
        }
        
        const groupRanges = getGroupRanges_();
        
        console.log("グループ範囲情報:");
        for (let i = 0; i < groupRanges.length; i++) {
            console.log(`Group ${i + 1}: start: ${groupRanges[i].start}, end: ${groupRanges[i].end}, keyRow: ${groupRanges[i].keyRow}`);
        }
        
        const preparedMessages = prepMailGenkouFromSS_(groupRanges);
        return preparedMessages;
    } catch (error) {
        try {
            SpreadsheetApp.getUi().alert(
                'エラー',
                error.message,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        } catch (uiError) {
            console.error('メール送信処理でエラーが発生:', error.message);
        }
        throw error;
    }
}

/**
 * =====================================================
 * ヘルパー関数
 * =====================================================
 */

// プレースホルダー置換
function replacePlaceholders1_(text, userData) {
    const sei = userData && userData.sei ? userData.sei : '';
    const mei = userData && userData.mei ? userData.mei : '';
    return text
        .replace(/\{sei\}/g, sei)
        .replace(/\{mei\}/g, mei);
}

// 列名を列番号に変換する
function columnToNumber1_(column) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
        result *= 26;
        result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return result;
}


//スプレッドシートから予約送信グループの範囲とキー行を取得する関数
function getGroupRanges_() {
    const sheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID).getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH);

    // メアドの列番号を取得
    const addressColNumber = columnToNumber1_(MANU_CONFIG_MAIL.COL_ADDRESS);

    // A列からメアド列までを一度に取得
    const data = sheet.getRange(1, 1, sheet.getLastRow(), addressColNumber).getValues();

    // 最終行を判定（ユーザーID列のデータを使用）
    let lastRow = data.reduceRight((acc, row, idx) => {
        if (acc === 0 && row[addressColNumber - 1] !== "") acc = idx + 1;
        return acc;
    }, 0);

    // A列に値がある行番号のリストを取得（1行目は見出し行なのでスキップ）
    const startRowsNumbers = data.slice(1, lastRow).reduce((acc, row, idx) => {
        if (row[0] !== "") acc.push(idx);
        return acc;
    }, []);

    // 各グループの行範囲とキー行を特定
    const groupRanges = [];
    for (let i = 0; i < startRowsNumbers.length; i++) {
        const eachStartRow = startRowsNumbers[i];
        const eachEndRow = (i + 1 < startRowsNumbers.length) ? startRowsNumbers[i + 1] - 1 : lastRow - 2;

        const sheetStRow = eachStartRow + 2;
        const sheetEdRow = eachEndRow + 2;

        groupRanges.push({ start: sheetStRow, end: sheetEdRow, keyRow: sheetStRow });
    }

    return groupRanges;
}


/**
 * =====================================================
 * 実際の処理
 * =====================================================
 */
/**
 * メール送信用のメッセージ原稿を準備する関数
 */
function prepMailGenkouFromSS_(groupRanges) {
    const sheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID).getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH);

    // 各送信行為ごとに送る内容をまとめて、次の関数に渡すための変数
    const packedMaterialsWithKeyrow = [];

    // 引数として受け取った groupRanges を使用し、キー行を抽出
    const keyRowsToProcess = groupRanges.map(group => group.keyRow);

    // 選択したキー行を使ってメイン処理を行う
    for (const row of keyRowsToProcess) {
        try {
            // 本文の取得と処理
            const bodyColumn = MANU_CONFIG_MAIL.COL_HONBUN;
            const bodyColumnNum = columnToNumber1_(bodyColumn);
            const bodyCell = sheet.getRange(row, bodyColumnNum);
            let bodyContent = '';

            // 本文セルの処理
            const richTextValue = bodyCell.getRichTextValue();
            const runs = richTextValue.getRuns();
            let foundInRichText = false;

            // まずリッチテキストリンクをチェック
            for (let run of runs) {
                const linkUrl = run.getLinkUrl();
                if (linkUrl && linkUrl.includes('google.com/')) {
                    try {
                        const fileId = GeneralUtils.getFileIdFromDriveUrl_(linkUrl);
                        Logger.log(`docsファイルID抽出成功: ${fileId}`);
                        const mimeType = GeneralUtils.getMimeType_(fileId);

                        if (mimeType === 'application/vnd.google-apps.document') {
                            const doc = DocumentApp.openById(fileId);
                            bodyContent = doc.getBody().getText();
                            foundInRichText = true;
                            break;
                        } else {
                            console.error(`本文列に非対応の形式のファイルが指定されています: ${mimeType}`);
                            continue;
                        }
                    } catch (error) {
                        console.error(`Googleドキュメントの処理でエラー: ${error.message}`);
                        continue;
                    }
                }
            }

            // リッチテキストで見つからなかった場合のみ、直接入力テキストを使用
            if (!foundInRichText) {
                const directText = bodyCell.getValue();
                if (directText) {
                    bodyContent = directText;
                }
            }

            // 件名の取得（本文取得後に移動）
            const subjectColumn = MANU_CONFIG_MAIL.COL_KENMEI;
            const subjectColumnNum = columnToNumber1_(subjectColumn);
            const subjectCell = sheet.getRange(row, subjectColumnNum);
            let subject = subjectCell.getValue();

            // 件名が空欄の場合、本文1行目を件名に
            if (!subject && bodyContent) {
                const lines = bodyContent.split('\n');
                subject = lines[0];
                bodyContent = lines.slice(1).join('\n');
            }

            // 送信者情報の取得
            const senderName = sheet.getRange(row, columnToNumber1_(MANU_CONFIG_MAIL.COL_SENDER_NAME)).getValue();
            const senderEmail = sheet.getRange(row, columnToNumber1_(MANU_CONFIG_MAIL.COL_SENDER_EMAIL_ADDRESS)).getValue();

            // 空欄チェック
            if (!senderName || !senderEmail) {
                throw new Error(`${row}行目: 送信者名または送信者メールアドレスが未入力です。`);
            }

            // 添付画像の処理
            const attachments = [];

            // 添付画像の列を順に処理
            for (let i = 1; i <= 3; i++) {
                const imageColumn = MANU_CONFIG_MAIL.COL_ATTACHMENTS[`IMAGE_${i}`];
                const imageNameColumn = MANU_CONFIG_MAIL.COL_ATTACHMENTS[`IMAGE_NAME_${i}`];
                
                if (imageColumn && imageNameColumn) {
                    const imageCell = sheet.getRange(row, columnToNumber1_(imageColumn));
                    const imageNameCell = sheet.getRange(row, columnToNumber1_(imageNameColumn));
                    
                    // セルが空でない場合のみ処理を行う
                    if (imageCell.getValue()) {
                        const imageBlob = getImageBlob_(imageCell, imageNameCell);
                        if (imageBlob) {
                            attachments.push(imageBlob);
                        }
                    }
                }
            }

            // 処理結果を配列に追加
            packedMaterialsWithKeyrow.push({
                keyRow: row,
                messageContent: {
                    subject: subject,
                    body: bodyContent,
                    attachments: attachments,
                    senderName: senderName,    // 追加
                    senderEmail: senderEmail   // 追加
                }
            });

        } catch (error) {
            console.error(`行 ${row} の処理中にエラーが発生: ${error.message}`);
            continue;
        }
    }

    // メッセージと宛先の紐付け処理へ
    return prepareMessagesForUsers_(groupRanges,packedMaterialsWithKeyrow);
}

/**
 * 画像のリンクについて、原稿シートに画像URL直打ちでも、ファイル名リンクにしても読み取ってくれる関数
 * @param {Range} imageCell - 画像URLまたはリンクが含まれるセル
 * @param {Range} nameCell - 画像のファイル名が含まれるセル
 * @return {Blob|null} - 取得した画像のBlobデータ、取得できない場合はnull
 */
function getImageBlob_(imageCell, nameCell) {
  if (!imageCell || !imageCell.getValue()) {
    Logger.log("imageCell が null または空です。");
    return null;
  }

  try {
    let blob;
    const url = imageCell.getValue().trim();

    let richTextValue = imageCell.getRichTextValue();
    if (richTextValue && richTextValue.getRuns().some(run => run.getLinkUrl())) {
      Logger.log("画像欄にリッチテキストを検出しました。");

      const runs = richTextValue.getRuns();
      for (let run of runs) {
        const linkUrl = run.getLinkUrl();
        if (linkUrl && linkUrl.includes('google.com/')) {
          try {
            const fileId = GeneralUtils.getFileIdFromDriveUrl_(linkUrl);
            const file = DriveApp.getFileById(fileId);
            const mimeType = file.getMimeType();
            Logger.log(`MIMEタイプ: ${mimeType}`);

            if (mimeType.includes('image')) {
              blob = file.getBlob();
              
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
          } catch (e) {
            Logger.log(`画像のGoogleドライブURL処理エラー: ${e.message}`);
            continue;
          }
        }
      }
    }

    // 既存の処理（リッチテキストで画像が見つからなかった場合のフォールバック）
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
 * 作った原稿をどこに送ればいいのか特定。特定できたらプレースホルダー置換
 */
function prepareMessagesForUsers_(groupRanges,packedMaterialsWithKeyrow) {
  console.log('=== 原稿と宛先の紐づけ開始 ===');

  const ss = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID);
  const genkouSheet = ss.getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH);
  const mainSheet = ss.getSheetByName(MANU_CONFIG_MAIL.MAIN_SH);

  // groupRanges とpackedMaterialsWithKeyrow を keyRow で結合
  const combinedData = groupRanges.map(group => {
    const correspondingGenkou =packedMaterialsWithKeyrow.find(item => item.keyRow === group.keyRow);
    return {
      keyRow: group.keyRow,
      sheetStRow: group.start,
      sheetEdRow: group.end,
      messageContent: correspondingGenkou ? correspondingGenkou.messageContent : null
    };
  });

  // ユーザーデータマップを作成 (キー: メールアドレス, 値: {sei: 姓, mei: 名})
  const userDataMap = new Map();
  const lastRow = mainSheet.getLastRow();
  const mailAddresses = mainSheet.getRange(`${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.ADDRESS}2:${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.ADDRESS}${lastRow}`).getValues();
  const lastNames = mainSheet.getRange(`${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.LAST_NAME}2:${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.LAST_NAME}${lastRow}`).getValues();
  const firstNames = mainSheet.getRange(`${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.FIRST_NAME}2:${MANU_CONFIG_MAIL.MAIN_SH_COLUMNS.FIRST_NAME}${lastRow}`).getValues();

  for (let i = 0; i < mailAddresses.length; i++) {
    const mailAddress = mailAddresses[i][0];
    if (mailAddress) {
      const sei = lastNames[i][0] || '';
      const mei = firstNames[i][0] || '';
      userDataMap.set(mailAddress, { sei, mei });
    }
  }

  // 各送信行為ごとに分けてループ処理
  const messageGroups = [];

  combinedData.forEach((group, groupIndex) => {
    let processedCount = 0;
    let skippedCount = 0;
    let duplicatedCount = 0;

    console.log(`▼ グループ${groupIndex + 1} (${group.sheetStRow}行目～${group.sheetEdRow}行目)`);
    const keyRow = group.keyRow;

    // 宛先特定: メールアドレス列から対象範囲の値を取得
    const mailAddressRange = genkouSheet.getRange(
      group.sheetStRow,
      columnToNumber1_(MANU_CONFIG_MAIL.COL_ADDRESS),
      group.sheetEdRow - group.sheetStRow + 1,
      1
    );
    const mailAddresses = mailAddressRange.getValues().flat().filter(addr => addr !== '');

    console.log(`送信先数: ${mailAddresses.length}件`);

    // マッチング処理
    const matchedUserData = new Map();
    let matchedUserCount = 0;
    mailAddresses.forEach(address => {
      if (userDataMap.has(address)) {
        matchedUserData.set(address, userDataMap.get(address));
        matchedUserCount++;
      }
    });
    console.log(`マッチしたユーザーデータ数: ${matchedUserCount}件`);

    // マッチしたメールアドレス、姓、名を出力
    for (const [address, userData] of matchedUserData) {
      console.log(`- メールアドレス: ${address}, 姓: ${userData.sei}, 名: ${userData.mei}`);
    }

    const now = new Date();
    const processedAddresses = new Set();

    // 各宛先ごとに原稿と紐づけループ処理
    mailAddresses.forEach((address) => {

      if (processedAddresses.has(address)) {
        console.warn(`警告: メールアドレス ${address} が重複しています。このアドレスはスキップされます。 - グループ: ${groupIndex + 1}, キー行: ${keyRow}`);
        duplicatedCount++;
        return;
      }

      if (userDataMap.has(address)) {
        const userData = matchedUserData.get(address);

        // 送信IDの生成
        const timestampStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
        const addressSuffix = address.replace(/[^a-zA-Z0-9]/g, '').slice(-8);
        const sendingId = `${timestampStr}_GRP${groupIndex + 1}_ADR${addressSuffix}`;

        // メール本文のプレースホルダー置換
        const messageContent = { ...group.messageContent };
        if (messageContent.body) {
          messageContent.body = replacePlaceholders1_(messageContent.body, userData);
        }
        if (messageContent.subject) {
          messageContent.subject = replacePlaceholders1_(messageContent.subject, userData);
        }

        processedCount++;
        processedAddresses.add(address);

        // 宛先と紐づいた原稿一式をまとめる
        messageGroups.push({
          groupIndex: groupIndex,
          keyRow: keyRow,
          address: address,
          userData: userData,
          messageContent: messageContent,
          timestamp: now,
          sendingId: sendingId
        });

      } else {
        console.error(`エラー: メールアドレス ${address} が統一シートに見つかりません。このアドレスへの送信をスキップします。 - グループ: ${groupIndex + 1}, キー行: ${keyRow}`);
        skippedCount++;
      }
    });
    console.log(`グループ${groupIndex + 1}の処理完了: 処理済み: ${processedCount}件, スキップ: ${skippedCount}件, 重複: ${duplicatedCount}件`);
  });

  // nowOrDelay_にgroupedData（keyRowと宛先数）と、messageGroups（宛先と紐づいた原稿一式）を渡す
  nowOrDelay_({
    groupedData: combinedData.map(group => ({
      keyRow: group.keyRow,
      targetUserCount: messageGroups.filter(msg => msg.keyRow === group.keyRow).length 
    })),
    messageGroups: messageGroups
  });
}


/**
 * 予約送信の日時を生成し、過去の日時をバリデーションする関数
 * @param {number} keyRow - キー行番号
 * @returns {string} 'MM/dd HH:mm' 形式の日時文字列
 * @throws {Error} バリデーションエラー時
 */
function createReservationTime_(keyRow) {
  const genkouSheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID)
      .getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH);

  const dateValue = genkouSheet.getRange(keyRow, 1).getValue();  // A列 日付
  const hour = genkouSheet.getRange(keyRow, 3).getValue();       // C列 時
  const minute = genkouSheet.getRange(keyRow, 4).getValue();     // D列 分

  // null/undefined チェック（0は許容。そうしないと0時とか0分に送れない）
  if (!dateValue || 
      hour === null || hour === undefined || 
      minute === null || minute === undefined) {
      throw new Error(`${keyRow}行目: 日付・時刻のいずれかが未入力です。`);
  }

  // NOW の場合の処理
  if (typeof dateValue === 'string' && dateValue.toString().toUpperCase().includes('NOW')) {
      return 'NOW';
  }

  // 日付型じゃない場合のエラー
  if (!(dateValue instanceof Date)) {
      throw new Error(`${keyRow}行目: A列は日付型かnowで入力してください。（現在の値: ${dateValue}）`);
  }

  // 日付から月日を抽出
  const month = dateValue.getMonth() + 1;
  const day = dateValue.getDate();

  // 予約日時のDateオブジェクトを生成
  const reservationDate = new Date();
  reservationDate.setMonth(month - 1);
  reservationDate.setDate(day);
  reservationDate.setHours(parseInt(hour));
  reservationDate.setMinutes(parseInt(minute));
  reservationDate.setSeconds(0);

  // 現在時刻との比較
  const now = new Date();
  if (reservationDate < now) {
      const formattedReservationTime = `${month}月${day}日 ${hour}時${minute}分`;
      throw new Error(
          `${keyRow}行目: ${formattedReservationTime}は過去の日時です。\n` +
          `予約時刻は現在時刻より後の時刻を指定してください。`
      );
  }

  // ゼロ埋めして2桁にする
  const paddedMonth = String(month).padStart(2, '0');
  const paddedDay = String(day).padStart(2, '0');
  const paddedHour = String(hour).padStart(2, '0');
  const paddedMinute = String(minute).padStart(2, '0');

  // MM/dd HH:mm 形式で返す
  return `${paddedMonth}/${paddedDay} ${paddedHour}:${paddedMinute}`;
}


/**
 * 即時送信か予約送信かを判定し、振り分けを行う。
 * 即時送信だったら、ここで実送信→ログシート書込み→送信成功フラグ更新→原稿削除まで行う
 * 予約送信だったら、ログシート（予約ストレージ）書込み→原稿削除
 * 各keyRowごとに、keyRowの個数分ループ処理
 */

function nowOrDelay_(preparedData) {
  // prepareMessagesForUsers_から渡されたデータを解凍し、必要な情報取得
  const { groupedData, messageGroups } = preparedData;

  // 各keyRowグループのキー行をチェックし、即時か予約か判定→ここで全keyRowグループ分、drop〜〜関数をループさせてる
  groupedData.forEach(group => { //アロー関数でgroupedDataの中身をgroup.●●、で取り出せるようにしてる
    
    const keyRow = group.keyRow;
    const sendTypeCell = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID)
      .getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH)
      .getRange(keyRow, 1); // A列
    const sendType = sendTypeCell.getValue();

    // keyRow に紐づく messageGroups のデータを取得
    const targetMessageGroups = messageGroups.filter(group => group.keyRow === keyRow);


    // 即時送信の場合！
    if (sendType.toString().toUpperCase().includes('NOW')) {
      console.log(`keyRow ${keyRow} は即時送信`);

      // 即時送信用のタイムスタンプを生成
      const now = new Date();
      const timestampForLog = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

      // 送信結果を格納する配列
      const results = [];
      let allSucceeded = true; // 全て成功したかどうかのフラグ

      // 即時送信の場合は、グループ内の各メールアドレスに対して、ここでループ処理を行う必要がある
      // ただしログシートには、まとめて記録
      targetMessageGroups.forEach(group => {

        // メール送信用の関数を呼び出す
        const eachResult = sendMailMessage_(
          group.address, 
          group.messageContent.subject,
          group.messageContent.body,
          group.messageContent.attachments,
          group.groupIndex, 
          group.keyRow, 
          group.sendingId,
          group.messageContent.senderName,    // 追加
          group.messageContent.senderEmail    // 追加
        );
        results.push(eachResult);

        if (!eachResult.success) { // 1つでも送信失敗したらフラグを false に
            allSucceeded = false;
        }
      }); // ここまでループして、まとめたデータを一括でログシート記録処理↓に渡す

      // まとめてログシートに書き込むためにデータ渡す
      if (targetMessageGroups.length > 0) {
        dropToStorage_('NOW', keyRow, targetMessageGroups, {
          keyRow: group.keyRow,
          targetUserCount: targetMessageGroups.length
        }, timestampForLog);
      }

      // 処理済みフラグ更新関数呼び出し
      if (results.length > 0) {
        flagOfSuccessOrFalse_(results);
      }

      // 送信完了したグループの原稿をクリア
      if (allSucceeded) {
        
        // 行範囲の一覧を取得
        const groupRanges = getGroupRanges_();
        
        // 処理したkeyRowに対応する行範囲を特定
        const targetGroup = groupRanges.find(group => group.keyRow === keyRow);
        // クリア処理
        if (targetGroup) {
          clearYouzumiGenkou_(targetGroup.start, targetGroup.end);
        } else {
          console.error(`キー行 ${keyRow} に対応するグループが見つかりませんでした`);
        }
      }

    } else {

      // 予約送信の場合！
      try {
        const reservationTime = createReservationTime_(keyRow);
        console.log(`keyRow ${keyRow}は予約送信, 予約日時: ${reservationTime}`);

        // 予約時刻までスプシに完全な状態で保管のため、次の関数呼び出す
        dropToStorage_(reservationTime, keyRow, targetMessageGroups, group, null);

        // 原稿クリア処理
        const groupRanges = getGroupRanges_();
        const targetGroup = groupRanges.find(group => group.keyRow === keyRow);
        
        if (targetGroup) {
          clearYouzumiGenkou_(targetGroup.start, targetGroup.end);
        } else {
          console.error(`キー行 ${keyRow} に対応するグループが見つかりませんでした`);
        }

      } catch (error) {
        console.error(`予約送信設定エラー: keyRow ${keyRow}, エラー内容: ${error.message}, スタックトレース: ${error.stack}`);
        // エラーが発生した場合は、このグループの処理をスキップして、次のグループの処理を続ける。
      }
    }
  });
}

/**
 * 各データをログシート（予約ストレージ）に書込む
 * 予約送信ならトリガーもセット
 */
function dropToStorage_(nowOrReservationTime, keyRow, targetMessageGroups, currentGroup, timestampForLog) {

  const ss = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID);
  
  console.log("▼ keyRow" ,keyRow,"のグループの予約ストレージ（送信ログ）書込み開始");
  ss.toast(`キー行${keyRow}のグループの予約ストレージ書込みを開始します。※即時送信（now）の場合は送信完了後のログ書込み`, "処理状況");
  console.log("予約日時:", nowOrReservationTime);
  console.log("グループデータ:", currentGroup);
  console.log("メッセージグループ:", JSON.stringify(targetMessageGroups.map(g => ({
  ...g,
  groupIndex: g.groupIndex + 1,
  messageContent: {
    ...g.messageContent,
    body: g.messageContent.body.substring(0, 100) + '...'
  }
  }))));
  

  // 任意送信ログシートを取得
  const logSheet = ss.getSheetByName(MANU_CONFIG_MAIL.LOG_SH);

  // Google ドライブのフォルダを取得
  const folder = DriveApp.getFolderById(MANU_CONFIG_MAIL.DRIVE_FOLDER_ID);

  // ログシートに書き込むデータ
  let logData = [];


  // 今から処理するグループは、全部で○人に送る△番目のグループです、と後続処理に伝える準備
  if (!currentGroup) {
    console.error(`キー行 ${keyRow} に対応するグループデータが見つかりません`);
    return;
  }

  const groupIndex = targetMessageGroups[0].groupIndex + 1;
  const targetUserCount = currentGroup.targetUserCount; 

  // targetMessageGroups が配列であることを念のため確認
  if (!Array.isArray(targetMessageGroups)) {
    console.error('targetMessageGroups is not an array:', targetMessageGroups);
    return;
  }

  // 各keyRowグループ内に存在する各メアドに対しての処理
  targetMessageGroups.forEach(group => {

    if (!group || !group.messageContent) {
      console.error('Invalid group structure:', group);
      return;
    }

    const sendingId = group.sendingId;
    const address = group.address;
    // 添付ファイル情報を最初に取得
    const messageAttachments = group.messageContent.attachments || [];

    // このログのせいで遅い可能性あるから一旦コメントアウト
    //console.log(`予約ｽﾄﾚｰｼﾞ（送信ログ）用データ蓄積中: ${address}`);


    // 送信予約時刻をフォーマット
    let formattedReservationTime = '';
    let timestamp = '';

    if (nowOrReservationTime === 'NOW') {
      timestamp = timestampForLog;
      formattedReservationTime = Utilities.formatDate(new Date(timestampForLog), 'Asia/Tokyo', 'MM/dd HH:mm') + ' (即時送信)';
    } else {
      timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      formattedReservationTime = nowOrReservationTime;
    }

    // メール内容をJSONファイルとして保存
    const mailContent = {
      subject: group.messageContent.subject,
      body: group.messageContent.body,
      attachments: [],
      senderName: group.messageContent.senderName,   
      senderEmail: group.messageContent.senderEmail
    };

    // 添付ファイルの処理
    if (messageAttachments.length > 0) {
      messageAttachments.forEach((attachment, index) => {
        try {
          // 添付ファイル用のフォルダに保存
          const fileName = `attachment_${sendingId}_${index + 1}_${attachment.getName()}`;
          const file = folder.createFile(attachment);
          const filePath = file.getUrl();
          
          mailContent.attachments.push({
            path: filePath,
            name: attachment.getName(),
            type: attachment.getContentType()
          });
        } catch (error) {
          console.error(`添付ファイル処理エラー: ${error.message}`);
        }
      });
    }

    // JSONファイルとして保存
    const fileName = `mail_${sendingId}.json`;
    const file = folder.createFile(fileName, JSON.stringify(mailContent), 'application/json');
    const filePath = file.getUrl();
    
    // 全列分の配列を初期化（関数冒頭で取得したlogSheetを再利用）
    let logRow = Array(logSheet.getLastColumn()).fill('');

    // 冒頭設定に従って各値を配置
    const cols = MANU_CONFIG_MAIL.LOG_SH_COLUMNS;
    logRow[columnToNumber1_(cols.SENDING_ID) - 1] = sendingId;
    logRow[columnToNumber1_(cols.KEY_ROW) - 1] = keyRow;
    logRow[columnToNumber1_(cols.GROUP_INDEX) - 1] = groupIndex;
    logRow[columnToNumber1_(cols.TARGET_USER_COUNT) - 1] = targetUserCount;
    logRow[columnToNumber1_(cols.RESERVATION_TIME) - 1] = formattedReservationTime;
    logRow[columnToNumber1_(cols.PROCESSED) - 1] = 0;
    logRow[columnToNumber1_(cols.ADDRESS) - 1] = address;
    logRow[columnToNumber1_(cols.SEI) - 1] = group.userData.sei;
    logRow[columnToNumber1_(cols.MEI) - 1] = group.userData.mei;
    logRow[columnToNumber1_(cols.TIMESTAMP) - 1] = timestamp;
    logRow[columnToNumber1_(cols.TRIGGER_ID) - 1] = '';
    logRow[columnToNumber1_(cols.FILE_PATH) - 1] = filePath;
    logRow[columnToNumber1_(cols.KENMEI) - 1] = group.messageContent.subject;
    logRow[columnToNumber1_(cols.HONBUN_PREVIEW) - 1] = group.messageContent.body.substring(0, 100).replace(/\r?\n/g, ' ');

    // 添付ファイル情報を追加
    if (messageAttachments.length > 0) {
        for (let i = 1; i <= 3; i++) {
            const attachment = messageAttachments[i - 1];
            if (attachment) {
                logRow[columnToNumber1_(cols[`IMAGE_${i}`]) - 1] = attachment.url || '';
                logRow[columnToNumber1_(cols[`IMAGE_NAME_${i}`]) - 1] = attachment.getName() || '';
            }
        }
    }

    // ログデータに追加
    logData.push(logRow);
  });

  // keyRowグループ単位でまとめて書き込み
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logData.length, logData[0].length).setValues(logData);

  //ここに書込み完了ログ出したほうがいい？
  
  // 予約送信の場合はトリガーを設定！
  if (nowOrReservationTime !== 'NOW') {
    setTriggerForDelay_(nowOrReservationTime, keyRow, targetMessageGroups);
  }
}


/**
 * 予約送信用のトリガーを設定する関数
 * @param {string} reservationTime - 'MM/dd HH:mm' 形式の予約時刻
 * @param {number} keyRow - 任意送信原稿シートのキー行番号
 * @param {array} targetMessageGroups - 送信グループごとのメッセージ情報を含む配列（修正箇所: 引数を追加）
 */
function setTriggerForDelay_(reservationTime, keyRow, targetMessageGroups) {
  try {
    const ss = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID);

    // 予約時刻をDateオブジェクトに変換
    const [datePart, timePart] = reservationTime.split(' ');
    const [month, day] = datePart.split('/');
    const [hours, minutes] = timePart.split(':');
    
    const triggerDate = new Date();
    triggerDate.setMonth(parseInt(month) - 1);
    triggerDate.setDate(parseInt(day));
    triggerDate.setHours(parseInt(hours));
    triggerDate.setMinutes(parseInt(minutes));
    triggerDate.setSeconds(0);

    // トリガーを作成
    const trigger = ScriptApp.newTrigger('triggerPointForMailManu')
      .timeBased()
      .at(triggerDate)
      .create();


    // トリガーIDをログシートに記録
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID)
      .getSheetByName(MANU_CONFIG_MAIL.LOG_SH);
    
    // このkeyRowに紐づく現在処理中のすべての送信IDのベース部分を取得
    const currentSendingIdBases = targetMessageGroups
      .filter(group => group.keyRow === keyRow) // keyRowが一致するグループのみをフィルタリング
      .map(group => {
        // 送信IDから「タイムスタンプ_GRP番号」の部分を抽出
        const idParts = group.sendingId.split('_');
        return `${idParts[0]}_${idParts[1]}_${idParts[2]}`; // 例: 20231027_103000_GRP1
      })
      .filter((value, index, self) => self.indexOf(value) === index); // 重複除去

    // ログシートのデータを取得
    const lastRow = logSheet.getLastRow();
    const logData = logSheet.getRange(2, 1, lastRow - 1, 23).getValues(); // A列からW列まで

    // 各行をチェックし、現在処理中の送信グループに属する行のみにトリガーIDを設定
    logData.forEach((row, index) => {
      const rowSendingId = row[0]; // A列の送信ID
      if (rowSendingId) {
        // この行の送信IDから「タイムスタンプ_GRP番号」部分までを抽出
        const idParts = rowSendingId.split('_');
        const rowSendingIdBase = `${idParts[0]}_${idParts[1]}_${idParts[2]}`;

        // 現在処理中の送信グループに属する行であれば、トリガーIDを設定
        if (currentSendingIdBases.includes(rowSendingIdBase)) {
          logSheet.getRange(index + 2, 23).setValue(trigger.getUniqueId()); // W列にトリガーIDを設定
        }
      }
    });

    const logMessage = `トリガーを設定しました - 予約時刻: ${reservationTime}, キー行: ${keyRow}, トリガーID: ${trigger.getUniqueId()}`;
    console.log(logMessage);
    ss.toast(logMessage);

  } catch (error) {
    console.error(`トリガー設定中にエラーが発生しました - キー行: ${keyRow}`, error);
    throw error;
  }
}


/**
 * =====================================================
 * ここから予約送信の場合は時間が空く
 * =====================================================
 */

/**
 * トリガーに叩かれて、ストレージのデータを引っ張ってメール送信関数にわたす
 */
function readStorageAndPassToMail_(e) {  // イベントパラメータを追加
  try {
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID).getSheetByName(MANU_CONFIG_MAIL.LOG_SH);

    // トリガーイベントから直接IDを取得
    const triggerId = e.triggerUid;
    console.log("発火したトリガーID:", triggerId);

    if (!triggerId) {
      console.warn('トリガーIDを取得できませんでした。');
      return;
    }

    // 処理対象の行を特定（同じトリガーIDで未送信のもの）
    const lastRow = logSheet.getLastRow();

    // 今後の処理に予約ストレージシートの各データ使うから、シート全部一括取得
    const data = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();
    const targetRows = [];

    data.forEach((row, index) => {
      const rowNum = index + 2;
      const processed = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.PROCESSED) - 1]; // F列の処理済みフラグ
      const currentTriggerId = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.TRIGGER_ID) - 1]; // W列のトリガーID
      const keyRow = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.KEY_ROW) - 1]; // B列のキー行
      const sendingId = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.SENDING_ID) -1]; // A列の送信ID
      const mailAddress = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.ADDRESS) - 1] // G列のメールアドレス


      if (currentTriggerId === triggerId && !processed) {
        targetRows.push({
          rowNum: rowNum,
          sendingId: sendingId,
          mailAddress: mailAddress,
          keyRow: keyRow,
          messageData: []
        });

        // filePath（JSONファイル）から件名、本文、添付ファイル情報を取得
        try {
          const filePath = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.FILE_PATH) - 1];
          if (filePath) {
            const fileId = filePath.replace(/.*\/d\/([^\/]+)\/.*/, '$1');
            const file = DriveApp.getFileById(fileId);
            const mailContent = JSON.parse(file.getBlob().getDataAsString());
            targetRows[targetRows.length - 1].messageData = mailContent;
          }
        } catch (error) {
          console.error(`メール内容の取得に失敗: ${error.message}`);
        }
      }
    });

    console.log(`送信対象データ: ${targetRows.length}件`);

    const results = []; // 送信結果を格納する配列

    // メール送信ループ
    for (const row of targetRows) {

        // F列を「処理中」の意味の「2」に更新
        logSheet.getRange(row.rowNum, columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.PROCESSED)).setValue(2);

        // メール送信関数の呼び出し
        const eachResult = sendMailMessage_(
          row.mailAddress,
          row.messageData.subject,
          row.messageData.body,
          row.messageData.attachments,
          null,
          null,
          row.sendingId,
          row.messageData.senderName,  
          row.messageData.senderEmail 
        );
        results.push(eachResult); // 各宛先の送信結果を蓄積
    }

    // 全送信完了後に一括で処理済みフラグを更新
    flagOfSuccessOrFalse_(results);

    // 全送信完了後、発火元のトリガーを直接削除
    if (triggerId) {
      const trigger = ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === triggerId);
      if (trigger) {
        ScriptApp.deleteTrigger(trigger);
        console.log(`トリガー削除完了: ${triggerId}`);
      }
    }
  
  } catch (error) { 
    console.error('予約送信処理でエラーが発生しました:', error.message, 'スタックトレース:', error.stack);
  } 
}    


/**
 * 送信結果に基づいてログシートの処理済みフラグを更新する関数
 * @param {Array<{sendingId: string, success: boolean}>} results - 送信結果の配列
 */
function flagOfSuccessOrFalse_(results) {
  if (!results || results.length === 0) {
    console.warn('flagOfSuccessOrFalse_: results が空です。処理をスキップします。');
    return;
  }

  try {
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID).getSheetByName(MANU_CONFIG_MAIL.LOG_SH);
    const lastRow = logSheet.getLastRow();

    // ログシートの全データを取得（A列から処理済みフラグのF列まで）
    const logData = logSheet.getRange(2, 1, lastRow - 1, columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.PROCESSED)).getValues();

    // 送信IDをキーとするマップを作成し、高速化を図る
    const sendingIdMap = new Map();
    results.forEach(result => {
      sendingIdMap.set(result.sendingId, result.success);
    });

    // ログシートのデータをループし、送信IDが一致する行の処理済みフラグを更新
    let updateCount = 0;
    logData.forEach((row, index) => {
      const sendingId = row[columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.SENDING_ID) - 1]; // A列の送信ID
      if (sendingIdMap.has(sendingId)) {
        const success = sendingIdMap.get(sendingId);
        const rowNum = index + 2; // 実際の行番号（見出し行を考慮）

        if (success) {
          logSheet.getRange(rowNum, columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.PROCESSED)).setValue(1); // 成功なら1
          updateCount++;
        } else {
          logSheet.getRange(rowNum, columnToNumber1_(MANU_CONFIG_MAIL.LOG_SH_COLUMNS.PROCESSED)).setValue('エラー'); // 失敗なら「エラー」
          updateCount++;
        }
      }
    });
    console.log(`flagOfSuccessOrFalse_: ${updateCount} 件の処理済みフラグを更新しました。`);

  } catch (error) {
    console.error('flagOfSuccessOrFalse_ でエラーが発生しました:', error.message, 'スタックトレース:', error.stack);
  }
}


/**
 * メールを実際に送信する関数 (Gmail API版)
 * @param {string} mailAddress - 送信先メールアドレス
 * @param {string} subject - メールの件名
 * @param {string} body - メールの本文
 * @param {Array} attachments - 添付ファイルの配列
 * @param {number|null} groupIndex - グループインデックス
 * @param {number|null} keyRow - キー行番号
 * @param {string} sendingId - 送信ID
 * @returns {Object} 送信結果 {sendingId: string, success: boolean}
 */
function sendMailMessage_(mailAddress, subject, body, attachments, groupIndex, keyRow, sendingId, senderName, senderEmail){
  try {
    // デバッグモードの場合は、メール内容をログに出力して、送信処理をスキップ
    if (MANU_CONFIG_MAIL.DEBUG_MODE) {
      const attachmentNames = attachments && attachments.length > 0 ? attachments.map(a => a.getName()).join(', ') : 'なし';
      console.log(`デバッグモード送信スキップ - ID: ${sendingId}, To: ${mailAddress}, 件名: ${subject}, 添付: ${attachmentNames}`);
      return { sendingId: sendingId, success: true }; // 送信成功とみなす
    }

    // 引数から取った送信者名とメールアドレスをRFC2047に従ってエンコード
    const encodedSenderName = '=?UTF-8?B?' + Utilities.base64Encode(
        Utilities.newBlob('').setDataFromString(senderName, 'UTF-8').getBytes()
    ) + '?=';
    const sender = `${encodedSenderName} <${senderEmail}>`;
    
    // 件名を一旦UTF-8のバイト配列に変換→base64エンコード→RFC2047形式でラップ
    const encodedSubject = '=?UTF-8?B?' + Utilities.base64Encode(
      Utilities.newBlob('').setDataFromString(subject, 'UTF-8').getBytes()
    ) + '?=';

    // MIMEメッセージの基本ヘッダー
    const headers = [
      'MIME-Version: 1.0',
      `From: ${sender}`,
      `To: ${mailAddress}`,
      `Subject: ${encodedSubject}`,
      'Content-Type: multipart/mixed; boundary="mixed_boundary"',
      ''
    ];

    // メッセージパーツの配列
    const parts = [];

    // メール本文をUTF-8でエンコード
    const bodyBytes = Utilities.newBlob('').setDataFromString(body, 'UTF-8').getBytes();
    parts.push(
      '--mixed_boundary',
      'Content-Type: text/plain; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      Utilities.base64Encode(bodyBytes)
    );

    // 添付ファイルの処理
    if (attachments && attachments.length > 0) {
      for (let [index, attachment] of attachments.entries()) {
        try {
          let attachmentBlob;
          let fileName;
          let contentType;

          if (typeof attachment === 'object' && attachment.path) {
            // 予約送信からの呼び出しの場合
            const fileId = attachment.path.replace(/.*\/d\/([^\/]+)\/.*/, '$1');
            const file = DriveApp.getFileById(fileId);
            attachmentBlob = file.getBlob();
            fileName = attachment.name;
            contentType = attachment.type;
          } else {
            // 即時送信からの呼び出しの場合
            attachmentBlob = attachment;
            fileName = attachment.getName();
            contentType = attachment.getContentType();
          }

          // ファイル名をRFC2047でエンコード
          const encodedFileName = '=?UTF-8?B?' + Utilities.base64Encode(
            Utilities.newBlob('').setDataFromString(fileName, 'UTF-8').getBytes()
          ) + '?=';

          parts.push(
            '--mixed_boundary',
            `Content-Type: ${contentType}; name="${encodedFileName}"`,
            'Content-Transfer-Encoding: base64',
            `Content-Disposition: attachment; filename="${encodedFileName}"`,
            `X-Attachment-Id: f_${sendingId}_${index}`,
            '',
            Utilities.base64Encode(attachmentBlob.getBytes())
          );

        } catch (error) {
          console.error(`添付ファイル処理エラー: ${error.message}`);
        }
      }
    }

    // 終端境界
    parts.push('--mixed_boundary--');

    // 最終的なメッセージの構築
    const message = headers.concat(parts).join('\r\n');

    // Gmail APIでメール送信
    Gmail.Users.Messages.send(
      {
        raw: Utilities.base64EncodeWebSafe(message)
      },
      'me'
    );

    console.log(`■メール送信完了 - ID: ${sendingId}, To: ${mailAddress}`);
    return { sendingId: sendingId, success: true };

  } catch (error) { // エラー時↓

    console.error(`メール送信エラー - ID: ${sendingId}, To: ${mailAddress}, エラー: ${error.message}`);
    
    if (error.message.includes('Invalid email')) {
      console.error('メールアドレスの形式が不正です');
    } else if (error.message.includes('Daily email quota exceeded')) {
      console.error('1日のメール送信制限を超えました');
    } else if (error.message.includes('Service invoked too many times')) {
      console.error('APIの呼び出し回数制限を超えました。しばらく待ってから再試行してください。');
    }
    
    return { sendingId: sendingId, success: false };
  }
}


/**
* 送信完了した原稿の内容をクリアする関数
* @param {number} startRow - クリア開始行
* @param {number} endRow - クリア終了行
*/
function clearYouzumiGenkou_(startRow, endRow) {
  try {
    const genkouSheet = SpreadsheetApp.openById(MANU_CONFIG_MAIL.SS_ID).getSheetByName(MANU_CONFIG_MAIL.GENKOU_SH);
    //送信者名の列は関数入ってるからクリアしない
    // Q列(COL_SENDER_NAME)の前までクリア
    const firstRange = genkouSheet.getRange(
      startRow, 
      columnToNumber1_('A'), 
      endRow - startRow + 1,
      columnToNumber1_(MANU_CONFIG_MAIL.COL_SENDER_NAME) - columnToNumber1_('A')
    );
    firstRange.clearContent();

    // Q列(COL_SENDER_NAME)の次の列から最後までクリア
    const secondRange = genkouSheet.getRange(
      startRow,
      columnToNumber1_(MANU_CONFIG_MAIL.COL_SENDER_NAME) + 1,
      endRow - startRow + 1,
      columnToNumber1_(MANU_CONFIG_MAIL.COL_BORDER) - columnToNumber1_(MANU_CONFIG_MAIL.COL_SENDER_NAME)
    );
    secondRange.clearContent();

    console.log(`行${startRow}から${endRow}までの原稿クリアが完了しました（送信者名は保持）`);
  } catch (error) {
    console.error(`原稿クリア処理中にエラーが発生: ${error.message}`);
    throw error;
  }
}


// 最後に公開インターフェースを定義
return {
    manuGmailMain: manuGmailMain,
    readStorageAndPassToMail_: readStorageAndPassToMail_
};

})();

/**
 * グローバルスコープ関数
 */
// スプシのボタンからメイン関数を動かしたいから、露出させとく必要ある
function manuGmailMain() {
    return ManualMailSender.manuGmailMain();
}

// 全部カプセル化されてるとトリガーが叩けない
function triggerPointForMailManu(e) {
    return ManualMailSender.readStorageAndPassToMail_(e);
}

