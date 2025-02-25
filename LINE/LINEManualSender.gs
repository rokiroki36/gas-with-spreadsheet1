/**
 * =====================================================
 * LINE任意（Manual）送信Bot  
 *
 * スプシの原稿用シートの情報を元に、各宛先ユーザーIDに送る。
 * A列にnowってあれば即時送信
 * それ以外は予約送信、と条件分岐が自動でできる。
 * 送信ログシートは予約送信待ちのメッセのストレージも兼ねる。
 * 予約送信ならトリガー作って、指定時刻になったらトリガーが発火して、ストレージから情報引っ張って、LINE実送信関数に渡す
 * =====================================================
 */
const ManualSender = (function() {

const MANU_CONFIG = {

    DEBUG_MODE: false, // true: デバッグモード, false: 本番モード
    
    SS_ID: '1Burek5CFlxcN80ZtIAMLcvhgn62DzfNzzWq6V1CXXXX',
    MAIN_SH: '統一シート',
    GENKOU_SH: '任意送信原稿',
    LOG_SH: '予約ｽﾄﾚｰｼﾞ兼送信ﾛｸﾞ',

    LINE: {
        API_ENDPOINT: 'https://api.line.me/v2/bot/message/push',
        TOKEN_PROPERTY: 'LINE_ACCESS_TOKEN'
    },

    //任意送信原稿シート内の列の設定
    COL_CONTENTS: {
        CONTENT_1: 'E',
        CONTENT_2: 'G',
        CONTENT_3: 'I'
    },
    COL_ALT_TEXT: {
        ALT_TEXT_1: 'F',
        ALT_TEXT_2: 'H',
        ALT_TEXT_3: 'J'
    },

    COL_USERID: 'K',
    COL_BORDER: 'N', // 用済みの原稿をここまで消すよという列
    
    // 画像サイズ読み取り関数がバグった時に、ﾘｯﾁﾒｯｾの縦幅指定できるようにしとく
    COL_HEIGHT: {
    IMAGE_HEIGHT_1:'L',
    IMAGE_HEIGHT_2:'M',
    IMAGE_HEIGHT_3:'N',
    },

    // 統一シートの列設定
    MAIN_SH_COLUMNS: {
        USER_ID: 'DS', // LINEユーザーID
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
        USER_ID: 'G',
        SEI: 'H',
        MEI: 'I',
        MESSAGE_TYPE_1: 'J',
        MESSAGE_FILE_PATH_1: 'K',
        TEXT_PREVIEW_1: 'L',
        ALT_TEXT_1: 'M',
        MESSAGE_TYPE_2: 'N',
        MESSAGE_FILE_PATH_2: 'O',
        TEXT_PREVIEW_2: 'P',
        ALT_TEXT_2: 'Q',
        MESSAGE_TYPE_3: 'R',
        MESSAGE_FILE_PATH_3: 'S',
        TEXT_PREVIEW_3: 'T',
        ALT_TEXT_3: 'U',
        TIMESTAMP: 'V',
        TRIGGER_ID: 'W'
    },
    DRIVE_FOLDER_ID: '1m8DnseISuaVGqlJofTfKAF-QBVxzMEhT'

};

/**
 * =====================================================
 * メインエントリーポイント
 * =====================================================
 */
function manuLINEMain() {
    try {
        const groupRanges = getGroupRanges_();
        
        console.log("グループ範囲情報:");
        for (let i = 0; i < groupRanges.length; i++) {
            console.log(`Group ${i + 1}: start: ${groupRanges[i].start}, end: ${groupRanges[i].end}, keyRow: ${groupRanges[i].keyRow}`);
        }
        
        const preparedMessages = prepLINEGenkouFromSS_(groupRanges);
        return preparedMessages;
    } catch (error) {
        try {
            SpreadsheetApp.getUi().alert(
                'エラー',
                error.message,
                SpreadsheetApp.getUi().ButtonSet.OK
            );
        } catch (uiError) {
            console.error('LINE送信処理でエラーが発生:', error.message);
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
    const sheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.GENKOU_SH);

    // ユーザーIDの列番号を取得
    const userIdColumnNumber = columnToNumber1_(MANU_CONFIG.COL_USERID);

    // A列からユーザーID列までを一度に取得
    const data = sheet.getRange(1, 1, sheet.getLastRow(), userIdColumnNumber).getValues();

    // 最終行を判定（ユーザーID列のデータを使用）
    let lastRow = data.reduceRight((acc, row, idx) => {
        if (acc === 0 && row[userIdColumnNumber - 1] !== "") acc = idx + 1;
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
 * LINE 送信用のメッセージ原稿を準備する関数
 */
function prepLINEGenkouFromSS_(groupRanges) {
    const sheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.GENKOU_SH);

    // 各送信行為ごとに送る内容をまとめて、次の関数に渡すための変数
    const packedMaterialsWithKeyrow = [];

    // 引数として受け取った groupRanges を使用し、キー行を抽出
    // (将来的に部品関数化する際に、ここを変えれば他の関数と接続できるはず)
    const keyRowsToProcess = groupRanges.map(group => group.keyRow);

    // 選択したキー行を使ってメイン処理を行う
    for (const row of keyRowsToProcess) {
        let messageMaterials = [];

        // 送信するコンテンツの列数を数える
        let sendContentCount = 0;
        for (const key in MANU_CONFIG.COL_CONTENTS) {
            const column = MANU_CONFIG.COL_CONTENTS[key];
            const columnNum = columnToNumber1_(column);
            if (sheet.getRange(row, columnNum).getValue() !== "") {
                sendContentCount++;
            }
        }

        // コンテンツごとにメッセージを作成
        let contentIndex = 1;

        for (const contentKey in MANU_CONFIG.COL_CONTENTS) {
            if (contentIndex > sendContentCount) break;

            const contentColumn = MANU_CONFIG.COL_CONTENTS[contentKey];
            const contentColumnNum = columnToNumber1_(contentColumn);
            const contentCell = sheet.getRange(row, contentColumnNum);
            let messageMaterial;

            try {
                const richTextValue = contentCell.getRichTextValue();
                let targetUrl = '';
                let displayText = '';

                // リッチテキストからURLを取得
                if (richTextValue && richTextValue.getRuns().some(run => run.getLinkUrl())) {
                    const runs = richTextValue.getRuns();
                    const linkRun = runs.find(run => run.getLinkUrl());
                    targetUrl = linkRun.getLinkUrl();
                    displayText = linkRun.getText();
                } else {
                    targetUrl = contentCell.getDisplayValue();
                }

                if (targetUrl.includes('google.com/')) {
                    try {
                        const fileId = GeneralUtils.getFileIdFromDriveUrl_(targetUrl);
                        const file = DriveApp.getFileById(fileId);
                        const mimeType = file.getMimeType();

                        if (mimeType.includes('image')) {
                            file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
                            //公開用URL設定
                            const publicUrl = `https://drive.google.com/uc?id=${fileId}&export=download`;
                            try {
                                const dimensions = GeneralUtils.getImageDimensions_(targetUrl);
                                messageMaterial = prepImageMessage_(contentKey, publicUrl, row, dimensions.height, dimensions.width);
                            } catch (error) {
                                // 画像サイズ取得失敗時は臨時の高さ設定を使用
                                const heightColumn = MANU_CONFIG.COL_HEIGHT[`IMAGE_HEIGHT_${contentIndex}`];
                                const heightValue = heightColumn ? sheet.getRange(row, columnToNumber1_(heightColumn)).getValue() : null;
                                messageMaterial = prepImageMessage_(contentKey, publicUrl, row, heightValue);
                            }

                        } else if (mimeType === 'application/vnd.google-apps.document') {
                            const doc = DocumentApp.openById(fileId);
                            messageMaterial = {
                                message: { type: 'text', text: doc.getBody().getText() },
                                type: 'text'
                            };
                        } else {
                            const errorText = displayText ? 
                                `エラー：ドキュメントでも画像でもないファイルです リンクテキスト：${displayText} ファイル形式：${mimeType} URL：${targetUrl}` :
                                `エラー：ドキュメントでも画像でもないファイルです URL：${targetUrl} ファイル形式：${mimeType}`;
                            messageMaterial = { message: { type: 'text', text: errorText }, type: 'text' };
                        }
                    } catch (error) {
                        messageMaterial = {
                            message: { type: 'text', text: `エラー：Googleドライブファイルの処理に失敗しました URL：${targetUrl}` },
                            type: 'text'
                        };
                    }
                } else if (targetUrl.match(/\.(jpe?g|png|gif)$/i)) {
                    const heightColumn = MANU_CONFIG.COL_HEIGHT[`IMAGE_HEIGHT_${contentIndex}`];
                    const heightValue = heightColumn ? sheet.getRange(row, columnToNumber1_(heightColumn)).getValue() : null;
                    messageMaterial = prepImageMessage_(contentKey, targetUrl, row, heightValue);
                } else {
                    // 通常のテキストとして処理
                    messageMaterial = {
                        message: { type: 'text', text: contentCell.getDisplayValue() },
                        type: 'text'
                    };
                }

                messageMaterials.push(messageMaterial);
                contentIndex++;

            } catch (error) {
                console.error(`Error processing content at row ${row}, column ${contentColumn}: ${error.message}`);
                messageMaterials.push({
                    message: { type: 'text', text: `エラー：処理に失敗しました 詳細：${error.message}` },
                    type: 'text'
                });
                contentIndex++;
            }
        }

        // messageMaterials の内容をログに出力
        console.log(`Row ${row}:`);
        messageMaterials.forEach((material, index) => {
            if (material.message.type === 'flex') {
                console.log(
                    `  メッセージ ${index + 1}: Flexメッセージ - 代替テキスト: ${material.message.altText}, 画像URL: ${material.message.contents.hero.url}`
                );
            } else if (material.message.type === 'text') {
                console.log(
                    `  メッセージ ${index + 1}: テキストメッセージ - ${material.message.text.substring(0, 50)}${material.message.text.length > 50 ? '...' : ''}`
                );
            }
        });

        packedMaterialsWithKeyrow.push({
            keyRow: row,
            messageMaterials: messageMaterials
        });
    }

    return prepareMessagesForUsers_(groupRanges, packedMaterialsWithKeyrow);
}

/**
 * 画像メッセージを作成する関数
 * @param {string} contentKey - コンテンツのキー（CONTENT_1など）
 * @param {string} imageUrl - 画像のURL
 * @param {number} row - スプレッドシートの行番号
 * @param {number} height - 画像の高さ（オプション）
 * @param {number} width - 画像の幅（オプション）
 * @return {Object} LINE送信用の画像メッセージオブジェクト
 */
function prepImageMessage_(contentKey, imageUrl, row, height, width = 1040) {
    // デフォルト値の設定（widthもheightも引数でもらってればそっち優先）
    width = width || 1040;
    height = height || 1040;

    // アスペクト比の計算と調整
    let aspectRatio;
    const ratio = height / width;
    
    if (ratio > 3) {
        // 比率が3を超える場合、高さを調整
        const adjustedHeight = Math.floor(width * 3);
        aspectRatio = `${width}:${adjustedHeight}`;
    } else {
        aspectRatio = `${width}:${height}`;
    }

    // 代替テキストの取得
    const sheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.GENKOU_SH);
    const contentColumn = MANU_CONFIG.COL_CONTENTS[contentKey];
    let altText = "";

    // 代替テキストの取得 (MANU_CONFIGのペア情報を使用)
    for (const altTextKey in MANU_CONFIG.COL_ALT_TEXT) {
        if (MANU_CONFIG.COL_ALT_TEXT[altTextKey] === String.fromCharCode(contentColumn.charCodeAt(0) + 1)) {
            // 対応する代替テキスト列を検索
            const altTextColumn = MANU_CONFIG.COL_ALT_TEXT[altTextKey];
            const altTextColumnNum = columnToNumber1_(altTextColumn);
            altText = sheet.getRange(row, altTextColumnNum).getValue();
            break;
        }
    }

    // メッセージオブジェクトの作成
    return {
        message: {
            type: 'flex',
            altText: altText || "画像メッセージ",
            contents: {
                type: 'bubble',
                size: 'giga',
                hero: {
                    type: 'image',
                    url: imageUrl,
                    size: 'full',
                    aspectRatio: aspectRatio,
                    aspectMode: 'fit'
                }
            }
        },
        type: 'image'
    };
}

/**
 * 作った原稿をどこに送ればいいのか特定。特定できたらプレースホルダー置換
 */
function prepareMessagesForUsers_(groupRanges, packedMaterialsWithKeyrow) {

  console.log('=== 原稿と宛先の紐づけ開始 ===');

  const ss = SpreadsheetApp.openById(MANU_CONFIG.SS_ID);
  const genkouSheet = ss.getSheetByName(MANU_CONFIG.GENKOU_SH);
  const mainSheet = ss.getSheetByName(MANU_CONFIG.MAIN_SH);

  // groupRanges と packedMaterialsWithKeyrow を keyRow で結合
  const combinedData = groupRanges.map(group => {
    const correspondingZairyou = packedMaterialsWithKeyrow.find(item => item.keyRow === group.keyRow);
    return {
      keyRow: group.keyRow,
      sheetStRow: group.start, // 渡された配列は既にスプシに合わせて+2されてることを明示
      sheetEdRow: group.end,     // 同上
      messageMaterials: correspondingZairyou ? correspondingZairyou.messageMaterials : []
    };
  });

  // ユーザーデータマップを作成 (キー: ユーザーID, 値: {sei: 姓, mei: 名})
  // これは全ての送信行為で共通なので、ループの外で作成する
  const userDataMap = new Map();
  const lastRow = mainSheet.getLastRow();
  const userIdsFromMain = mainSheet.getRange(`${MANU_CONFIG.MAIN_SH_COLUMNS.USER_ID}2:${MANU_CONFIG.MAIN_SH_COLUMNS.USER_ID}${lastRow}`).getValues();
  const lastNames = mainSheet.getRange(`${MANU_CONFIG.MAIN_SH_COLUMNS.LAST_NAME}2:${MANU_CONFIG.MAIN_SH_COLUMNS.LAST_NAME}${lastRow}`).getValues();
  const firstNames = mainSheet.getRange(`${MANU_CONFIG.MAIN_SH_COLUMNS.FIRST_NAME}2:${MANU_CONFIG.MAIN_SH_COLUMNS.FIRST_NAME}${lastRow}`).getValues();

  for (let i = 0; i < userIdsFromMain.length; i++) {
    const userId = userIdsFromMain[i][0];
    if (userId) {
      const sei = lastNames[i][0] || '';
      const mei = firstNames[i][0] || '';
      userDataMap.set(userId, { sei, mei });
    }
  }

  // 各送信行為ごとに分けてループ処理
  const messageGroups = [];

  combinedData.forEach((group, groupIndex) => {
    let processedCount = 0;
    let skippedCount = 0;
    let duplicatedCount = 0; // 重複カウント用変数を追加

    console.log(`▼ グループ${groupIndex + 1} (${group.sheetStRow}行目～${group.sheetEdRow}行目)`);
    const keyRow = group.keyRow;

    // 宛先特定: GENKOU_SH シートの N 列から、sheetStRow 行目から sheetEdRow 行目までのユーザーIDを取得
    const userIdsRange = genkouSheet.getRange(group.sheetStRow, columnToNumber1_(MANU_CONFIG.COL_USERID),
      group.sheetEdRow - group.sheetStRow + 1, 1);
    const userIds = userIdsRange.getValues().flat().filter(id => id !== '');

    console.log(`送信先数: ${userIds.length}件`);

    // データマップとマッチング
    const matchedUserData = new Map();
    let matchedUserCount = 0;
    userIds.forEach(userId => {
      if (userDataMap.has(userId)) {
        matchedUserData.set(userId, userDataMap.get(userId));
        matchedUserCount++;
      }
    });

    console.log(`マッチしたユーザーデータ数: ${matchedUserCount}件`);

    // マッチしたユーザーID、姓、名を出力
    for (const [userId, userData] of matchedUserData) {
      console.log(`- ユーザーID: ${userId}, 姓: ${userData.sei}, 名: ${userData.mei}`);
    }

    //このあと各所でnow使うからループ外で宣言しとく  
    const now = new Date();
    const processedUserIds = new Set(); // 処理済みユーザーIDリスト

    // 各ユーザーごとにループ処理
    userIds.forEach((userId) => {
      // 重複チェック
      if (processedUserIds.has(userId)) {
        console.warn(`警告: ユーザーID ${userId} が重複しています。このユーザーIDはスキップされます。 - グループ: ${groupIndex + 1}, キー行: ${keyRow}`);
        duplicatedCount++;
        return; // 重複している場合はスキップ
      }

      if (userDataMap.has(userId)) {
        const userData = matchedUserData.get(userId);

        // 送信IDの生成
        const timestampStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
        const userIdSuffix = userId.slice(-8);
        const sendingId = `${timestampStr}_GRP${groupIndex + 1}_USR${userIdSuffix}`;

        // メッセージ処理
        const processedMessages = group.messageMaterials.map((material, materialIndex) => {

          const newMessage = { ...material.message };

          if (newMessage.type === 'text' && newMessage.text && newMessage.text.match(/\{sei\}|\{mei\}/g)) {
            newMessage.text = replacePlaceholders1_(newMessage.text, userData);
          } else if (newMessage.type === 'flex' && newMessage.altText && newMessage.altText.match(/\{sei\}|\{mei\}/g)) {
            newMessage.altText = replacePlaceholders1_(newMessage.altText, userData);
          }

          return { message: newMessage };
        });

        processedCount++;
        processedUserIds.add(userId); // 処理済みリストに追加

        // 送信セット情報を追加
        messageGroups.push({
          groupIndex: groupIndex,
          keyRow: keyRow,
          userId: userId,
          userData: userData,
          messages: processedMessages,
          timestamp: now,
          sendingId: sendingId
        });

        console.log(`    ユーザー ${userId} の処理が完了`);
      } else {
        console.error(`エラー: ユーザーID ${userId} が統一シートに見つかりません。このユーザーへの送信をスキップします。 - グループ: ${groupIndex + 1}, キー行: ${keyRow}`);
        skippedCount++;
      }
    });
    console.log(`グループ${groupIndex + 1}の処理完了: 処理済み: ${processedCount}件, スキップ: ${skippedCount}件, 重複: ${duplicatedCount}件`);
  });

  // nowOrDelay_にgroupedData（keyRowと1回の送信行為に何個メッセージ含まれるのか、と宛先数）と、messageGroups（宛先と紐づいた原稿一式）を渡す
  nowOrDelay_({
    groupedData: combinedData.map(group => ({
      keyRow: group.keyRow,
      messageCount: group.messageMaterials.length,
      targetUserCount: group.sheetEdRow - group.sheetStRow + 1
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
  const genkouSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID)
      .getSheetByName(MANU_CONFIG.GENKOU_SH);

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
 */
function nowOrDelay_(preparedData) {
  // 引数を解凍
  const { groupedData, messageGroups } = preparedData;

  // 各keyRowグループのキー行をチェックし、即時か予約か判定→ここで全keyRowグループ分、drop〜〜関数をループさせてる
  groupedData.forEach(group => {
    const keyRow = group.keyRow;
    const sendTypeCell = SpreadsheetApp.openById(MANU_CONFIG.SS_ID)
      .getSheetByName(MANU_CONFIG.GENKOU_SH)
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

      // 即時送信の場合は、グループ内の各ユーザーに対して、ここでループ処理を行う必要がある
      targetMessageGroups.forEach(group => {
        // LINE 送信用の関数を呼び出す
        const eachResult = sendMessagesToLINE1_(
          group.userId,
          group.messages,
          group.groupIndex,
          group.keyRow,
          group.sendingId
        );
        results.push(eachResult);

        if (!eachResult.success) { // 1つでも送信失敗したらフラグを false に
          allSucceeded = false;
        }
      });

      // まとめてログシートに書き込み
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
        // グループ範囲の一覧を取得
        const groupRanges = getGroupRanges_();
        
        // 処理したkeyRowに対応するグループを特定
        const targetGroup = groupRanges.find(group => group.keyRow === keyRow);
        
        if (targetGroup) {
          console.log(`原稿クリア開始: ${targetGroup.start}行目〜${targetGroup.end}行目`);
          clearYouzumiGenkou_(targetGroup.start, targetGroup.end);
          console.log(`原稿クリア完了: ${targetGroup.start}行目〜${targetGroup.end}行目`);
        } else {
          console.error(`キー行 ${keyRow} に対応するグループが見つかりませんでした`);
        }
      }

    } else {

      // 予約送信の場合
      try {
        const reservationTime = createReservationTime_(keyRow);
        console.log(`予約送信: keyRow ${keyRow}, 予約日時: ${reservationTime}`);

        // 予約時刻までスプシに完全な状態で保管
        dropToStorage_(reservationTime, keyRow, targetMessageGroups, {keyRow: group.keyRow,targetUserCount: targetMessageGroups.length}, null);

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
  console.log("▼ keyRow" ,keyRow,"のグループの予約ストレージ（送信ログ）書込み開始");
  console.log("予約日時:", nowOrReservationTime);
  console.log("グループデータ:", currentGroup);

  // 任意送信ログシートを取得
  const logSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.LOG_SH);

  // Googleドライブのフォルダを取得
  const folder = DriveApp.getFolderById(MANU_CONFIG.DRIVE_FOLDER_ID);

  // ログシートに書き込むデータ
  let logData = [];

  // 今から処理するグループの情報を取得
  if (!currentGroup) {
    console.error(`キー行 ${keyRow} に対応するグループデータが見つかりません`);
    return;
  }

  const groupIndex = targetMessageGroups[0].groupIndex + 1;
  const targetUserCount = currentGroup.targetUserCount;

  // 各keyRowグループ内に存在する各宛先に対しての処理
  targetMessageGroups.forEach(group => {
    if (!group || !group.messages) {
      console.error('Invalid group structure:', group);
      return;
    }

    const sendingId = group.sendingId;
    const userId = group.userId;

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

    // ログシートから実際の最後の列を取得し、その分の配列を初期化
    let logRow = Array(logSheet.getLastColumn()).fill('');

    // 各メッセージの処理とJSONファイル作成
    for (let i = 0; i < Math.min(group.messages.length, 3); i++) {
      const message = group.messages[i];
      let messageType = message.message.type;
      let messageFilePath = '';
      let textPreview = '';
      let altText = '';

      try {
        // JSONファイルとして保存
        const fileName = `message_${sendingId}_${i + 1}.json`;
        const file = folder.createFile(fileName, JSON.stringify(message.message), 'application/json');
        messageFilePath = file.getUrl();

        // ログシートに書く内容の下準備 
        if (messageType === 'text') { 
          textPreview = message.message.text.replace(/\r?\n/g, ' ').substring(0, 50);
        } 
        else if (messageType === 'flex') {
          altText = message.message.altText;
          if (message.message.contents?.hero?.url) { // 添付画像のリンクをログシートに置いて、すぐ確認できるようにする
            textPreview = message.message.contents.hero.url;
          }
        }

        // １〜３通目の内容をログシートに書く
        const cols = MANU_CONFIG.LOG_SH_COLUMNS;
        logRow[columnToNumber1_(cols[`MESSAGE_TYPE_${i + 1}`]) - 1] = messageType;
        logRow[columnToNumber1_(cols[`MESSAGE_FILE_PATH_${i + 1}`]) - 1] = messageFilePath;
        logRow[columnToNumber1_(cols[`TEXT_PREVIEW_${i + 1}`]) - 1] = textPreview;
        logRow[columnToNumber1_(cols[`ALT_TEXT_${i + 1}`]) - 1] = altText;

      } catch (error) {
        console.error(`メッセージ ${i + 1} の処理中にエラー: ${error.message}`);
      }
    }

    // メッセージ内容以外の項目
    const cols = MANU_CONFIG.LOG_SH_COLUMNS;
    logRow[columnToNumber1_(cols.SENDING_ID) - 1] = sendingId;
    logRow[columnToNumber1_(cols.KEY_ROW) - 1] = keyRow;
    logRow[columnToNumber1_(cols.GROUP_INDEX) - 1] = groupIndex;
    logRow[columnToNumber1_(cols.TARGET_USER_COUNT) - 1] = targetUserCount;
    logRow[columnToNumber1_(cols.RESERVATION_TIME) - 1] = formattedReservationTime;
    logRow[columnToNumber1_(cols.PROCESSED) - 1] = 0;
    logRow[columnToNumber1_(cols.USER_ID) - 1] = userId;
    logRow[columnToNumber1_(cols.SEI) - 1] = group.userData.sei;
    logRow[columnToNumber1_(cols.MEI) - 1] = group.userData.mei;
    logRow[columnToNumber1_(cols.TIMESTAMP) - 1] = timestamp;
    logRow[columnToNumber1_(cols.TRIGGER_ID) - 1] = '';

    // ログデータに追加
    logData.push(logRow);
  });

  // まとめて書き込み
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logData.length, logData[0].length).setValues(logData);

  console.log(`${logData.length}件のログデータを書き込みました`);

  // 予約送信の場合はトリガーを設定
  if (nowOrReservationTime !== 'NOW') {
    setTriggerForDelay_(nowOrReservationTime, keyRow, targetMessageGroups);
  }
}


/**
 * 予約送信用のトリガーを設定する関数
 * @param {string} reservationTime - 'MM/dd HH:mm' 形式の予約時刻
 * @param {number} keyRow - 任意送信原稿シートのキー行番号
 * @param {array} messageGroups - 送信グループごとのメッセージ情報を含む配列（修正箇所: 引数を追加）
 */
function setTriggerForDelay_(reservationTime, keyRow, messageGroups) {
  try {
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
    const trigger = ScriptApp.newTrigger('triggerPointForLINEManu')
      .timeBased()
      .at(triggerDate)
      .create();

    // トリガーIDをログシートに記録
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID)
      .getSheetByName(MANU_CONFIG.LOG_SH);
    
    // このkeyRowに紐づく現在処理中のすべての送信IDのベース部分を取得
    const currentSendingIdBases = messageGroups
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

    console.log(`トリガーを設定しました - 予約時刻: ${reservationTime}, キー行: ${keyRow}, トリガーID: ${trigger.getUniqueId()}`);

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
 * トリガーに叩かれて、ストレージのデータを引っ張ってLINE送信関数にわたす
 */
function readStorageAndPassToLINE1_(e) {
  try {
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.LOG_SH);

    // トリガーイベントから直接IDを取得
    const triggerId = e.triggerUid;
    console.log("発火したトリガーID:", triggerId);

    if (!triggerId) {
      console.warn('トリガーIDを取得できませんでした。');
      return;
    }

    // 処理対象の行を特定（同じトリガーIDで未送信のもの）
    const lastRow = logSheet.getLastRow();

    // 「送信ID」があるA列(indexとしては0)から「トリガーID」があるW列(indexとしては22)までを取得すればOK
    const data = logSheet.getRange(2, 1, lastRow - 1, columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.TRIGGER_ID)).getValues();
    const targetRows = [];

    data.forEach((row, index) => {
      const rowNum = index + 2;
      const processed = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.PROCESSED) - 1]; // F列の処理済みフラグの値を取得
      const currentTriggerId = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.TRIGGER_ID) - 1]; // W列のトリガーIDの値を取得
      const keyRow = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.KEY_ROW) - 1]; // B列のキーとなる行の値を取得
      const sendingId = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.SENDING_ID) -1]; // A列のsendingIDの値を取得
      const userId = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.USER_ID) - 1] // G列のuserIdの値を取得
      

      if (currentTriggerId === triggerId && !processed) {
        targetRows.push({
          rowNum: rowNum,
          sendingId: sendingId,
          userId: userId,
          keyRow: keyRow,
          messageData: [] // 1~3通のメッセージにまつわる各データ
        });

        // 1〜3通目のメッセージに関してループ処理
        for (let i = 0; i < 3; i++) {
          const messageType = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS[`MESSAGE_TYPE_${i + 1}`]) - 1];
          const filePath = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS[`MESSAGE_FILE_PATH_${i + 1}`]) - 1];
          const altText = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS[`ALT_TEXT_${i + 1}`]) - 1];
          
          // テキストは内容そのものがJSONに保存してあるけど
          // 画像の場合は、URLが保存されてるだけだから、再解凍みたいな作業が必要
          if (messageType) {  
            try {
              if (messageType === 'flex' && filePath) { 
             
                const fileId = filePath.replace(/.*\/d\/([^\/]+)\/.*/, '$1');
                const file = DriveApp.getFileById(fileId);
                const jsonString = file.getBlob().getDataAsString();

                const messageData = JSON.parse(jsonString);

                if (messageData.contents && messageData.contents.hero && messageData.contents.hero.url) {
                  const imageUrl = messageData.contents.hero.url;
                  targetRows[targetRows.length - 1].messageData.push({
                    type: messageType, //flex
                    filePath: filePath,
                    altText: altText,
                    imageUrl: imageUrl 
                  });
                } else {
                  console.error(`メッセージ ${i + 1}: 画像URL が見つかりません - sendingId: ${row[0]}`);
                }
              } 

              else { // textの場合はシンプルに内容取り出して格納するだけ
                targetRows[targetRows.length - 1].messageData.push({
                  type: messageType, //text
                  filePath: filePath,
                  altText: altText
                });
              }
            } catch (error) {
              console.error(`メッセージ ${i + 1}: エラー発生 - sendingId: ${row[0]}, エラー内容: ${error.message}, スタックトレース: ${error.stack}`);
              // エラー発生時は、そのメッセージはスキップして次のメッセージの処理を続ける。
            }
          }
        }
      }
    });

    console.log(`送信対象データ: ${targetRows.length}件`);

    // 各宛先に送信ループ開始！

    const results = []; // 送信結果を格納する配列

    for (const row of targetRows) {
      
        // F列を「処理中」の意味の「2」に更新
        logSheet.getRange(row.rowNum, columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.PROCESSED)).setValue(2);

        // LINE送信関数の呼び出し！
        const eachResult = sendMessagesToLINE1_(
          row.userId, 
          row.messageData, 
          null, 
          null, 
          row.sendingId
        );
        results.push(eachResult); // 各宛先の送信結果を蓄積
    } 

    // 全送信完了後に一括で処理済みフラグを更新
    flagOfSuccessOrFalse_(results);

    // 全送信完了後、発火元のトリガーを削除
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
    const logSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.LOG_SH);
    const lastRow = logSheet.getLastRow();

    // ログシートの全データを取得（A列から処理済みフラグのF列まで）
    const logData = logSheet.getRange(2, 1, lastRow - 1, columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.PROCESSED)).getValues();

    // 送信IDをキーとするマップを作成し、高速化を図る
    const sendingIdMap = new Map();
    results.forEach(result => {
      sendingIdMap.set(result.sendingId, result.success);
    });

    // ログシートのデータをループし、送信IDが一致する行の処理済みフラグを更新
    let updateCount = 0;
    logData.forEach((row, index) => {
      const sendingId = row[columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.SENDING_ID) - 1]; // A列の送信ID
      if (sendingIdMap.has(sendingId)) {
        const success = sendingIdMap.get(sendingId);
        const rowNum = index + 2; // 実際の行番号（見出し行を考慮）

        if (success) {
          logSheet.getRange(rowNum, columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.PROCESSED)).setValue(1); // 成功なら1
          updateCount++;
        } else {
          logSheet.getRange(rowNum, columnToNumber1_(MANU_CONFIG.LOG_SH_COLUMNS.PROCESSED)).setValue('エラー'); // 失敗なら「エラー」
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
 * LINE メッセージを実際に送信する関数
 * @param {string} userId - 送信先ユーザーの LINE ユーザーID
 * @param {object[]} messages - 送信するメッセージの配列（呼び出し元１はmessageオブジェクトをそのまま、呼び出し元２はtype, filePath, altTextのobjectを渡してくるので、場合分けが必要）
 * @param {number} groupIndex - 送信グループのインデックス（この関数では使わなくても、即時送信時はログシートに書き込むために、次の関数に渡す必要がある）
 * @param {number} keyRow - 任意送信原稿シートのキー行番号（同上）
 * @param {string} sendingId - 送信ID（オプション）
 * @return {boolean} 送信が成功した場合は true、失敗した場合は false を返す
 */
function sendMessagesToLINE1_(userId, messages, groupIndex, keyRow, sendingId) {
  const accessToken = PropertiesService.getScriptProperties().getProperty(MANU_CONFIG.LINE.TOKEN_PROPERTY);
  const url = MANU_CONFIG.LINE.API_ENDPOINT;

  // メッセージデータの配列を処理
  const lineMessages = messages.map((message, index) => {
    //呼び出し元１：nowOrDelay_() の場合
    if (message.message) {
        return message.message;

    //呼び出し元２：readStrageAndPassToLINE1() の場合
    } else if (message.filePath) {
      
      // message.type が text か flex 以外ならエラーを投げる
      if (message.type !== 'text' && message.type !== 'flex') {
          console.error(`メッセージ${index + 1}のタイプが不正です: 送信ID: ${sendingId}, エラー内容: 不明なメッセージタイプ ${message.type}`);
          return null;
      }

      try {
        const fileId = message.filePath.replace(/.*\/d\/([^\/]+)\/.*/, '$1');
        const file = DriveApp.getFileById(fileId);
        const messageData = JSON.parse(file.getBlob().getDataAsString());

        // タイプによって処理を分岐
        if (message.type === 'text') {
          return {
            type: 'text',
            text: messageData.text // textがnullやundefinedでも、エラーにしない
          };
        } else if (message.type === 'flex') {
          return {
            type: 'flex',
            altText: message.altText, // altTextがnullやundefinedでも、エラーにしない
            contents: messageData.contents // contentsがnullやundefinedでも、エラーにしない
          };
        }
      } catch (error) {
        console.error(`メッセージ${index + 1}の処理中にエラーが発生: 送信ID: ${sendingId}, エラー内容: ${error}`);

        if (error.message.includes("Cannot find file with ID")) {
          console.error("ファイルが見つかりません。ファイルが削除されたか、IDが間違っている可能性があります。");
        } else if (error.name === "SyntaxError") {
          console.error("メッセージデータのJSON形式が正しくありません。ドライブ上のファイルの該当箇所を確認してください");
        } else {
          console.error("メッセージデータの処理中に予期せぬエラーが発生しました。");
        }

        return null; //取得エラーが発生したら、nullを返して、そのメッセージはスキップ
      }
    } else {
      // 想定外のメッセージオブジェクトの場合
      console.error(`メッセージ${index + 1}が不正な形式です: 送信ID: ${sendingId}`);
      return null;
    }
  }).filter(message => message !== null); // null のメッセージを除外

  const payload = {
    'to': userId,
    'messages': lineMessages
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + accessToken,
    },
    'payload': JSON.stringify(payload)
  };

  try {
    // デバッグモードの場合は、メッセージ内容をログに出力して、送信処理をスキップ
    if (MANU_CONFIG.DEBUG_MODE) {
      console.log(`デバッグモード: メッセージ送信をスキップします`);
      console.log(`ユーザーID: ${userId}`);
      console.log(`送信ID: ${sendingId}`);
      console.log(`メッセージ内容:`);
      console.log(JSON.stringify(lineMessages, null, 2));
      return { sendingId: sendingId, success: true }; // 送信成功とみなす
    }

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      console.log(`メッセージ送信成功: ユーザーID: ${userId}, 送信ID: ${sendingId}, レスポンス: ${responseBody}`);
     return { sendingId: sendingId, success: true }; // 送信成功

    } else {
      console.error(`メッセージ送信失敗: ユーザーID: ${userId}, 送信ID: ${sendingId}, ステータスコード: ${responseCode}, レスポンス: ${responseBody}`);

      if (responseCode === 400) {
        console.error("リクエストエラーです。送信メッセージの内容が正しいか確認してください。");
      } else if (responseCode === 401) {
        console.error("認証エラーです。LINEアクセストークンが正しいか確認してください。");
      } else if (responseCode === 429) {
        console.error("レート制限エラーです。送信頻度を下げてください。");
      } else {
        console.error("予期せぬエラーが発生しました。");
      }

      return { sendingId: sendingId, success: false }; // 送信失敗
    }
  } catch (error) {
    console.error(`メッセージ送信中にエラーが発生しました: ユーザーID: ${userId}, 送信ID: ${sendingId}, エラー内容: ${error}`);
    
    if(error.message.includes("Address unavailable or invalid")){
        console.error("送信先ユーザーIDが間違っています");
    } else if (error.message.includes("timed out")) {
        console.error("LINE Messaging APIへのリクエストがタイムアウトしました");
    }
     return { sendingId: sendingId, success: false }; // 送信失敗
  }
}


/**
 * 送信完了した原稿の内容をクリアする関数
 * @param {number} startRow - クリア開始行
 * @param {number} endRow - クリア終了行
 */
function clearYouzumiGenkou_(startRow, endRow) {
  try {
    const genkouSheet = SpreadsheetApp.openById(MANU_CONFIG.SS_ID).getSheetByName(MANU_CONFIG.GENKOU_SH);
    
   const startColumn = 'A'; // クリア開始列
   const endColumn = MANU_CONFIG.COL_BORDER;
   const range = genkouSheet.getRange(startRow, columnToNumber1_(startColumn), endRow - startRow + 1, columnToNumber1_(endColumn) - columnToNumber1_(startColumn) + 1);
   range.clearContent();
    
    console.log(`行${startRow}から${endRow}までの原稿クリアが完了しました`);
  } catch (error) {
    console.error(`原稿クリア処理中にエラーが発生: ${error.message}`);
    throw error;
  }
}

// 最後に公開インターフェースを定義
return {
    manuLINEMain: manuLINEMain,
    readStorageAndPassToLINE1_: readStorageAndPassToLINE1_
};

})();


/**
 * グローバルスコープ関数
 */
// スプシのボタンからメイン関数を動かしたいから、露出させとく必要ある
function manuLINEMain() {
    return ManualSender.manuLINEMain();
}

// 全部カプセル化されてるとトリガーが叩けない
function triggerPointForLINEManu(e) {
    return ManualSender.readStorageAndPassToLINE1_(e);
}
