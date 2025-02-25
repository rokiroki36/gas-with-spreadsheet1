/**
* LINEから受信したテキストメッセージの処理を振り分けるモジュール
* 主に以下の2つの処理に振り分け：
* 1. キーワードに基づく自動応答
* 2. 電話番号下4桁とLINEユーザーIDの紐付け
*/
const TextEventProcessors = {
   /**
    * テキストメッセージの振り分けを実行
    * 優先順位：
    * 1. キーワードマッチング
    * 2. 電話番号処理
    */
    processMessage_: async function(userInfo, event) {
      GeneralUtils.logDebug_('TextEventProcessors start', userInfo);

      // キーワードマッチングチェック
      const matchedKeyword = await this.shouldProcessAutoResponse_(userInfo); // 戻り値をmatchedKeywordに変更

      if (matchedKeyword) { // 条件をmatchedKeywordに変更
        GeneralUtils.logDebug_('キーワード一致: AutoResponseProcessorへ振り分け');
        try {
          userInfo.matchedKeyword = matchedKeyword; // userInfoにキーワードをセット
          const autoResponseResult = await AutoResponseProcessor.process_(userInfo);
          return {
            handled: true,
            processor: 'キーワード応答',
            result: autoResponseResult
          };
        } catch (error) {
          GeneralUtils.logError_('キーワード処理委譲時にエラー', error);
          return {
            handled: false,
            error: error.message
          };
        }
      }


        // 電話番号判定
        //shouldProcessPhoneNumberに回すと、数字以外削られるから、なにか条件付け足すならこの前の工程で
        const shouldProcessPhone = this.shouldProcessPhoneNumber_(userInfo);
        GeneralUtils.logDebug_('電話番号判定結果:', shouldProcessPhone);

        if (shouldProcessPhone) {
            try {
                GeneralUtils.logDebug_('PhoneNumberProcessor処理開始');
                const result = await PhoneNumberProcessor.mainPhoneProcess_(userInfo);
                GeneralUtils.logDebug_('PhoneNumberProcessor処理結果取得', result);

                if (result.handled) {
                    if (result.code === 'ALREADY_REGISTERED') {
                        GeneralUtils.logDebug_('既に登録済みユーザー: 処理終了');
                    }
                    return {
                        handled: true,
                        processor: '電話番号処理',
                        result: result
                    };
                }
            } catch (error) {
                GeneralUtils.logError_('電話番号処理委譲時にエラー', error);
                return {
                    handled: false,
                    error: error.message
                };
            }
        }

        // どちらの処理も行われなかった場合
        GeneralUtils.logDebug_('処理振り分け結果', {
            handled: false,
            reason: "キーワードにも電話番号形式にも該当せず"
        });
        return {
            handled: false,
            debugMessage: "振り分け対象となる条件にマッチしませんでした",
            source: event.source,
            type: event.type,
            message: userInfo.messageText,
            timestamp: event.timestamp
        };
    },


    // キーワードマッチングチェック
    shouldProcessAutoResponse_: async function(userInfo) {
      try {
        // 自動応答シートからキーワード列(B列)のみを取得
        const ss = SpreadsheetApp.openById(AutoResponseProcessor.CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(AutoResponseProcessor.CONFIG.SHEET_NAMES.AUTO_RESPONSE);
        const keywords = sheet.getRange(2, GeneralUtils.columnToNumber_(AutoResponseProcessor.CONFIG.COLUMN_NAMES.KEYWORD), sheet.getLastRow() - 1).getValues().flat();

        // ユーザーのメッセージがキーワードのいずれかを含むか判定
        const matchedKeyword = keywords.find(keyword => userInfo.messageText.includes(keyword));

        GeneralUtils.logDebug_('shouldProcessAutoResponse_ - キーワードマッチ判定結果', { matched: Boolean(matchedKeyword), keyword: matchedKeyword });

        // マッチしたキーワードを返す
        return matchedKeyword || null; // マッチしたキーワードがあればそれを返し、なければnullを返す
      } catch (error) {
        GeneralUtils.logError_('shouldProcessAutoResponse_ - エラー発生', error);
        return null; // エラー発生時はnullを返す
      }
    },
    
    // 電話番号ぽい数字があるかチェック
    shouldProcessPhoneNumber_: function(userInfo) {
    try {
        // 改行とハイフンを先に除去
        const preCleanedText = userInfo.messageText.replace(/[\r\n-]/g, '');

        // 数字の塊を区別し、_SEP_を挿入
        const separatedText = preCleanedText.replace(/(\d+)/g, "$1_SEP_");

        // 数字以外を全て除去 (ただし "_SEP_" は残す)
        const cleanedText = separatedText.replace(/[^\d_SEP_]/g, '');

        // 数字の塊ごとに処理する準備
        const numberGroups = cleanedText.split("_SEP_").filter(group => group.length > 0);

        const potentialPhoneNumbers = [];

        // 10桁以上か→0から始まるか、でチェック
        for (const group of numberGroups) {
            if (group.length >= 10) {
                if (group.startsWith('0')) {
                    potentialPhoneNumbers.push(group);
                }
            }
        }

        // 複数の電話番号候補が見つかった場合
        if (potentialPhoneNumbers.length > 1) {
            GeneralUtils.logError_('電話番号の可能性あるものが２つ以上あります: ' + potentialPhoneNumbers.join(', '));
            return false;
        }

        // 条件を満たす数字の塊が1つだけ存在する場合、userInfo.processedTextに格納→上のメインの電番判定に戻す
          if (potentialPhoneNumbers.length === 1) {
            userInfo.processedText = potentialPhoneNumbers[0];
            return true;
          }

        return false;

    } catch (error) {
        GeneralUtils.logError_('電話番号判定処理でエラー', error);
        return false;
    }
}
};



