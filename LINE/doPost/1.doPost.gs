/**
 * Webhookからのリクエストを処理する関数
 * 
 * LINEプラットフォームからのイベントを受け取り、適切なハンドラーに振り分ける
 * 
 * @param {Object} e - Webhookイベントのデータ
 * @return {Object} 処理結果のJSONレスポンス
 */
function doPost(e) {
  const requestId = Utilities.getUuid(); // リクエストごとの一意のIDを生成
  GeneralUtils.logDebug_(`Webhook received (RequestID: ${requestId})`, e.postData.contents);

  // Webhookイベントの重複をチェック（60秒以内の同一イベントIDを無視）
  const eventData = JSON.parse(e.postData.contents);
  const webhookEventId = eventData.events[0].webhookEventId;
  const processedEventIds = CacheService.getScriptCache();

  if (processedEventIds.get(webhookEventId)) {
    GeneralUtils.logDebug_(`重複するWebhookイベントを無視 (RequestID: ${requestId}, EventID: ${webhookEventId})`);
    return ContentService.createTextOutput(JSON.stringify({ status: 'duplicate ignored' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  processedEventIds.put(webhookEventId, 'processed', 60);

  try {
    // イベントを種類別に振り分け
    const sortedEvents = distributeEvents_(e, requestId); // requestIdを渡す

    // メッセージイベントが含まれている場合、メッセージハンドラーで処理
    if (sortedEvents.message) {
      LineHandlers.MessageHandler.process_(sortedEvents.message, requestId); // requestIdを渡す
    }

    // 処理成功のレスポンスを返す
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // エラー発生時はログを記録し、エラーレスポンスを返す
    GeneralUtils.logError_(`doPost (RequestID: ${requestId})`, error);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
