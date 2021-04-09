| クラス | フィールド | 説明 |
|:---|:---|:---|
|[メールボックス](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync(eventType: \| Office.EventType string, handler: (type: Office.EventType) => void, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[addHandlerAsync(eventType: \| Office.EventType string, handler: (type: Office.EventType) => void, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[getCallbackTokenAsync(options: Office.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: Office.AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|REST API または Exchange Web Services (EWS) の呼び出しに使用されるトークンを含む文字列を取得します。|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult ) => <void> void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[removeHandlerAsync(eventType: \| Office.EventType string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|この電子メール アカウントの REST エンドポイントの URL を取得します。|
