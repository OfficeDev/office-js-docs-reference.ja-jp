| クラス | フィールド | 説明 |
|:---|:---|:---|
|[メールボックス](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync(eventType: Office.EventType \| 文字列、ハンドラー: any、callback?: (asyncResult: Office。AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[addHandlerAsync(eventType: Office.EventType \| 文字列、ハンドラー: any、オプション: Office。AsyncContextOptions、callback?: (asyncResult: Office。AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[getCallbackTokenAsync(options: Office.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: Office.AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|REST API または Web Services (EWS) の呼び出しに使用Exchange文字列を取得します。|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync(eventType: Office.EventType \| 文字列、callback?: (asyncResult: Office。AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[removeHandlerAsync(eventType: Office.EventType \| 文字列、オプション: Office。AsyncContextOptions、callback?: (asyncResult: Office。AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|この電子メール アカウントの REST エンドポイントの URL を取得します。|
