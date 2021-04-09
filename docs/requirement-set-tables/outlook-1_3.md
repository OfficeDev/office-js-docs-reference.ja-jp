| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[close()](/javascript/api/outlook/outlook.appointmentcompose#close--)|構成されている現在のアイテムを閉じます。|
||[notificationMessages](/javascript/api/outlook/outlook.appointmentcompose#notificationmessages)|アイテムの通知メッセージを取得します。|
||[saveAsync(callback: (asyncResult: Office.AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-callback--asyncresult-)|項目を非同期的に保存します。|
||[saveAsync(options: Office.AsyncContextOptions, callback: (asyncResult: Office.AsyncResult ) => <string> void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-options--callback--asyncresult-)|項目を非同期的に保存します。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[notificationMessages](/javascript/api/outlook/outlook.appointmentread#notificationmessages)|アイテムの通知メッセージを取得します。|
|[Body](/javascript/api/outlook/outlook.body)|[getAsync(coercionType: Office.CoercionType \| string, callback?: (asyncResult: Office.AsyncResult ) => <string> void)](/javascript/api/outlook/outlook.body#getasync-coerciontype--callback--asyncresult-)|現在の本文を指定された形式で返します。|
||[getAsync(coercionType: Office.CoercionType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.body#getasync-coerciontype--options--callback--asyncresult-)|現在の本文を指定された形式で返します。|
||[setAsync(data: string, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setasync-data--callback--asyncresult-)|本文全体を指定されたテキストに置換します。|
||[setAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setasync-data--options--callback--asyncresult-)|本文全体を指定されたテキストに置換します。|
|[メールボックス](/javascript/api/outlook/outlook.mailbox)|[convertToEwsId(itemId: string, restVersion: MailboxEnums.RestVersion \| string)](/javascript/api/outlook/outlook.mailbox#converttoewsid-itemid--restversion-)|REST 形式のアイテム ID を EWS 形式に変換します。|
||[convertToRestId(itemId: string, restVersion: MailboxEnums.RestVersion \| string)](/javascript/api/outlook/outlook.mailbox#converttorestid-itemid--restversion-)|EWS 形式のアイテム ID を REST 形式に変換します。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[close()](/javascript/api/outlook/outlook.messagecompose#close--)|構成されている現在のアイテムを閉じます。|
||[notificationMessages](/javascript/api/outlook/outlook.messagecompose#notificationmessages)|アイテムの通知メッセージを取得します。|
||[saveAsync(callback: (asyncResult: Office.AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#saveasync-callback--asyncresult-)|項目を非同期的に保存します。|
||[saveAsync(options: Office.AsyncContextOptions, callback: (asyncResult: Office.AsyncResult ) => <string> void)](/javascript/api/outlook/outlook.messagecompose#saveasync-options--callback--asyncresult-)|項目を非同期的に保存します。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[notificationMessages](/javascript/api/outlook/outlook.messageread#notificationmessages)|アイテムの通知メッセージを取得します。|
|[NotificationMessageDetails](/javascript/api/outlook/outlook.notificationmessagedetails)|[icon](/javascript/api/outlook/outlook.notificationmessagedetails#icon)|`Resources`セクションのマニフェストで定義されているアイコンへの参照。|
||[key](/javascript/api/outlook/outlook.notificationmessagedetails#key)|通知メッセージの識別子。|
||[message](/javascript/api/outlook/outlook.notificationmessagedetails#message)|通知メッセージのテキスト。|
||[persistent](/javascript/api/outlook/outlook.notificationmessagedetails#persistent)|メッセージを永続的に設定する必要がある場合に指定します。|
||[type](/javascript/api/outlook/outlook.notificationmessagedetails#type)|メッセージの数 `ItemNotificationMessageType` を指定します。|
|[NotificationMessages](/javascript/api/outlook/outlook.notificationmessages)|[addAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#addasync-key--jsonmessage--callback--asyncresult-)|アイテムに通知を追加します。|
||[addAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#addasync-key--jsonmessage--options--callback--asyncresult-)|アイテムに通知を追加します。|
||[getAllAsync(callback?: (asyncResult: Office.AsyncResult<NotificationMessageDetails[]>) => void)](/javascript/api/outlook/outlook.notificationmessages#getallasync-callback--asyncresult-)|アイテムのすべてのキーとメッセージを返します。|
||[getAllAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<NotificationMessageDetails[]>) => void)](/javascript/api/outlook/outlook.notificationmessages#getallasync-options--callback--asyncresult-)|アイテムのすべてのキーとメッセージを返します。|
||[removeAsync(key: string, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#removeasync-key--callback--asyncresult-)|アイテムの通知メッセージを削除します。|
||[removeAsync(key: string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#removeasync-key--options--callback--asyncresult-)|アイテムの通知メッセージを削除します。|
||[replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#replaceasync-key--jsonmessage--callback--asyncresult-)|指定のキーが含まれる通知メッセージを別のメッセージに置換します。|
||[replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#replaceasync-key--jsonmessage--options--callback--asyncresult-)|指定のキーが含まれる通知メッセージを別のメッセージに置換します。|
