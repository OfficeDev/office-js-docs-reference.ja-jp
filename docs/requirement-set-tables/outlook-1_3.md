| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[close ()](/javascript/api/outlook/outlook.appointmentcompose#close--)|構成されている現在のアイテムを閉じます。|
||[notificationMessages](/javascript/api/outlook/outlook.appointmentcompose#notificationmessages)|アイテムの通知メッセージを取得します。|
||[saveAsync (callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-callback--asyncresult-)|項目を非同期的に保存します。|
||[saveAsync (options: AsyncContextOptions, callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#saveasync-options--callback--asyncresult-)|項目を非同期的に保存します。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[notificationMessages](/javascript/api/outlook/outlook.appointmentread#notificationmessages)|アイテムの通知メッセージを取得します。|
|[Body](/javascript/api/outlook/outlook.body)|[getAsync (coercionType: CoercionType \| string, options?: AsyncContextOptions, callback?: (asyncresult: <string> ) => void)。](/javascript/api/outlook/outlook.body#getasync-coerciontype--options--callback--asyncresult-)|現在の本文を指定された形式で返します。|
||[setAsync (data: string, options?: AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setasync-data--options--callback--asyncresult-)|本文全体を指定されたテキストに置換します。|
|[メールボックス](/javascript/api/outlook/outlook.mailbox)|[convertToEwsId (itemId: string, Office.mailboxenums.restversion: MailboxEnums \| 文字列)](/javascript/api/outlook/outlook.mailbox#converttoewsid-itemid--restversion-)|REST 形式のアイテム ID を EWS 形式に変換します。|
||[convertToRestId (itemId: string, Office.mailboxenums.restversion: MailboxEnums \| 文字列)](/javascript/api/outlook/outlook.mailbox#converttorestid-itemid--restversion-)|EWS 形式のアイテム ID を REST 形式に変換します。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[close ()](/javascript/api/outlook/outlook.messagecompose#close--)|構成されている現在のアイテムを閉じます。|
||[notificationMessages](/javascript/api/outlook/outlook.messagecompose#notificationmessages)|アイテムの通知メッセージを取得します。|
||[saveAsync (callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#saveasync-callback--asyncresult-)|項目を非同期的に保存します。|
||[saveAsync (options: AsyncContextOptions, callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#saveasync-options--callback--asyncresult-)|項目を非同期的に保存します。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[notificationMessages](/javascript/api/outlook/outlook.messageread#notificationmessages)|アイテムの通知メッセージを取得します。|
|[NotificationMessageDetails](/javascript/api/outlook/outlook.notificationmessagedetails)|[icon](/javascript/api/outlook/outlook.notificationmessagedetails#icon)|`Resources`セクションのマニフェストで定義されているアイコンへの参照。|
||[key](/javascript/api/outlook/outlook.notificationmessagedetails#key)|通知メッセージの識別子。|
||[message](/javascript/api/outlook/outlook.notificationmessagedetails#message)|通知メッセージのテキスト。|
||[引き続き](/javascript/api/outlook/outlook.notificationmessagedetails#persistent)|メッセージを永続的にする必要があるかどうかを指定します。|
||[type](/javascript/api/outlook/outlook.notificationmessagedetails#type)|メッセージのを指定し `ItemNotificationMessageType` ます。|
|[NotificationMessages](/javascript/api/outlook/outlook.notificationmessages)|[addAsync (key: string, JSONmessage: NotificationMessageDetails, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#addasync-key--jsonmessage--options--callback--asyncresult-)|アイテムに通知を追加します。|
||[getAllAsync (options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult<NotificationMessageDetails [] >) => void)](/javascript/api/outlook/outlook.notificationmessages#getallasync-options--callback--asyncresult-)|アイテムのすべてのキーとメッセージを返します。|
||[removeAsync (key: string, options?: AsyncContextOptions, callback?: (asyncResult: <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#removeasync-key--options--callback--asyncresult-)|アイテムの通知メッセージを削除します。|
||[replaceAsync (key: string, JSONmessage: NotificationMessageDetails, options?: AsyncContextOptions, callback?: (asyncResult: <void> ) => void)](/javascript/api/outlook/outlook.notificationmessages#replaceasync-key--jsonmessage--options--callback--asyncresult-)|指定のキーが含まれる通知メッセージを別のメッセージに置換します。|
