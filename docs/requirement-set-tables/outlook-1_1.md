| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addFileAttachmentAsync (uri: string, attachmentName: string, options?: AsyncContextOptions & {isInline: boolean}, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|ファイルを添付ファイルとしてメッセージまたは予定に追加します。|
||[addItemAttachmentAsync (itemId: any, attachmentName: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。|
||[body](/javascript/api/outlook/outlook.appointmentcompose#body)|アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。|
||[isInline](/javascript/api/outlook/outlook.appointmentcompose#isinline)||
||[removeAttachmentAsync (attachmentId: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|メッセージまたは予定から添付ファイルを削除します。|
|[AppointmentForm](/javascript/api/outlook/outlook.appointmentform)|[body](/javascript/api/outlook/outlook.appointmentform#body)|アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[body](/javascript/api/outlook/outlook.appointmentread#body)|アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。|
|[AttachmentDetails](/javascript/api/outlook/outlook.attachmentdetails)|[attachmentType](/javascript/api/outlook/outlook.attachmentdetails#attachmenttype)|添付ファイルの種類を示す値を取得します。|
||[contentType](/javascript/api/outlook/outlook.attachmentdetails#contenttype)|添付ファイルの MIME コンテンツ タイプを取得します。|
||[id](/javascript/api/outlook/outlook.attachmentdetails#id)|添付ファイルの Exchange 添付ファイル ID を取得します。|
||[isInline](/javascript/api/outlook/outlook.attachmentdetails#isinline)|添付ファイルをアイテムの本文に表示するかどうかを示す値を取得します。|
||[name](/javascript/api/outlook/outlook.attachmentdetails#name)|添付ファイルの名前を取得します。|
||[size](/javascript/api/outlook/outlook.attachmentdetails#size)|添付ファイルのサイズをバイト単位で取得します。|
|[Body](/javascript/api/outlook/outlook.body)|[getTypeAsync (options?: AsyncContextOptions, callback?: (asyncResult<: CoercionType>) => void) ()](/javascript/api/outlook/outlook.body#gettypeasync-options--callback--asyncresult-)|コンテンツの形式が HTML とテキストのどちらであるかを示す値を取得します。|
||[prependAsync (data: string, options?: AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#prependasync-data--options--callback--asyncresult-)|アイテム本文の先頭に指定の内容を追加します。|
||[setSelectedDataAsync (data: string, options?: AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.body#setselecteddataasync-data--options--callback--asyncresult-)|本文の選択部分を、指定のテキストに置き換えます。|
|[Location](/javascript/api/outlook/outlook.location)|[getAsync (callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.location#getasync-callback--asyncresult-)|予定の場所を取得します。|
||[getAsync (options: AsyncContextOptions, callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.location#getasync-options--callback--asyncresult-)|予定の場所を取得します。|
||[setAsync (location: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.location#setasync-location--options--callback--asyncresult-)|予定の場所を設定します。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addFileAttachmentAsync (uri: string, attachmentName: string, options?: AsyncContextOptions & {isInline: boolean}, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|ファイルを添付ファイルとしてメッセージまたは予定に追加します。|
||[addItemAttachmentAsync (itemId: any, attachmentName: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.messagecompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。|
||[bcc](/javascript/api/outlook/outlook.messagecompose#bcc)|メッセージの **Bcc** (ブラインドカーボンコピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。|
||[body](/javascript/api/outlook/outlook.messagecompose#body)|アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。|
||[isInline](/javascript/api/outlook/outlook.messagecompose#isinline)||
||[removeAttachmentAsync (attachmentId: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|メッセージまたは予定から添付ファイルを削除します。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[body](/javascript/api/outlook/outlook.messageread#body)|アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。|
|[PhoneNumber](/javascript/api/outlook/outlook.phonenumber)|[type](/javascript/api/outlook/outlook.phonenumber#type)|電話番号の種類を識別する文字列を取得します。自宅、勤務先、モバイル、未確定。|
|[受信者](/javascript/api/outlook/outlook.recipients)|[addAsync (recipients: (string \| Emailuser.displayname \| emailaddressdetails) [], options?: AsyncContextOptions, callback?: (asyncresult: <void> ) => void)](/javascript/api/outlook/outlook.recipients#addasync-recipients-)|予定やメッセージの既存の受信者に、受信者のリストを追加します。|
||[getAsync (callback: (asyncResult: Office. AsyncResult<EmailAddressDetails [] >) => void)](/javascript/api/outlook/outlook.recipients#getasync-callback--asyncresult-)|予定やメッセージの受信者リストを取得します。|
||[getAsync (options: AsyncContextOptions, callback: (asyncResult: Office. AsyncResult<EmailAddressDetails [] >) => void)](/javascript/api/outlook/outlook.recipients#getasync-options--callback--asyncresult-)|予定やメッセージの受信者リストを取得します。|
||[setAsync (recipients: (string \| Emailuser.displayname \| emailaddressdetails) [], callback: (Asyncresult: Office. asyncresult <void> ) => void)](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|予定やメッセージの受信者リストを設定します。|
||[setAsync (recipients: (string \| Emailuser.displayname \| emailaddressdetails) [], Options: AsyncContextOptions, callback: (Asyncresult: Office. asyncresult <void> ) => void)](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|予定やメッセージの受信者リストを設定します。|
|[[件名]](/javascript/api/outlook/outlook.subject)|[getAsync (callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.subject#getasync-callback--asyncresult-)|予定またはメッセージの件名を取得します。|
||[getAsync (options: AsyncContextOptions, callback: (asyncResult: <string> ) => void)](/javascript/api/outlook/outlook.subject#getasync-options--callback--asyncresult-)|予定またはメッセージの件名を取得します。|
||[setAsync (subject: string, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.subject#setasync-subject--options--callback--asyncresult-)|予定またはメッセージの件名を設定します。|
|[Time](/javascript/api/outlook/outlook.time)|[getAsync (callback: (asyncResult: <Date> ) => void)](/javascript/api/outlook/outlook.time#getasync-callback--asyncresult-)|予定の開始または終了の時刻を取得します。|
||[getAsync (options: AsyncContextOptions, callback: (asyncResult: <Date> ) => void)](/javascript/api/outlook/outlook.time#getasync-options--callback--asyncresult-)|予定の開始または終了の時刻を取得します。|
||[setAsync (dateTime: Date, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.time#setasync-datetime--options--callback--asyncresult-)|予定の開始または終了の時刻を設定します。|
