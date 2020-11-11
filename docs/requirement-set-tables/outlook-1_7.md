| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[Addhandler Async (eventType: \| AsyncContextOptions, options?:, callback?: (asyncresult: Office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|指定した会議の開催者を取得します。|
||[繰り返さ](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|予定の定期的なパターンを取得または設定します。|
||[Removeハンドラ Async (eventType: \| AsyncContextOptions, callback?: (asyncresult: office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[系列 Id](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|インスタンスが属する系列の id を取得します。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[Addhandler Async (eventType: \| AsyncContextOptions, options?:, callback?: (asyncresult: Office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[繰り返さ](/javascript/api/outlook/outlook.appointmentread#recurrence)|予定の定期的なパターンを取得します。|
||[Removeハンドラ Async (eventType: \| AsyncContextOptions, callback?: (asyncresult: office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[系列 Id](/javascript/api/outlook/outlook.appointmentread#seriesid)|インスタンスが属する系列の ID を取得します。|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getAsync (オプション?: AsyncContextOptions, callback?: (asyncResult: <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|メッセージの from 値を取得します。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[Addhandler Async (eventType: \| AsyncContextOptions, options?:, callback?: (asyncresult: Office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|メッセージの送信者の電子メール アドレスを取得します。|
||[Removeハンドラ Async (eventType: \| AsyncContextOptions, callback?: (asyncresult: office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[系列 Id](/javascript/api/outlook/outlook.messagecompose#seriesid)|インスタンスが属する系列の ID を取得します。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[Addhandler Async (eventType: \| AsyncContextOptions, options?:, callback?: (asyncresult: Office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|サポートされているイベントのイベント ハンドラーを追加します。|
||[繰り返さ](/javascript/api/outlook/outlook.messageread#recurrence)|予定の定期的なパターンを取得します。|
||[Removeハンドラ Async (eventType: \| AsyncContextOptions, callback?: (asyncresult: office. asyncresult <void> ) => void) を指定します。](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|サポートされているイベントの種類のイベント ハンドラーを削除します。|
||[系列 Id](/javascript/api/outlook/outlook.messageread#seriesid)|インスタンスが属する系列の id を取得します。|
|[Organizer](/javascript/api/outlook/outlook.organizer)|[getAsync (オプション?: AsyncContextOptions, callback?: (asyncResult: <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|予定の開催者の値を {@link Office. EmailAddressDetails として取得します。 | EmailAddressDetails} オブジェクト|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[[編集後の受信者] フィールド](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|[ **Bcc** ] フィールド内の受信者が変更されたかどうかを取得します。|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|[ **Cc** ] フィールドの受信者が変更されたかどうかを取得します。|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|任意出席者が変更されたかどうかを取得します。|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|必要な出席者が変更されたかどうかを取得します。|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|リソースが変更されたかどうかを取得します。|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|[宛先] フィールド内の受信者が変更されたかどう **かを取得** します。|
|[Recurrence](/javascript/api/outlook/outlook.recurrence)|[getAsync (オプション?: AsyncContextOptions, callback?: (asyncResult: <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|定期的な予定の現在の繰り返しオブジェクトを返します。|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|一連の定期的な予定のプロパティを取得または設定します。|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|一連の定期的な予定のプロパティを取得または設定します。|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|一連の定期的な予定の種類を取得または設定します。|
||[seriesTime](/javascript/api/outlook/outlook.recurrence#seriestime)|{@Link SeriesTime | SeriesTime} オブジェクトを使用すると、定期的な予定の開始日と終了日を管理できます。|
||[setAsync (recurrencePattern: after, options?: AsyncContextOptions, callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|予定の定期的なパターンを設定します。|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[繰り返さ](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|月の日付を表します。|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|曜日または1日の種類を表します。たとえば、週末の曜日や曜日を表します。|
||[分](/javascript/api/outlook/outlook.recurrenceproperties#days)|この定期的なアイテムの日付のセットを表します。|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|選択されている最初の曜日を表します。それ以外の場合、既定値は現在のユーザーの設定の値です。|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|同じ定期的なアイテムのインスタンス間の期間を表します。|
||[month](/javascript/api/outlook/outlook.recurrenceproperties#month)|月を表します。|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|月の最初の週の場合は、選択した月の週の番号 (たとえば、"最初") を表します。|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|定期的なアイテムのタイムゾーンの名前を表します。|
||[交互](/javascript/api/outlook/outlook.recurrencetimezone#offset)|会議シリーズが開始された日付のローカルタイムゾーンと UTC との間の分単位の差を表す整数値。|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration ()](/javascript/api/outlook/outlook.seriestime#getduration--)|定期的な一連の予定に含まれる通常のインスタンスの期間 (分単位) を取得します。|
||[getEndDate ()](/javascript/api/outlook/outlook.seriestime#getenddate--)|次に示す定期的なパターンの終了日を取得します。|
||[getEndTime ()](/javascript/api/outlook/outlook.seriestime#getendtime--)|ユーザーまたはユーザーのいずれかのタイムゾーンで定期的な予定または定期的な予定の会議出席依頼のインスタンスの終了時刻を取得します。|
||[getStartDate ()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|次のアイテムの定期的なパターンの開始日を取得します。|
||[getStartTime ()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|ユーザー/アドインが設定するタイムゾーンのうち、定期的なパターンの通常の予定インスタンスの開始時刻を取得します。|
||[setDuration (分: 数値)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|定期的なパターンのすべての予定の期間を設定します。|
||[setEndDate (date: string)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|定期的な予定の系列の終了日を設定します。|
||[setEndDate (year: number, month: number, day: number)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|定期的な予定の系列の終了日を設定します。|
||[setStartDate (date: string)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|定期的な予定の系列の開始日を設定します。|
||[setStartDate (year: number、month: number、day: number)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|定期的な予定の系列の開始日を設定します。|
||[setStartTime (時間: 数値、分: 数値)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|定期的な予定の系列のすべてのインスタンスの開始時刻を、定期的なパターンが設定されている任意のタイムゾーンで設定します。|
||[setStartTime (time: string)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|定期的な予定の系列のすべてのインスタンスの開始時刻を、定期的なパターンが設定されている任意のタイムゾーンで設定します。|
