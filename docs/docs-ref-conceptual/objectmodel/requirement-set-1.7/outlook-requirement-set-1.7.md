# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook アドイン API 要件は、1.7 を設定します。

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

## <a name="whats-new-in-17"></a>1.7 の新機能は何ですか。

要件セット 1.7 には、すべての[要件の設定 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md)の機能が含まれています。 それは、次の機能を追加します。

- 定期的な予定と会議出席依頼には、メッセージに関する新しい Api を追加します。
- 使用する作成モードで item.from プロパティを変更するには。
- RecurrenceChanged、RecipientsChanged、AppointmentTimeChanged イベントのサポートを追加します。

### <a name="change-log"></a>変更ログ

- [](/javascript/api/outlook_1_7/office.from)追加: を取得するメソッドを提供する新しいオブジェクトを追加、値からです。
- [開催者](/javascript/api/outlook_1_7/office.organizer)の追加: 開催者の値を取得するメソッドを提供する新しいオブジェクトを追加します。
- [定期的なアイテム](/javascript/api/outlook_1_7/office.recurrence)を追加します。 メソッドを取得すると予定の定期的なパターンを設定のみ取得するメッセージの定期的なパターンは、会議出席依頼を提供する新しいオブジェクトを追加します。
- [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone)を追加します。 定期的なパターンのタイム ゾーンの構成を表す新しいオブジェクトを追加します。
- [SeriesTime](/javascript/api/outlook_1_7/office.seriestime)を追加: を取得し、定期的な一連の日付と予定の時刻を設定しの一連の定期的な会議出席依頼の日時を取得するメソッドを提供する新しいオブジェクトを追加します。
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback)を追加する: サポートされているイベントのイベント ハンドラーを追加する新しいメソッドを追加します。
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)の変更: 変更を取得するのには、作成モードでの値からです。
- 作成モードで、開催者の値を取得する[Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer)を変更したを変更します。
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence)を追加します。 新しいプロパティを取得または設定、予定アイテムの定期的なパターンを管理するメソッドを提供するオブジェクトを追加します。 このプロパティを使用して、会議の定期的なパターンを取得することもできるアイテムを要求します。
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback)を追加: イベント ハンドラーを削除する新しいメソッドを追加します。
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string)を追加: する回の定期的なアイテムの id を取得する新しいプロパティが属しているを追加します。
- [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days)を追加: 新しい 1 日の週または日の種類を指定する列挙型を追加します。
- [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month)を追加: 新しい月を指定する列挙型を追加します。
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone)を追加: 新しい定期的なアイテムに適用するタイム ゾーンを指定する列挙型を追加します。
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype)を追加: 新しい定期的なアイテムの種類を指定する列挙型を追加します。
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber)を追加: 新しい月の週を指定する列挙型を追加します。
- [Office.EventType](/javascript/api/office/office.eventtype)を変更: 変更の追加によって、RecurrenceChanged、RecipientsChanged、および AppointmentTimeChanged のイベントをサポートするために`RecurrenceChanged`、`RecipientsChanged`と`AppointmentTimeChanged`エントリそれぞれ。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)