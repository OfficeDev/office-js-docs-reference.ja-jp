# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このマニュアルでは、**プレビュー**が[要件を設定](/javascript/office/requirement-sets/outlook-api-requirement-sets)します。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。 この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。

プレビューの要件のセットには、すべての[要件の設定 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md)の機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

- [](/javascript/api/outlook/office.from)取得するメソッドを提供する新しいオブジェクトを追加、値からです。
- [開催者](/javascript/api/outlook/office.organizer)開催者の値を取得するメソッドを提供する新しいオブジェクトが追加されます。
- [定期的なアイテム](/javascript/api/outlook/office.recurrence)の取得し予定の定期的なパターンを設定するがのみ会議出席依頼には、メッセージの定期的なパターンを取得するメソッドを提供する新しいオブジェクトが追加されます。
- [SeriesTime](/javascript/api/outlook/office.seriestime) - を取得し、定期的な一連の日付と予定の時刻を設定しの一連の定期的な会議出席依頼の日時を取得するメソッドを提供する新しいオブジェクトを追加します。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options`。この値は、イベントの実行をキャンセルするために使用されます。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - は、添付ファイルの base64 エンコードをメッセージまたは予定を新しいメソッドを追加します。
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - は、サポートされているイベントのイベント ハンドラーを追加する新しいメソッドを追加します。
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - の変更を取得するのには、作成モードでの値からです。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されると渡される初期化データを返す新しい機能が追加されました。
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer)の作成モードで、開催者の値を取得するように変更します。
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - は、新しいプロパティを取得または設定、予定アイテムの定期的なパターンを管理するメソッドを提供するオブジェクトを追加します。 このプロパティを使用して、会議の定期的なパターンを取得することもできるアイテムを要求します。
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - は、イベント ハンドラーを削除する新しいメソッドを追加します。
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - 対象は、一連の出来事の id を取得する新しいプロパティが追加されます。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Microsoft Graph API の[アクセス トークンの取得](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - は、1 日の週または日の種類を指定する新しい列挙型を追加します。
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - は、月を指定する新しい列挙を追加します。
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - は、新しい定期的なアイテムに適用するタイム ゾーンを指定する列挙型を追加します。
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - は、新しい定期的なアイテムの種類を指定する列挙型を追加します。
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - は、月の週を指定する新しい列挙型を追加します。
- [Office.EventType](/javascript/api/office/office.eventtype) - の変更の追加によって、RecurrenceChanged、RecipientsChanged、AppointmentTimeChanged、および OfficeThemeChanged のイベントをサポートするために`RecurrencePatternChanged`、 `RecipientsChanged`、 `AppointmentTimeChanged`、および`OfficeThemeChanged`エントリそれぞれ。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)