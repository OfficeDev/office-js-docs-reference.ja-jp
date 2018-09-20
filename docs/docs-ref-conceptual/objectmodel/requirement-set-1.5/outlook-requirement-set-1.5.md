# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook アドイン API 要件セット 1.5

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントではの[要件は、設定](/javascript/office/requirement-sets/outlook-api-requirement-sets)以外の最新の要件のセットです。

## <a name="whats-new-in-15"></a>1.5 の新機能

要件セット 1.5 には、[要件セット 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) のすべての機能が含まれています。次の機能が追加されました。

- [ピン留め可能な作業ウィンドウ](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)のサポートが追加されました。
- [REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) の呼び出しのサポートが追加されました。
- インラインで添付ファイルにマークを付ける機能が追加されました。
- 作業ウィンドウまたはダイアログを閉じる機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。
- [Office.EventType](office.md#eventtype-string)を追加: イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれています。
- [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。
- [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--)を追加しました。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)