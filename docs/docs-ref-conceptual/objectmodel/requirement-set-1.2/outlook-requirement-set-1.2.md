# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook アドイン API 要件セット 1.2

Outlook アドイン API のサブセット Office 用の JavaScript API にはには、オブジェクト、メソッド、プロパティが含まれています、イベントが、Outlook で使用することができることを追加で。

> [!NOTE]
> このドキュメントではの[要件は、設定](/javascript/office/requirement-sets/outlook-api-requirement-sets)以外の最新の要件のセットです。 

## <a name="whats-new-in-12"></a>1.2 の新機能

要件セット 1.2 には、[要件セット 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) のすべての機能が含まれています。アドインを使用して、メッセージの件名または本文内のいずれかで、ユーザーのカーソル位置にテキストを挿入する機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)を追加: 件名またはメッセージの本文から選択したデータを非同期的に返します。
- [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) が追加されました。メッセージの本文または件名に非同期的にデータを挿入します。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) が変更されました。`formData` パラメーターに `attachments` プロパティが追加されました。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)