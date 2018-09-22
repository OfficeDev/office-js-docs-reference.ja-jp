# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook アドイン API 要件は、1.6 を設定します。

Outlook アドイン API のサブセット Office 用の JavaScript API にはには、オブジェクト、メソッド、プロパティが含まれています、イベントが、Outlook で使用することができることを追加で。

## <a name="whats-new-in-16"></a>1.6 の新機能は何ですか。

要件セット 1.6 には、すべての[要件は 1.5、設定](../requirement-set-1.5/outlook-requirement-set-1.5.md)の機能が含まれています。 それは、次の機能を追加します。

- 追加された Api を使用するアドインのコンテキスト、エンティティを取得するか、正規表現は、アドインをアクティブにするのには、ユーザーが選択されていると一致します。
- 新しいメッセージ フォームを開くに新しい API を追加します。
- アドインをユーザーのメールボックスのアカウントの種類を決定するための機能を追加します。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities)を追加: については、ユーザーが選択されて強調表示された一致するエンティティを取得する新しい関数を追加します。 強調表示された一致は、コンテキスト アドインに適用されます。
- [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)を追加: で強調表示されている一致するマニフェストの XML ファイルで定義されている正規表現に一致する文字列の値を返す新しい関数を追加します。 強調表示された一致は、コンテキスト アドインに適用されます。
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)を追加: 新しいメッセージ フォームを表示する新しい関数を追加します。
- [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)を追加する: ユーザーのアカウントの種類を示すユーザー プロファイルに新しいメンバーを追加します。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [作業の開始](https://docs.microsoft.com/outlook/add-ins/quick-start)