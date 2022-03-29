| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|現在のシステム カルチャ設定に基づく情報を提供します。|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|数値の小数点として使用される文字列を取得します。|
||[ThousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|数値の 10 進数の左側に数字のグループを区切る文字列を取得します。|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|ユーザーのシステム区切り記号が有効Excel指定します。|
|[コメント](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|コメントに記載されているエンティティ (人など) を取得します。|
||[解決済み](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|コメント スレッドの状態。|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|リッチ コメント コンテンツ (コメントのメンションなど) を取得します。|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|特別に書式設定された文字列とメンションの一覧を使用してコメント コンテンツを更新します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|指定したセルで、指定した内容の新しいコメントを作成します。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|コメントに記載されているエンティティの電子メール アドレス。|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|エンティティの ID。|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|コメントに記載されているエンティティの名前。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|コメントに記載されているエンティティ (人など)。|
||[解決済み](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|コメントの返信の状態。|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|豊富なコメント コンテンツ (コメント内のメンションなど)。|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|特別に書式設定された文字列とメンションの一覧を使用してコメント コンテンツを更新します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|コメントのコメント返信を作成します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|コメント内で言及されているすべてのエンティティ (人など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|コメントのリッチ コンテンツを指定します (例: メンション付きコメント コンテンツ、最初に言及したエンティティの ID 属性は 0、2 番目に指定したエンティティの ID 属性は 1)。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|languagecode2-country/regioncode2 形式のカルチャ名 ("zh-cn" や "ja-us" など) を取得します。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|数値を表示する文化的に適切な形式を定義します。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|数値の小数点として使用される文字列を取得します。|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|数値の 10 進数の左側に数字のグループを区切る文字列を取得します。|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|セルの値、書式設定、および数式を現在の範囲から移動先の範囲に移動し、それらのセルの古い情報を置き換える。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|範囲の書式設定のインデントを調整します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|現在のブックを保存します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|特定のワークシートで 1 つ以上の行の非表示状態が変更された場合に発生します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|計算が完了した範囲のアドレス。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|特定のワークシートで 1 つ以上の行の非表示状態が変更された場合に発生します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|イベントがトリガーされた方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|データが変更されたワークシートの ID を取得します。|
