| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|現在のシステムのカルチャ設定に基づく情報を提供します。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|数値の小数点の記号として使用される文字列を取得します。|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Excel のシステム区切り記号を有効にするかどうかを指定します。|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|コメントに記載されているエンティティ (ユーザーなど) を取得します。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|リッチコメントの内容 (コメント内のメンションなど) を取得します。|
||[解析](/javascript/api/excel/excel.comment#resolved)|コメントスレッドの状態。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress: Range \| string, content: CommentRichContent \| String, contenttype?: Excel)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|指定したセルで、指定した内容の新しいコメントを作成します。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|コメントに記載されているエンティティの電子メールアドレス。|
||[id](/javascript/api/excel/excel.commentmention#id)|エンティティの id。|
||[name](/javascript/api/excel/excel.commentmention#name)|Comment で言及されているエンティティの名前。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|コメントに記載されているエンティティ (ユーザーなど)。|
||[解析](/javascript/api/excel/excel.commentreply#resolved)|コメントの返信状態。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|リッチコメントの内容 (コメント内のメンションなど)。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add (content: CommentRichContent \| string, contenttype?: Excel)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|コメント内で言及されているすべてのエンティティ (人物など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|コメントのリッチコンテンツを指定します (たとえば、メンションを含むコメントコンテンツ、最初に説明したエンティティの id 属性は0、2番目に指定されたエンティティの id 属性は1です)。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|カルチャ名を languagecode2-country/regioncode2 の形式で取得します (例: "zh-cn-cn" または "en-us")。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|数字を表示するためのカルチャに適した形式を定義します。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|数値の小数点の記号として使用される文字列を取得します。|
||[番号 Groupseparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|セルの値、書式設定、および数式を現在の範囲から移動先の範囲に移動し、そのセルの古い情報を置き換えます。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[adjustIndent (金額: 数値)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|範囲の書式のインデントを調整します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|計算を完了した範囲のアドレス。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|イベントがトリガーされた方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
