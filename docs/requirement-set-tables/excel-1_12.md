| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|グラフ軸タイトルのテキストの向きを指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフ系列の 1 つのディメンションから値を取得します。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツ タイプを取得します。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|変更されたイベントのトリガー方法を表す変更の種類を取得します。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|コメントが追加された場合に発生します。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|コメント コレクション内のコメントまたは返信が変更された場合 (返信が削除される場合を含む) に発生します。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|コメント コレクション内のコメントが削除された場合に発生します。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|コメントの ID を表します。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|コメントに属する関連する返信の ID を表します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|返信のコンテンツ タイプ。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示する文化的に適切な形式を定義します。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付区切り記号として使用される文字列を取得します。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付値の書式文字列を取得します。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式文字列を取得します。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[コンパレータ](/javascript/api/excel/excel.pivotdatefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotdatefilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|フィルター条件の範囲の下限 `between` 。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|フィルター条件の範囲の上限 `between` 。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|`equals`、、 `before` `after` およびフィルター条件の場合は、比較を丸 1 日 `between` として行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|1 つ以上のフィールドの現在のピボットフィルターを設定し、フィールドに適用します。|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|フィールドのすべてのフィルターからすべての条件をクリアします。|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|指定した種類のフィールドのフィルターからすべての既存の条件をクリアします (現在適用されている場合)。|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|フィールドに現在適用されているフィルターを取得します。|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|フィールドに適用されたフィルターが何かあるか確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|PivotField の現在適用されている日付フィルター。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|PivotField の現在適用されているラベル フィルター。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|PivotField の現在適用されている手動フィルター。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|PivotField の現在適用されている値フィルター。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[コンパレータ](/javascript/api/excel/excel.pivotlabelfilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|フィルター条件の範囲の下限 `between` 。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|、、およびフィルター条件 `beginsWith` `endsWith` に使用される `contains` サブ文字列。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|フィルター条件の範囲の上限 `between` 。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理する選択したアイテムの一覧。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定のピボットフィールドに複数のピボットフィルターを適用できる場合を指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|コレクション内の最初のピボットテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[コンパレータ](/javascript/api/excel/excel.pivotvaluefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|フィルター条件の範囲の下限 `between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターが上位/下位の N 項目、上/下の N パーセント、または上/下の N 合計のフィルターの値を指定します。|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上/下のフィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" しきい値数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|フィルター条件の範囲の上限 `between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理するフィールドで選択した "value" の名前。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|同じワークシートまたは複数のワークシート内のセルのすべての直接の前例を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|範囲と重なるピボットテーブルのスコープ付きコレクションを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|セルが流出するアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存される場合を表します。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|このコレクション内の `RangeAreas` オブジェクトの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|コレクション内の `RangeAreas` 位置に基づいてオブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|コレクション内の `RangeAreas` ワークシート ID または名前に基づいてオブジェクトを返します。|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|コレクション内の `RangeAreas` ワークシート名または ID に基づいてオブジェクトを返します。|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|A1 スタイルのアドレスの配列を返します。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|オブジェクトを返 `RangeAreasCollection` します。|
||[範囲](/javascript/api/excel/excel.workbookrangeareas#ranges)|オブジェクト内のこのオブジェクトを構成する範囲を返 `RangeCollection` します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシート レベルのカスタム プロパティのコレクションを取得します。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|指定されたキーにマップする新しいカスタム プロパティを追加します。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
