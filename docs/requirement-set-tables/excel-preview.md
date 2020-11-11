| クラス | フィールド | 説明 |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[プロバイダー](/javascript/api/excel/excel.linkeddatatype#dataprovider)|リンクされたデータ型のデータプロバイダーの名前を指定します。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|リンクされたデータ型が最後に更新されたときに、ブックが開かれてからのローカルタイムゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前を指定します。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|リンクされたデータ型が `refreshMode` "定期的" に設定されている場合に更新される頻度 (秒単位)。|
||[示し](/javascript/api/excel/excel.linkeddatatype#refreshmode)|リンクされたデータ型のデータを取得するメカニズムを指定します。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|リンクされたデータ型の一意の id。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|リンクされたデータ型によってサポートされるすべての更新モードを含む配列を返します。|
||[requestRefresh ()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode (refreshMode: LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新しいリンクされたデータ型の一意の id。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|リンクされたデータ型をサービス id で取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|コレクション内のインデックスによって、リンクされたデータ型を取得します。|
||[getItemOrNullObject (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|ID でリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシートビューをアクティブにします。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシートビューを削除します。|
||[重複 (名前?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシートビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シートビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定した名前の新しいシートビューを作成します。|
||[enterTemporary ()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シートビューを作成してアクティブにします。|
||[exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシートビューを終了します。|
||[getActive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシートビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシートビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシートビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスによってシートビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|ピボットテーブルの代替テキストタイトル。|
||[各アイテムを表示する (display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|各アイテムの後に空白行を表示するかどうかを設定します。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|ピボットテーブル内の空のセルに自動的に入力されるテキスト `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|ピボットテーブルの空のセルにを設定するかどうかを指定し `emptyCellText` ます。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルです。|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|ピボットテーブルのすべてのフィールドで [すべてのアイテムのラベルを繰り返す] 設定を設定します。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|ピボットテーブルにフィールドヘッダーを表示するかどうかを指定します (フィールドのタイトルとフィルターのドロップダウン)。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|ブックを開くときにピボットテーブルを更新するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents 元 ()](/javascript/api/excel/excel.range#getprecedents--)|`WorkbookRangeAreas`同じワークシートまたは複数のワークシート内のセルのすべての参照元を含む範囲を表すオブジェクト型 (object) の値を取得します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[示し](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|更新モードが変更されたオブジェクトの一意の id です。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が正常に終了したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|更新要求が完了したオブジェクトの一意の id。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[注意](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求によって生成された警告を含む配列。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|スライサーに適用されるスタイルです。|
||[setStyle (style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|表に適用されるスタイルです。|
||[setStyle (style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|表に適用するスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの id を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの id を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes 型](/javascript/api/excel/excel.workbook#linkeddatatypes)|ブックの一部である、リンクされたデータ型のコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールドリストウィンドウをブックレベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートにあるシートビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|
