| クラス | フィールド | 説明 |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|リンクされたブックを指すリンクを壊す要求を行います。|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|リンクされたブックを指す元の URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|リンクされたブックから取得したデータを更新する要求を行います。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|リンクされたブックへのすべてのリンクを壊します。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|リンクされたブックに関する情報を URL で取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|リンクされたブックに関する情報を URL で取得します。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|すべてのブック リンクを更新する要求を行います。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|ブック リンクの更新モードを表します。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|指定した名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|コレクション内のインデックスによってシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|テーブルから複数の行を削除します。|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|指定したインデックスから、指定した数の行をテーブルから削除します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|リンクされたブックのコレクションを返します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|ワークシートに存在するシート ビューのコレクションを返します。|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|ワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|ワークシートの表示設定が変更された場合に発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|ワークシートがブック内で移動するときに発生します。|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|ワークシートコレクションでワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|ワークシート コレクションでワークシートの表示設定が変更された場合に発生します。|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|移動後のワークシートの新しい位置を取得します。|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|移動の前に、ワークシートの前の位置を取得します。|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|移動したワークシートの ID を取得します。|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|名前の変更後に、ワークシートの新しい名前を取得します。|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|名前が変更される前に、ワークシートの前の名前を取得します。|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|新しい名前を持つワークシートの ID を取得します。|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|イベントの種類を取得します。|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|表示設定の変更後に、ワークシートの新しい表示設定を取得します。|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|表示設定が変更される前に、ワークシートの以前の表示設定を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|表示が変更されたワークシートの ID を取得します。|
