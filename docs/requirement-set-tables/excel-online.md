| クラス | フィールド | 説明 |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|リンクされたブックを指すリンクを壊す要求を行います。|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|リンクされたブックを指す元の URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|リンクされたブックから取得したデータを更新する要求を行います。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|リンクされたブックへのすべてのリンクを壊します。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|リンクされたブックに関する情報を URL で取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|リンクされたブックに関する情報を URL で取得します。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|すべてのブック リンクを更新する要求を行います。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|ブック リンクの更新モードを表します。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|指定した名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|コレクション内のインデックスによってシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|テーブルから複数の行を削除します。|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|指定したインデックスから、指定した数の行をテーブルから削除します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|リンクされたブックのコレクションを返します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|ワークシートに存在するシート ビューのコレクションを返します。|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|ワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|ワークシートの表示設定が変更された場合に発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|ワークシートがブック内のユーザーによって移動された場合に発生します。|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|ワークシートコレクションでワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|ワークシート コレクションでワークシートの表示設定が変更された場合に発生します。|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|移動後のワークシートの新しい位置を取得します。|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|移動の前に、ワークシートの前の位置を取得します。|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|移動したワークシートの ID を取得します。|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|名前の変更後に、ワークシートの新しい名前を取得します。|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|名前が変更される前に、ワークシートの前の名前を取得します。|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|新しい名前を持つワークシートの ID を取得します。|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|イベントの種類を取得します。|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|表示設定の変更後に、ワークシートの新しい表示設定を取得します。|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|表示設定が変更される前に、ワークシートの以前の表示設定を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|表示が変更されたワークシートの ID を取得します。|
