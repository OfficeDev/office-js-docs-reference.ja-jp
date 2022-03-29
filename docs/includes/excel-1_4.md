| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|コレクションに含まれるバインドの数を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|ID を使用してバインド オブジェクトを取得します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|ワークシート上のグラフの数を返します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|グラフ名を使用してグラフを取得します。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|系列に含まれるグラフのポイントの数を返します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|コレクションに含まれるデータ系列の数を返します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|この名前に関連付けられたコメントを指定します。|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|指定された名前を削除します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|名前に関連付けられている範囲オブジェクトを返します。|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|名前をブックまたは特定のワークシートの範囲に設定する場合に指定します。|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|名前付きのアイテムの対象になるワークシートを返します。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|名前付きアイテムのスコープを設定するワークシートを返します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|指定のスコープのコレクションに新しい名前を追加します。|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|ユーザーのロケールを数式に使用して、指定のスコープのコレクションに新しい名前を追加します。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|コレクションに含まれる名前付きアイテムの数を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|その名前を `NamedItem` 使用してオブジェクトを取得します。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|コレクションに含まれるピボット テーブルの数を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|名前に基づいてピボットテーブルを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersectionornullobject-member(1))|指定した範囲の長方形の交差を表す範囲オブジェクトを取得します。|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getusedrangeornullobject-member(1))|指定した範囲オブジェクトのうち使用されている範囲を返します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|コレクション内のオブジェクトの `RangeView` 数を取得します。|
|[設定](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|設定を削除します。|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|設定の ID を表すキー。|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|この設定に格納されている値を表します。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array \| any)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|指定した設定をブックに設定または追加します。|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|コレクション内の設定の数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|キーを使用して設定エントリを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|キーを使用して設定エントリを取得します。|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|ドキュメントの設定が変更された場合に発生します。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|設定変更 `Setting` イベントを発生したバインドを表すオブジェクトを取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|コレクションに含まれるテーブルの数を取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|名前または ID でテーブルを取得します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|表の列数を取得します。|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|名前または ID によって、列オブジェクトを取得します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|表の行数を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|ブックに関連付けられている設定のコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getusedrangeornullobject-member(1))|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。|
||[名前](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|現在のワークシートにスコープされている名前のコレクション。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|コレクションに含まれるワークシートの数を取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|名前または ID を使用して、ワークシート オブジェクトを取得します。|
