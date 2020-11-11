| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|バインドを削除します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range: Range \| string, bindingtype: Excel. bindingtype, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|特定の範囲に新しいバインドを追加します。|
||[addFromNamedItem (name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromSelection (bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|現在の選択範囲に基づいて新しいバインドを追加します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|現在のピボットテーブルを含んでいるワークシート。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|ピボットテーブルを更新します。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|コレクション内のすべてのピボットテーブルを更新します。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|現在の範囲の表示されている行を表します。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|R1C1 スタイル表記の数式を表します。|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|現在の RangeView に関連付けられている親の範囲を取得します。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|指定したセルの Excel の数値書式コードを表します。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|RangeView のセル アドレスを表します。|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|表示される列の数を指定します。|
||[index](/javascript/api/excel/excel.rangeview#index)|RangeView のインデックスを表す値を返します。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|表示される行の数を指定します。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|範囲に関連付けられている範囲ビューのコレクションを表します。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|各セルのデータの種類を表します。|
||[values](/javascript/api/excel/excel.rangeview#values)|指定した範囲ビューの Raw 値を表します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|インデックスを使用して、RangeView 行を取得します。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを指定します。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを指定します。|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|列が、表を見やすくするために、奇数列の強調表示と異なる方法で表示される縞模様の書式を表示するかどうかを指定します。|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|表を見やすくするために、奇数行の強調表示に使用する縞模様の書式を行に表示するかどうかを指定します。|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|フィルターボタンを各列ヘッダーの上部に表示するかどうかを指定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[テーブル](/javascript/api/excel/excel.workbook#pivottables)|ブックに関連付けられているピボットテーブルのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[テーブル](/javascript/api/excel/excel.worksheet#pivottables)|ワークシートの一部になっているピボットテーブルのコレクション。|
