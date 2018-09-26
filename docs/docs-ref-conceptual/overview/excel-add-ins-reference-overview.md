# <a name="excel-javascript-api-overview"></a>Excel の JavaScript API の概要

アドイン Excel 2016 またはそれ以降のビルドには、Excel の JavaScript API を使用できます。 API で使用できる Excel オブジェクトの概要を次に示します。 各オブジェクトのページのリンクには、プロパティ、イベント、およびオブジェクトの使用可能な方法の説明が含まれています。 メニューからのリンクを調べて、詳細を確認してください。

便宜上、Excel の主要なオブジェクトの一部を以下に示します。 

- [ブック](/javascript/api/excel/excel.workbook): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。

- [Worksheet](/javascript/api/excel/excel.worksheet):ブック内のワークシートを表します。 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection):ブック内の **Worksheet** オブジェクトのコレクション。

- [Range](/javascript/api/excel/excel.range):1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。

- [Table](/javascript/api/excel/excel.table):データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。
    - [TableCollection](/javascript/api/excel/excel.tablecollection):ブックまたはワークシート内のテーブルのコレクション。
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection):テーブル内のすべての列のコレクション。
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection):テーブル内のすべての行のコレクション。

- [Chart](/javascript/api/excel/excel.chart):基になるデータを視覚的に表示する、ワークシート内の chart オブジェクトを表します。
    - [ChartCollection](/javascript/api/excel/excel.chartcollection):ワークシート内のグラフのコレクション。

- [TableSort](/javascript/api/excel/excel.tablesort):**Table** オブジェクトの並べ替え操作を管理するオブジェクトを表します。

- [RangeSort](/javascript/api/excel/excel.rangesort):**Range** オブジェクトの並べ替え操作を管理するオブジェクトを表します。

- [Filter](/javascript/api/excel/excel.filter):テーブルの列のフィルター処理を管理するオブジェクトを表します。

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection):**Worksheet** オブジェクトの保護を表します。

- [NamedItem](/javascript/api/excel/excel.nameditem):セルまたは値の範囲の定義済みの名前を表します。 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection):ブック内の **NamedItem** オブジェクトのコレクション。

- [Binding](/javascript/api/excel/excel.binding):ブックのセクションへのバインドを表す抽象クラス。
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection):ブック内の **Binding** オブジェクトのコレクション。

## <a name="excel-javascript-api-open-specifications"></a>Excel の JavaScript API 仕様を開く

ようにデザインし、Api を使用する Excel のアドインを開発、私たちがで使用できるようにフィードバックを[オープン API の仕様](../openspec.md)のページ。 新機能が Excel JavaScript Api では、パイプラインし、設計仕様書に必要な設定を入力するどのようなを確認します。

## <a name="excel-javascript-api-reference"></a>Excel の JavaScript API リファレンス

Excel の JavaScript API の詳細については、 [Excel の JavaScript API リファレンス ドキュメント](/javascript/api/excel)を参照してください。

## <a name="see-also"></a>関連項目

- [Excel アドインの概要](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office アドイン プラットフォームの概要](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [GitHub 上の Excel のアドインのサンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
