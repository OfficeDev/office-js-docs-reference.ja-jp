# <a name="excel-javascript-api-overview"></a><span data-ttu-id="e7afd-101">Excel の JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="e7afd-101">Excel JavaScript API overview</span></span>

<span data-ttu-id="e7afd-102">アドイン Excel 2016 またはそれ以降のビルドには、Excel の JavaScript API を使用できます。</span><span class="sxs-lookup"><span data-stu-id="e7afd-102">You can use the Excel JavaScript API to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="e7afd-103">API で使用できる Excel オブジェクトの概要を次に示します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-103">The following list shows the high-level Excel objects that are available in the API.</span></span> <span data-ttu-id="e7afd-104">各オブジェクトのページのリンクには、プロパティ、イベント、およびオブジェクトの使用可能な方法の説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e7afd-104">Each object page link contains a description of the properties, events, and methods that are available on the object.</span></span> <span data-ttu-id="e7afd-105">メニューからのリンクを調べて、詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="e7afd-105">Explore the links from the menu to learn more.</span></span>

<span data-ttu-id="e7afd-106">便宜上、Excel の主要なオブジェクトの一部を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-106">Some of the core Excel objects are listed below for convenience:</span></span> 

- <span data-ttu-id="e7afd-107">[ブック](/javascript/api/excel/excel.workbook): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。</span><span class="sxs-lookup"><span data-stu-id="e7afd-107">[Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.</span></span>

- <span data-ttu-id="e7afd-108">[Worksheet](/javascript/api/excel/excel.worksheet):ブック内のワークシートを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-108">[Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook.</span></span> 
    - <span data-ttu-id="e7afd-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection):ブック内の **Worksheet** オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.</span></span>

- <span data-ttu-id="e7afd-110">[Range](/javascript/api/excel/excel.range):1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-110">[Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.</span></span>

- <span data-ttu-id="e7afd-111">[Table](/javascript/api/excel/excel.table):データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-111">[Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.</span></span>
    - <span data-ttu-id="e7afd-112">[TableCollection](/javascript/api/excel/excel.tablecollection):ブックまたはワークシート内のテーブルのコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-112">[TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.</span></span>
    - <span data-ttu-id="e7afd-113">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection):テーブル内のすべての列のコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-113">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.</span></span>
    - <span data-ttu-id="e7afd-114">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection):テーブル内のすべての行のコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-114">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.</span></span>

- <span data-ttu-id="e7afd-115">[Chart](/javascript/api/excel/excel.chart):基になるデータを視覚的に表示する、ワークシート内の chart オブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-115">[Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.</span></span>
    - <span data-ttu-id="e7afd-116">[ChartCollection](/javascript/api/excel/excel.chartcollection):ワークシート内のグラフのコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-116">[ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.</span></span>

- <span data-ttu-id="e7afd-117">[TableSort](/javascript/api/excel/excel.tablesort):**Table** オブジェクトの並べ替え操作を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-117">[TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on **Table** objects.</span></span>

- <span data-ttu-id="e7afd-118">[RangeSort](/javascript/api/excel/excel.rangesort):**Range** オブジェクトの並べ替え操作を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-118">[RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on **Range** objects.</span></span>

- <span data-ttu-id="e7afd-119">[Filter](/javascript/api/excel/excel.filter):テーブルの列のフィルター処理を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-119">[Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.</span></span>

- <span data-ttu-id="e7afd-120">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection):**Worksheet** オブジェクトの保護を表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-120">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.</span></span>

- <span data-ttu-id="e7afd-121">[NamedItem](/javascript/api/excel/excel.nameditem):セルまたは値の範囲の定義済みの名前を表します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-121">[NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value.</span></span> 
    - <span data-ttu-id="e7afd-122">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection):ブック内の **NamedItem** オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-122">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.</span></span>

- <span data-ttu-id="e7afd-123">[Binding](/javascript/api/excel/excel.binding):ブックのセクションへのバインドを表す抽象クラス。</span><span class="sxs-lookup"><span data-stu-id="e7afd-123">[Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.</span></span>
    - <span data-ttu-id="e7afd-124">[BindingCollection](/javascript/api/excel/excel.bindingcollection):ブック内の **Binding** オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="e7afd-124">[BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.</span></span>

## <a name="excel-javascript-api-open-specifications"></a><span data-ttu-id="e7afd-125">Excel の JavaScript API 仕様を開く</span><span class="sxs-lookup"><span data-stu-id="e7afd-125">Excel JavaScript API open specifications</span></span>

<span data-ttu-id="e7afd-126">ようにデザインし、Api を使用する Excel のアドインを開発、私たちがで使用できるようにフィードバックを[オープン API の仕様](../openspec.md)のページ。</span><span class="sxs-lookup"><span data-stu-id="e7afd-126">As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page.</span></span> <span data-ttu-id="e7afd-127">新機能が Excel JavaScript Api では、パイプラインし、設計仕様書に必要な設定を入力するどのようなを確認します。</span><span class="sxs-lookup"><span data-stu-id="e7afd-127">Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="excel-javascript-api-reference"></a><span data-ttu-id="e7afd-128">Excel の JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="e7afd-128">Excel JavaScript API reference</span></span>

<span data-ttu-id="e7afd-129">Excel の JavaScript API の詳細については、 [Excel の JavaScript API リファレンス ドキュメント](/javascript/api/excel)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7afd-129">For detailed information about Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="see-also"></a><span data-ttu-id="e7afd-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="e7afd-130">See also</span></span>

- [<span data-ttu-id="e7afd-131">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="e7afd-131">Excel add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [<span data-ttu-id="e7afd-132">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="e7afd-132">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [<span data-ttu-id="e7afd-133">GitHub 上の Excel のアドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="e7afd-133">Excel add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
