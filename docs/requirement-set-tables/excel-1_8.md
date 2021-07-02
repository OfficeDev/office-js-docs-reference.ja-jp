| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|データの検証に使用する演算子。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|ソース カテゴリ ラベルのレベルを参照して、グラフ カテゴリ ラベル レベル列挙定数を指定します。|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|空白のセルをグラフにプロットする方法を指定します。|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|列や行がグラフのデータ系列として使用される方法を指定します。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|true の場合、可視セルだけがプロットされます。|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|グラフがアクティブ化されると発生します。|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|グラフが非アクティブ化された場合に発生します。|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|グラフのプロット領域を表します。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|ソース 系列名のレベルを参照して、グラフ系列名レベルの列挙定数を指定します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|値が値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.chart#style)|グラフのグラフ スタイルを指定します。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|アクティブ化されたグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|グラフをアクティブ化するワークシートの ID を取得します。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|ワークシートに追加されるグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|グラフが追加されるワークシートの ID を取得します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[配置](/javascript/api/excel/excel.chartaxis#alignment)|指定した軸目盛ラベルの配置を指定します。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|値軸がカテゴリの間でカテゴリ軸と交差する場合に指定します。|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multilevel)|軸がマルチレベルの場合に指定します。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|軸目盛ラベルの書式コードを指定します。|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|ラベルのレベル間の距離と、最初のレベルと軸線の間の距離を指定します。|
||[position](/javascript/api/excel/excel.chartaxis#position)|他の軸が交差する指定した軸位置を指定します。|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|他の軸が交差する軸位置を指定します。|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|他の軸が交差する指定した軸位置を設定します。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|グラフ軸目盛ラベルのテキストの向きを指定します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|グラフの塗りつぶしの書式設定を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|色、線のスタイル、太さなど、グラフ軸のタイトルの罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|グラフ軸のタイトルの塗りつぶしの書式設定を指定します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|グラフ要素の罫線の書式設定をクリアします。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|グラフがアクティブ化されると発生します。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|ワークシートに新しいグラフが追加された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|グラフが非アクティブ化された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|グラフが削除された場合に発生します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autotext)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|グラフのデータ ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|グラフ データ ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autotext)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|グラフ データ ラベルの水平方向の配置を指定します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|データ ラベルの形式コードを指定します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|データ ラベルのテキストの向きを示す角度を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|非アクティブ化されているグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|グラフが非アクティブ化されているワークシートの ID を取得します。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|ワークシートから削除されたグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|グラフが削除されるワークシートの ID を取得します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|グラフの凡例の凡例エントリの高さを指定します。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|グラフ凡例の凡例エントリのインデックスを指定します。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|グラフの凡例エントリの左の値を指定します。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|グラフ凡例エントリの上部を指定します。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|グラフ Legend の凡例エントリの幅を表します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|プロット領域の高さの値を指定します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|プロット領域の内側の高さの値を指定します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|プロット領域の内側の左の値を指定します。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|プロット領域の内側の上の値を指定します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|プロット領域の内側の幅の値を指定します。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|プロット領域の左の値を指定します。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|プロット領域の位置を指定します。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|グラフプロット領域の書式を指定します。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|プロット領域の上の値を指定します。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|プロット領域の幅の値を指定します。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|グラフプロット領域の罫線属性を指定します。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|背景の書式設定情報を含むオブジェクトの塗りつぶしの形式を指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|指定した系列のグループを指定します。|
||[爆発](/javascript/api/excel/excel.chartseries#explosion)|円グラフまたはドーナツ グラフのスライスの展開値を指定します。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|最初の円グラフまたはドーナツ グラフのスライスの角度を度 (垂直方向から時計回り) で指定します。|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True の場合Excelに対応する場合は、アイテム内のパターンを反転します。|
||[オーバーラップ](/javascript/api/excel/excel.chartseries#overlap)|横棒と縦棒の配置方法を指定します。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|系列内のすべてのデータ ラベルのコレクションを表します。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|円グラフまたは円グラフの 2 番目のセクションのサイズを、プライマリ 円グラフのサイズに対する割合で指定します。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|円グラフまたは円グラフの 2 つのセクションを分割する方法を指定します。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True の場合Excelデータ マーカーに異なる色またはパターンを割り当てる必要があります。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[label](/javascript/api/excel/excel.charttrendline#label)|グラフの近似曲線のラベルを表します。|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True の場合、トレンドラインの r-2 乗値がグラフに表示されます。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autotext)|傾向線ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|A1 スタイル表記を使用してグラフの傾向線ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|グラフの傾向線ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|グラフのトレンドライン ラベルの左端からグラフ領域の左端までの距離をポイントで表します。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|傾向線ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|グラフの傾向線ラベルの形式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|グラフの傾向線ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|グラフのトレンドライン ラベルの上端からグラフ領域の上端までの距離をポイントで表します。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|グラフの傾向線ラベルの垂直方向の配置を表します。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|色、線のスタイル、太さなど、罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|現在のグラフの傾向線ラベルの塗りつぶしの形式を指定します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|グラフの傾向線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|ユーザーの入力規則のカスタム数式。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy の位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|DataPivotHierarchy を既定値にリセットします。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|データを特定の集計計算として表示する必要がある場合に指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|DataPivotHierarchy のすべての項目を表示する場合に指定します。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|名前または ID で DataPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|名前に基づいて DataPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|現在の軸からピボット階層を削除します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|現在の範囲からデータの入力規則をクリアします。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|空白のセルに対してデータ検証を実行する場合に指定します。|
||[Prompt](/javascript/api/excel/excel.datavalidation#prompt)|ユーザーがセルを選択するときにプロンプトを表示します。|
||[type](/javascript/api/excel/excel.datavalidation#type)|データ検証の種類については、「詳細 `Excel.DataValidationType` 」を参照してください。|
||[有効](/javascript/api/excel/excel.datavalidation#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
||[ルール](/javascript/api/excel/excel.datavalidation#rule)|さまざまな種類のデータ検証条件を含むデータ検証ルール。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|エラー通知メッセージを表します。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|ユーザーが無効なデータを入力した場合にエラー通知ダイアログを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|データ検証アラートの種類については、「詳細」 `Excel.DataValidationAlertStyle` を参照してください。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|エラー通知ダイアログのタイトルを表します。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|プロンプトのメッセージを指定します。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|ユーザーがデータ検証を使用してセルを選択するときにプロンプトを表示する場合に指定します。|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|プロンプトのタイトルを指定します。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[カスタム](/javascript/api/excel/excel.datavalidationrule#custom)|データ検証条件のカスタム数式。|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|日付のデータ検証条件。|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|10 進数のデータ検証条件。|
||[リスト](/javascript/api/excel/excel.datavalidationrule#list)|リストのデータ検証条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|テキストの長さデータの検証条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|時刻のデータ検証条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|数値データの検証条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|データの検証に使用する演算子。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|FilterPivotHierarchy を既定値にリセットします。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|名前または ID で FilterPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|名前に基づいて FilterPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|セル ドロップダウンにリストを表示するかどうかを指定します。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|データの入力規則のリストのソース。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField の名前。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField の ID。|
||[アイテム](/javascript/api/excel/excel.pivotfield#items)|PivotField に関連付けられているピボット フィールドを返します。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField を並べ替えます。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField の小計。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|コレクション内のピボット フィールドの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|名前または ID で PivotField を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|名前でピボットフィールドを取得します。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy の名前。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy の ID。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|名前または ID で PivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|名前に基づいて PivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem の名前。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem の ID。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|PivotItem が表示される場合に指定します。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|コレクション内の PivotItems の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|名前または ID で PivotItem を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|名前で PivotItem を取得します。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|ピボットテーブルの列ラベルが存在する範囲を返します。|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|ピボットテーブルのデータ値が存在する範囲を返します。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|ピボットテーブルのフィルター エリアの範囲を返します。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|ピボットテーブルの行ラベルが存在する範囲を返します。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|ピボットテーブル レポートに列の総計が表示される場合に指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|ピボットテーブル レポートに行の総計が表示される場合に指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|このプロパティは、ピボット `SubtotalLocationType` テーブルのすべてのフィールドを示します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|ピボットテーブルを削除します。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|ピボットテーブルのデータ ピボット階層。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottable#hierarchies)|ピボットテーブルのピボット階層。|
||[レイアウト](/javascript/api/excel/excel.pivottable#layout)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|ピボットテーブルの行ピボット階層。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: \| \| Range string Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|指定したソース データに基づいてピボットテーブルを追加し、移動先範囲の左上のセルに挿入します。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|dataValidation オブジェクトを返します。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|RowColumnPivotHierarchy を既定値にリセットします。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|名前または ID で RowColumnPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|名前に基づいて RowColumnPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[ランタイム](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|現在の作業ウィンドウまたはコンテンツ アドインの JavaScript イベントを切り替えます。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|PivotField を使用して、型に応じて計算を基に計算を行います (該当する場合 `ShowAs` `ShowAsCalculation` は、それ以外 `null` の場合)。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|型に応じて該当する場合は、計算の基に設定するアイテム `ShowAs` `ShowAsCalculation` 。それ以外の場合 `null` 。|
||[計算](/javascript/api/excel/excel.showasrule#calculation)|`ShowAs`PivotField に使用する計算。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|セル内のテキスト配置が等しい分布に設定されている場合に、テキストが自動的にインデントされる場合に指定します。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|スタイルで適用されるテキストの向き。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|に `Automatic` 設定されている場合 `true` 、他のすべての値は、 を設定するときに無視 `Subtotals` されます。|
||[平均](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[製品](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[差異](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|数値 ID を返します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|ブックが `true` 読み取り専用モードで開いている場合に返します。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|ワークシートの計算時に発生します。|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|グリッド線がユーザーに表示される場合に指定します。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|ユーザーに見出しを表示する場合に指定します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|計算が発生したワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|ブック内のワークシートが計算される場合に発生します。|
