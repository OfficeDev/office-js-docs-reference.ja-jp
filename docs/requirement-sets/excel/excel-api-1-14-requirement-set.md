---
title: Excel JavaScript API 要件セット 1.14
description: ExcelApi 1.14 要件セットの詳細。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: a7e5874ef285ce72abf82cd3b4ee8087e6c98cd3
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502750"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>JavaScript API 1.14 Excel新機能

ExcelApi 1.14 には、グラフのデータ テーブル機能を制御するオブジェクト、数式のすべての先行セルを検索するメソッド、ワークシートの保護状態の変更を追跡するワークシート保護イベントが追加されました。 また、、 、、[`getItemOrNullObject`](/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties)など、`CommentCollection``ShapeCollection`オブジェクトの複数のメソッドを追加し、エラー処理`StyleCollection`を改善しました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [グラフ データ テーブル](/office/dev/add-ins/excel/excel-add-ins-charts#add-and-format-a-chart-data-table) | グラフ上のデータ テーブルの外観、書式設定、および表示を制御します。 | [グラフ](/javascript/api/excel/excel.chart)、 [ChartDataTable](/javascript/api/excel/excel.chartdatatable)、 [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [数式の前例](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents#get-the-precedents-of-a-formula) | 数式のすべての前のセルを返します。 | [範囲](/javascript/api/excel/excel.range) |
| クエリ | 名前、更新日、クエリ数など、Power Query 属性を取得します。 | [クエリ](/javascript/api/excel/excel.query)、 [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [ワークシート保護イベント](/office/dev/add-ins/excel/excel-add-ins-worksheets#detect-changes-to-the-worksheet-protection-state) | ワークシートの保護状態に対する変更と、それらの変更のソースを追跡します。 | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.14 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.14 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_14.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
