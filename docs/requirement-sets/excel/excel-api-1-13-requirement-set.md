---
title: Excel JavaScript API 要件セット 1.13
description: ExcelApi 1.13 要件セットの詳細。
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 79e493e25fd5c0a72cf62e0b1af9ce06fb262055
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486099"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>JavaScript API 1.13 Excel新機能

ExcelApi 1.13 では、Base64 エンコード文字列からブックにワークシートを挿入するメソッドと、ブックのアクティブ化を検出するイベントが追加されました。 また、API を追加して数式の変更を追跡し、数式の直接依存セルを特定することで、範囲内の数式のサポートを増やしました。 さらに、代替テキスト、スタイル、空のセル管理用の PivotLayout API を追加することで、ピボットテーブルのサポートを拡張しました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [数式の変更イベント](/office/dev/add-ins/excel/excel-add-ins-worksheets.md#detect-formula-changes) | 変更の原因となるイベントのソースと種類を含む、数式の変更を追跡します。 | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [数式の依存](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | 数式の直接依存セルを見つける。 | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [ワークシートの挿入](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | 別のブックのワークシートを現在のブックに Base64 エンコード文字列として挿入します。 | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [ピボットテーブル ピボットレイアウト](/office/dev/add-ins/excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Alt テキストと空のセル管理の新しいサポートを含む、PivotLayout クラスの拡張。 | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.13 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.13 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_13.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
