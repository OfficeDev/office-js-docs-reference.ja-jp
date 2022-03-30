---
title: Excel JavaScript API 要件セット 1.12
description: ExcelApi 1.12 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 002050f1b1c0b4420229146614182b53976f0f4c
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502358"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>JavaScript API 1.12 Excel新機能

ExcelApi 1.12 では、動的配列を追跡し、数式の直接の前例を見つけ出す API を追加することで、範囲の数式のサポートが増加しました。 ピボットテーブル フィルターの API コントロールも追加しました。 コメント、カルチャ設定、カスタム プロパティの機能領域でも改善が行われた。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメント イベント](/office/dev/add-ins/excel/excel-add-ins-comments#comment-events) | コメント コレクションに追加、変更、削除のイベントを追加します。| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 日付と時刻 [のカルチャ設定](/office/dev/add-ins/excel/excel-add-ins-workbooks#access-application-culture-settings) | 日付と時刻の書式設定に関する追加の文化設定にアクセスできます。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [アプリケーション](/javascript/api/excel/excel.application) |
| [直接の前例](/office/dev/add-ins/excel/excel-add-ins-ranges-precedents) | セルの数式の評価に使用される範囲を返します。| [範囲](/javascript/api/excel/excel.range#getdirectprecedents--) |
| ピボット フィルター | ピボットテーブルのフィールドに値駆動型フィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [範囲の流出](/office/dev/add-ins/excel/excel-add-ins-ranges-dynamic-arrays) | 動的配列の結果に関連付けられた範囲をアドイン [が検索](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) できます。 | [範囲](/javascript/api/excel/excel.range) |
| [ワークシート レベルのカスタム プロパティ](/office/dev/add-ins/excel/excel-add-ins-workbooks#worksheet-level-custom-properties) | ブック レベルのスコープに加えて、ワークシート レベルのカスタム プロパティのスコープを設定できます。 | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.12 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.12 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_12.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
