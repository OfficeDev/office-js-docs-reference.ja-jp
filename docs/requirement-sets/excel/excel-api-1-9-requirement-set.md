---
title: Excel JavaScript API 要件セット 1.9
description: ExcelApi 1.9 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ef0f5979c9cfd41b59b3e7ae86f96fb3d7b35746
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502743"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>JavaScript API 1.9 Excel新機能

1.9 の要件セットにより、500 件を超える新しい Excel API が 導入されました。 最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [Shapes](/office/dev/add-ins/excel/excel-add-ins-shapes) | 画像、幾何学な図形、テキスト ボックスを挿入、位置変更、書式設定します。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape) [Image](/javascript/api/excel/excel.image) |
| [オート フィルター](/office/dev/add-ins/excel/excel-add-ins-worksheets#filter-data) | 範囲にフィルターを追加します。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [エリア](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges) | 連続していない範囲をサポートします。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特別なセル](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges#get-special-cells-from-multiple-ranges) | ある範囲内に日付、コメント、数式を含むセルを取得します。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [検索](/office/dev/add-ins/excel/excel-add-ins-ranges-string-match) | ある範囲またはワークシート内で値や数式を見つけます。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [コピーと貼り付け](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | 範囲間で値、書式、数式をコピーします。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [計算](/office/dev/add-ins/excel/performance#suspend-calculation-temporarily) | Excel 計算エンジンを細かく操作できます。 | [アプリケーション](/javascript/api/excel/excel.application) |
| 新しいグラフ | 新しくサポートされたグラフである、マップ、箱ひげ図、ウォーターフォール、サンバースト、パレート、 じょうごをお試しください。 | [Chart](/javascript/api/excel/excel.charttype) |
| 範囲の形式 | 範囲の形式の新しい機能です。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.9 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.9 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_9.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
