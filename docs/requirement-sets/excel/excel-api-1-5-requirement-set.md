---
title: Excel JavaScript API 要件セット 1.5
description: ExcelApi 1.5 要件セットの詳細。
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: a809fd74a15c9d862879aeff27becbec2ebf8ee8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486386"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 の新機能

ExcelApi 1.5 では、カスタム XML パーツが追加されます。 これらは、ブック オブジェクトの [カスタム XML パーツ コレクション](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) からアクセスできます。

## <a name="custom-xml-part"></a>カスタム XML パーツ

* ID を使用してカスタム XML パーツを取得します。
* 名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。
* パーツに関連付けられた XML 文字列を取得します。
* パーツの ID と名前空間を指定します。
* 新しいカスタム XML パーツをブックに追加します。
* XML パーツ全体を設定します。
* カスタム XML パーツを削除します。
* xpath で識別される要素から、指定された名前を持つ属性を削除します。
* xpath で XML の内容を照会します。
* 属性の挿入、更新、および削除。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.5 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.5 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_5.md)]

## <a name="see-also"></a>関連項目

* [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
