---
title: Excel JavaScript API 要件セット 1.4
description: ExcelApi 1.4 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 555e75fe4d7a210e22a74081b29869a2dc17b0cb
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503471"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 の新機能

要件セット 1.4 の Excel JavaScript API に新たに追加された機能は次のとおりです。

## <a name="named-item-add-and-new-properties"></a>名前付きアイテムの追加と新しいプロパティ

新しいプロパティ:

* `comment`
* `scope` - ワークシートまたはブックのスコープ付きアイテム。
* `worksheet` - 名前付きアイテムのスコープが設定されているワークシートを返します。

新しいメソッド:

* `add(name: string, reference: Range or string, comment: string)` - 指定したスコープのコレクションに新しい名前を追加します。
* `addFormulaLocal(name: string, formula: string, comment: string)` - 数式のユーザーのロケールを使用して、指定したスコープのコレクションに新しい名前を追加します。

## <a name="settings-api-in-the-excel-namespace"></a>Excel 名前空間内の Setting API

[Setting](/javascript/api/excel/excel.setting) オブジェクトは、ドキュメントに永続的に適用される設定のキーと値のペアを表します。 `Excel.Setting` の機能は `Office.Settings` と同等ですが、共通 API のコールバック モデルではなくバッチ API 構文を使用します。

API には、キー `getItem()` を介して設定エントリを取得 `add()` し、指定した key:value 設定ペアをブックに追加する機能が含まれます。

## <a name="others"></a>Others

* テーブルの列名を設定します。
* テーブル列をテーブルの末尾に追加します。
* 一度に複数の行をテーブルに追加します。
* `range.getColumnsAfter(count: number)` および `range.getColumnsBefore(count: number)` を使用して、現在の Range オブジェクトの左右にある特定の数の列を取得します。
* [OrNullObject\* メソッドとプロパティ](/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties): この機能を使用すると、キーを使用してオブジェクトを取得できます。 オブジェクトが存在しない場合、返されるオブジェクトのプロパティ `isNullObject` は true になります。 これにより、開発者は例外処理を介してオブジェクトを処理することなく、オブジェクトが存在するかどうかを確認できます。 メソッド `*OrNullObject` は、ほとんどのコレクション オブジェクトで使用できます。

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.4 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.4 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_4.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
