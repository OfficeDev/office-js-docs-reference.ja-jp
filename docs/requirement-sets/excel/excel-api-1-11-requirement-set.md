---
title: Excel JavaScript API 要件セット 1.11
description: ExcelApi 1.11 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1d002e1884e8b18981b00160d8a498ec0ff3ea51
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503184"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>JavaScript API 1.11 Excel新機能

ExcelApi 1.11 では、コメントとブック レベルのコントロール (ブックの保存や閉じるなど) のサポートが強化されました。 また、ローカライズのアカウントに役立つカルチャ設定へのアクセスも追加されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| コメント [メンション](/office/dev/add-ins/excel/excel-add-ins-comments#mentions) |コメントを使用して他のブック ユーザーにタグを付け、通知します。 | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| コメント [の解決](/office/dev/add-ins/excel/excel-add-ins-comments#resolve-comment-threads) | コメント スレッドを解決し、解決状態を取得します。 | [コメント](/javascript/api/excel/excel.comment) |
| [カルチャの設定](/office/dev/add-ins/excel/excel-add-ins-workbooks#access-application-culture-settings) | 数値の書式設定など、ブックの文化システム設定を取得します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [アプリケーション](/javascript/api/excel/excel.application) |
| [切り取りと貼り付け (moveTo)](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | Range のセル内の切り取りExcel機能をレプリケートします。 | [範囲](/javascript/api/excel/excel.range) |
| ブックを[保存](/office/dev/add-ins/excel/excel-add-ins-workbooks.md#save-the-workbook)して[閉じる](/office/dev/add-ins/excel/excel-add-ins-workbooks#close-the-workbook) | ブックを保存して閉じます。 | [Workbook](/javascript/api/excel/excel.workbook) |
| ワークシート のイベント | ワークシートの計算と非表示の行に関するその他のイベントとイベント情報。 | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.11 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.11 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true) 以前の Excel API を参照してください。

[!INCLUDE[API table](../../includes/excel-1_11.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
