---
title: Excel JavaScript API オンライン専用要件セット
description: ExcelApiOnline 要件セットの詳細。
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ade48cec3e1db452206a1d35291b3f6f3e609566
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485914"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API オンライン専用要件セット

要件`ExcelApiOnline`セットは、ユーザーが使用できる機能のみを含む特別な要件セットExcel on the web。 この要件セットの API は、アプリケーションの実稼働 API と見なされます (文書化されていない動作や構造上の変更のExcel on the webされます。 `ExcelApiOnline`API は、他のプラットフォーム (Windows、Mac、iOS) の "プレビュー" API と見なされ、これらのプラットフォームではサポートされない場合があります。

要件セット内の `ExcelApiOnline` API がすべてのプラットフォームでサポートされている場合は、次にリリースされた要件セット () に追加されます`ExcelApi 1.[NEXT]`。 その新しい要件が公開されると、これらの API はから削除されます `ExcelApiOnline`。 これは、プレビューからリリースに移行する API と同様のプロモーション プロセスと考えて下さい。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 API の唯一のバージョンです。 これは、最新Excel on the webユーザーが常に 1 つのバージョンを使用できるためです。

次の表に、API の簡潔な概要を示しますが、後続の [API](#api-list) リスト テーブルでは、現在の API の詳細な一覧を `ExcelApiOnline` 示します。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| リンクされたブック | ブック間のリンクを管理します。ブックリンクの更新と破損のサポートを含む。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| ワークシートの移動イベント | コレクション内でワークシートを移動する場合、ワークシートの位置、変更元を検出します。 | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>推奨される使用法

API は `ExcelApiOnline` Excel on the webによってのみサポートされるので、アドインは、これらの API を呼び出す前に、要件セットがサポートされているのか確認する必要があります。 これにより、別のプラットフォームでオンライン専用 API を呼び出すのを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API がクロスプラットフォーム要件セットに入った後は、チェックを削除または編集する必要 `isSetSupported` があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行う場合は、必ずこれらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証 `ExcelApiOnline 1.1` 要件を指定することはできません。 Set 要素で使用する有効な値 [ではありません](/javascript/api/manifest/set)。

## <a name="api-list"></a>API リスト

次の表に、要件Excel含まれている JavaScript API の一覧を`ExcelApiOnline`示します。 すべての JavaScript API (`ExcelApiOnline`API Excel以前にリリースされた API を含む) の完全な一覧については、[JavaScript API Excel参照してください](/javascript/api/excel?view=excel-js-online&preserve-view=true)。

[!INCLUDE[API table](../../includes/excel-online.md)]
## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](excel-preview-apis.md)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
