---
title: Excel JavaScript プレビュー API
description: JavaScript API のExcel詳細。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3a7901d7d8dbd749843506060d347cdf94320e9b
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486365"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

次の表に、API の簡潔な概要を示しますが、後続の [API リスト テーブル](#api-list) には詳細な一覧が示されています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [データ型](/office/dev/add-ins/excel/excel-data-types-overview) | 書式付き番号と web Excelのサポートを含む、既存のデータ型の拡張。 | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)、 [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)、 [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)、 [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)、 [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)、 [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)、 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)、 [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)、 [StringCellValue](/javascript/api/excel/excel.stringcellvalue)、 [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)、 [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [データ型のエラー](/office/dev/add-ins/excel/excel-data-types-concepts.md#improved-error-support) | 拡張データ型をサポートするエラー オブジェクト。 | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)、[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)、[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)、[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)、[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)、[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)、[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)、[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)、[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)、[NullErrorCellValue、NumErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)、RefErrorCellValue、[](/javascript/api/excel/excel.numerrorcellvalue)[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)、[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue) [](/javascript/api/excel/excel.referrorcellvalue)|
| ドキュメント タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| ID | 表示名や電子メール アドレスなど、ユーザー ID を管理します。 | [IDENTITY](/javascript/api/excel/excel.identity)、[IdentityCollection、](/javascript/api/excel/excel.identitycollection)[IdentityEntity](/javascript/api/excel/excel.identityentity) |
| リンクされたデータ型 | 外部ソースからデータに接続されたデータExcelサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| テーブルのスタイル | フォント、罫線、塗りつぶしの色、および表のスタイルの他の側面のコントロールを提供します。 | [テーブル](/javascript/api/excel/excel.table)、[ピボットテーブル、](/javascript/api/excel/excel.pivottable)[スライサー](/javascript/api/excel/excel.slicer) |
| ワークシートの保護 | 承認されていないユーザーがワークシート内で指定した範囲に変更を加えなかねない。 | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中Excel JavaScript API の一覧を示します。 すべての JavaScript API (プレビュー API Excel以前にリリースされた API を含む) の完全な一覧については、[JavaScript API Excel参照してください](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

[!INCLUDE[API table](../../includes/excel-preview.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
