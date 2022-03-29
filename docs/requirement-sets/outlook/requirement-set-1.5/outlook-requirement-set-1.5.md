---
title: Outlook アドイン API 要件セット 1.5
description: メールボックス API 1.5 のOutlook JavaScript API および Office JavaScript API 用に導入された機能と API。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: de9c7b8f388e969ef00c6a7e9fb9e589a9a713a9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486710"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook アドイン API 要件セット 1.5

Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!NOTE]
> このドキュメントは、最新の要件セット以外の[要件セット](../outlook-api-requirement-sets.md)のためのものです。

## <a name="whats-new-in-15"></a>1.5 の新機能

要件セット 1.5 には、要件セット [1.4 のすべての機能が含まれています](../requirement-set-1.4/outlook-requirement-set-1.4.md)。 次の機能が追加されました。

- [ピン留め可能な作業ウィンドウ](/office/dev/add-ins/outlook/pinnable-taskpane)のサポートが追加されました。
- [REST API](/office/dev/add-ins/outlook/use-rest-api) の呼び出しのサポートが追加されました。
- インラインで添付ファイルにマークを付ける機能が追加されました。
- 作業ウィンドウまたはダイアログを閉じる機能が追加されました。

### <a name="change-log"></a>変更ログ

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) が追加されました。サポートされているイベントのイベント ハンドラーを追加します。
- [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): サポートされているイベントの種類のイベント ハンドラーを削除します。
- [Office.EventType](office.md#eventtype-string) が追加されました。イベント ハンドラーに関連付けられているイベントを指定し、ItemChanged イベントのサポートが含まれるようになります。
- [Office.context.mailbox.restUrl](office.context.mailbox.md#properties) が追加されました。この電子メール アカウントの REST エンドポイントの URL を取得します。
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) が変更されました。このメソッドの新しい署名付きの新しいバージョン (`getCallbackTokenAsync([options], callback)`) が追加されました。元のバージョンは引き続き使用でき、変更されていません。
- [Office.context.ui.closeContainer](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true#office-office-ui-closecontainer-member(1)) が追加されました。
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `options` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) が変更されました。`formData.attachments` と呼ばれる `isInline` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) が変更されました。`isInline` と呼ばれる `formData.attachments` ディクショナリの新しい値。イメージがインラインでメッセージ本文で使用されることを指定するために使用されます。

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/office/dev/add-ins/quickstarts/outlook-quickstart)
- [要求セットとサポートされているクライアント](../outlook-api-requirement-sets.md)
