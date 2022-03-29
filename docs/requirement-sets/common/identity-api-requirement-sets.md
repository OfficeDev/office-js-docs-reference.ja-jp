---
title: Identity API の要件セット
description: ID API 要件は、アドインOffice情報を設定します。
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 49579536fa085a6e5ebe95db350da9508798a95c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485998"
---
# <a name="identity-api-requirement-sets"></a>ID API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Identity API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、および id API 要件セットのビルド番号またはバージョン番号Officeします。

|  要件セット  | Office 2021 以降のWindows<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(両方のサブスクリプション<br> Mac 2019 以降Office 1 回の購入)   | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | ビルド 16.0.14326.20454 以降 | バージョン 2008 (ビルド 13127.20000) 以降 | 非サポート | 16.40 以降 | Microsoft Office SharePoint OnlineとOneDrive\* |

\*現在、要件セットは、Office on the webおよびドキュメントから開いているドキュメントMicrosoft Office SharePoint OnlineサポートOneDrive。

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook ID API 要件セット

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> イベント ベースのライセンス認証を使用する Outlook アドインでは、[OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth) インターフェイスは Office バージョン 2108 (ビルド 14326.20258) 以降の Windows Office でサポートされます。 次[のOffice。Auth インターフェイスは](/javascript/api/office/office.auth)バージョン 2109 (ビルド 14425.10000) 以降でサポートされています。 バージョンに応じて詳細を確認するには、[Office 2021](/officeupdates/update-history-office-2021) または [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) の更新履歴ページと、Office クライアントのバージョンと更新チャネルを見つける[方法を参照してください](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office アプリケーションと API 要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
