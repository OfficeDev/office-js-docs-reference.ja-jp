---
title: マニフェスト ファイルの IconUrl 要素
description: IconUrl 要素は、挿入 UX とストア内Officeアドインを表すイメージの URL をOfficeします。
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8f4026c9f97fc147b65532e370e18327bc07076c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487681"
---
# <a name="iconurl-element"></a>IconUrl 要素

挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|種類|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メール アドインの場合、アイコンは **FileManage**  >  アドイン UI (Outlook) または **設定** > **Manage** アドイン UI (Outlook on the web) に表示されます。 コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。 アドインを AppSource に発行する場合は、すべてのアドインの種類に対して、 [アイコンも AppSource](https://appsource.microsoft.com) で使用されます。

画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。 コンテンツ アプリおよび作業ウィンドウ アプリの場合、指定する画像は 32 x 32 ピクセルにする必要があります。 メール アプリの場合、画像の解像度は 64 x 64 ピクセルである必要があります。 また、[HighResolutionIconUrl](highresolutioniconurl.md) 要素を使用して、高 DPI 画面で実行Officeクライアント アプリケーションで使用するアイコンを指定する必要があります。 詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。

実行時に要素の値を `IconUrl` 変更することはできません。現在はサポートされていません。
