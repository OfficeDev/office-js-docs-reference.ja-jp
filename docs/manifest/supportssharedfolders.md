---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: SupportsSharedFolders 要素は、共有フォルダーと共有メールボックスのシナリオOutlookアドインを使用できるかどうかを定義します。
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 54ef1407a4c7e0c0999cb26def04eba7d820ecd8
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502071"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) シナリオで、アドインを使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.8 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.8](../requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8.md)

**SupportsSharedFolders 要素の例を次に示** します。

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
