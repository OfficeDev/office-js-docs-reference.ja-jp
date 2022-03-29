---
title: マニフェスト ファイルの ExtendedPermissions 要素
description: アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 78ff52c6b4e529709591c424bd01f540a85806ab
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486287"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 要素

アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可のコレクションを定義します。 要素 `ExtendedPermissions` は [VersionOverrides の子要素です](versionoverrides.md)。

> [!IMPORTANT]
> この要素のサポートは、要件セット 1.9 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [Mailbox 1.9](/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  いいえ   | アドインが関連付けられた API または機能にアクセスするために必要な拡張アクセス許可を定義します。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` 例

次に、要素の例を示 `ExtendedPermissions` します。

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
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>含まれる場所

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>含めることができるもの

[ExtendedPermission](extendedpermission.md)
