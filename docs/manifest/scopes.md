---
title: マニフェスト ファイルの Scopes 要素
description: Scopes 要素には、アドインが外部リソースに接続するために必要なアクセス許可が含まれる。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3b67977421a6ed0da205a051d0ecaf480c96fd5c
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486002"
---
# <a name="scopes-element"></a>Scopes 要素

Microsoft などの外部リソースに対してアドインに必要なアクセス許可がGraph。 Microsoft Graphリソースである場合、AppSource は Scopes 要素を使用して同意ダイアログ ボックスを作成します。 ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。

**アドインの種類:** 作業ウィンドウ、メール、コンテンツ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- コンテンツ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [IdentityAPI 1.3](../requirement-sets/common/identity-api-requirement-sets.md)

**スコープは** 、マニフェスト内の [WebApplicationInfo](webapplicationinfo.md) 要素の子要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Scope**                |  はい     |   アクセス許可の名前。たとえば、Files.Read.All またはプロファイルです。 |

## <a name="example"></a>例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
