---
title: マニフェスト ファイルの AllFormFactors 要素
description: すべてのフォーム ファクターについてアドインの設定を指定します。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 25d9b815cf7db440abce2c8859c9ebb15ea5df1a
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486485"
---
# <a name="allformfactors-element"></a>AllFormFactors 要素

すべてのフォーム ファクターについてアドインの設定を指定します。 現在、**AllFormFactors** を使用する機能はカスタム関数のみです。 **AllFormFactors** は、カスタム関数を使用するときの必須要素です。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

> [!NOTE]
> この要素は、Excel、mac、Windowsでサポートされます。 他のアプリケーションや iOS Office Android ではサポートされていません。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  はい |  アドインが機能を公開する場所を定義します。 |

## <a name="allformfactors-example"></a>AllFormFactors の例

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
