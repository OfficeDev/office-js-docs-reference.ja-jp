---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、カスタム関数がカスタム 関数で使用するメタデータ設定を定義Excel。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: b321f0059a6cfc6f2e8855f36128a076a1266755
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486440"
---
# <a name="metadata-element"></a>MetaData 要素

Excel でカスタム関数によって使用されるメタデータの設定を定義します。

**アドインの種類:** カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|  要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。 |

## <a name="example"></a>例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
