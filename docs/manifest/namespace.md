---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、カスタム関数がカスタム関数で使用する名前空間を定義Excel。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf05f7c950e2be03baf056284bdcb0bd6a5b5d0e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487640"
---
# <a name="namespace-element"></a>Namespace 要素

Excel でカスタム関数によって使用される名前空間を定義します。

**アドインの種類:** カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  いいえ  | [Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。 32 文字以内で指定できます。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<Namespace resid="namespace" />
```
