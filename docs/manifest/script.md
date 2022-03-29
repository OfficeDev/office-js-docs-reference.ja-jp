---
title: マニフェスト ファイルの Script 要素
description: Script 要素は、カスタム関数がカスタム 関数で使用するスクリプト設定を定義Excel。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: af63a1c5be397b5a7923715d5739cead7cd15017
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486572"
---
# <a name="script-element"></a>Script 要素

Excel でカスタム関数によって使用されるスクリプトの設定を定義します。

**アドインの種類:** カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/excel/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。|

## <a name="example"></a>例

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
