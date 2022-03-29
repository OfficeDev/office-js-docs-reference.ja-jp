---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 03/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1965405724076860b6aad8651cfb8c651f182c53
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486022"
---
# <a name="runtime-element"></a>Runtime 要素

共有 JavaScript ランタイムを使用して、さまざまなコンポーネントすべてが同じランタイムで実行されるアドインを構成します。 要素の子 [`<Runtimes>`](runtimes.md) 。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

 - 作業ウィンドウ 1.0
 - メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [SharedRuntime 1.1](../requirement-sets/common/shared-runtime-requirement-sets.md) (作業ウィンドウ アドインで使用する場合のみ)。

[!include[Runtimes support](../includes/runtimes-note.md)]

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

- [ランタイム](runtimes.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [Override](override.md) | なし | **Outlook**: LaunchEvent 拡張ポイント ハンドラーにデスクトップで必要Outlook [JavaScript ファイルの URL の場所を指定](extensionpoint.md#launchevent)します。 **重要**: 現時点では、定義できる要素は `<Override>` 1 つで、型である必要があります `javascript`。|

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 `resid` 32 `Url` `id` 文字以内で、要素内の要素の属性と一致する必要`Resources`があります。 |
|  [有効期間](#lifetime-attribute)  |  いいえ  | 既定値は is `lifetime` `short` であり、指定する必要はない。 Outlookベースのアクティブ化アドインでは、値のみを使用`short`します。 アドインで共有ランタイムを使用する場合Excelに値を明示的に設定します`long`。 |

### <a name="lifetime-attribute"></a>lifetime 属性

省略可能。 アドインの実行が許可されている時間の長さを表します。

**使用可能な値**

`short`: 既定。 イベント ベースのOutlookアドインでのみ使用されます。アドインがアクティブ化されると、プラットフォームで指定された最大時間実行されます。 現在、約 5 分です。 これは、ユーザーがサポートする唯一のOutlook。

`long`: 共有 JavaScript ランタイムを構成 [する場合にのみ使用されます](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)。 アドインは、ドキュメントを開いて無期限に実行できます。 たとえば、作業ウィンドウ コードは、ユーザーが作業ウィンドウを閉じても実行を続行します。 これは、共有ランタイムでサポートされている唯一の値です。

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [イベント ベースのOutlookアドインを構成する](/office/dev/add-ins/outlook/autolaunch)
