---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: b0ff18a44753c7c9d8a8f8886e40fb11c0e155d8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486007"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の子 [`<Host>`](host.md) 。

> [!NOTE]
> Windows `<Runtimes>` Office で実行する場合、マニフェストに要素を持つアドインは、必ずしも同じ Webview コントロールで実行されるとは限りません。 Windows および Office のバージョンで、どの Web ビュー コントロールが通常使用されるのかを判断する方法の詳細については、「Office アドインで使用[される](/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)ブラウザー」を参照してください。webView2 (Chromium ベース) `<Runtimes>` で Microsoft Edge を使用する場合に説明されている条件が満たされている場合、アドインは要素を持っているかどうかに応じ、そのブラウザーを使用します。 ただし、これらの条件`<Runtimes>`が満たされない場合、要素を持つアドインは、Windows またはバージョンに関係なく、常に Internet Explorer 11 をMicrosoft 365します。

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

[Host](host.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [ランタイム](runtime.md) | はい |  アドインのランタイム。 **重要**: 現時点では、定義できる要素は 1 つ `<Runtime>` のみです。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [イベント ベースのOutlookアドインを構成する](/office/dev/add-ins/outlook/autolaunch)
