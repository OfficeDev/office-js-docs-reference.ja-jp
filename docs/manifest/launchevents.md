---
title: マニフェスト ファイルの LaunchEvents
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 02/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 762c85158390f0b541d6fcedf41d7ce49dd7f042
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487645"
---
# <a name="launchevents-element"></a>LaunchEvents 要素

サポートされているイベントに基づいてアクティブ化するアドインを構成します。 要素の子 [`<ExtensionPoint>`](extensionpoint.md) 。 詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](/office/dev/add-ins/outlook/autolaunch)。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

## <a name="syntax"></a>構文

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a>含まれる場所

[ExtensionPoint](extensionpoint.md) (**LaunchEvent** メール アドイン)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | はい |  サポートされているイベントを JavaScript ファイル内の関数にマップして、アドインのアクティブ化を行います。 |

## <a name="see-also"></a>関連項目

- [LaunchEvent](launchevent.md)
