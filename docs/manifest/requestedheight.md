---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメール アドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: e0589e81e8905c4fc8c7a8e50ec7c14038035677
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486581"
---
# <a name="requestedheight-element"></a>RequestedHeight 要素

コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。

**アドインの種類:** コンテンツ、メール

## <a name="syntax"></a>構文

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>含まれる場所

- [DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能
- [DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) の値は、**DetectedEntity** 拡張ポイントの場合は 140 ~ 450、[**CustomPane**](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) 拡張ポイントの場合は 32 ~ 450 です (非推奨)
