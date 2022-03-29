---
title: マニフェスト ファイルの TabletSettings 要素
description: TabletSettings 要素は、メール アドインをタブレットで使用するときに適用されるコントロール設定を指定します。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3b2af4a220e01017c268125c3db3edc61cf7ec82
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485842"
---
# <a name="tabletsettings-element"></a>TabletSettings 要素

メール アドインがタブレットで使用されるときに適用される制御の設定を指定します。

> [!IMPORTANT]
> この`TabletSettings`要素は、従来の Outlook on the web (通常はオンプレミスの Exchange サーバーの古いバージョンに接続されている) および Outlook 2013 でWindows。 Android と iOS Outlookサポートするには、「[Add-ins for Outlook Mobile」を参照してください](/office/dev/add-ins/outlook/outlook-mobile-addins)。

**アドインの種類:** メール

## <a name="syntax"></a>構文

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>含まれる場所

[Form](form.md)
