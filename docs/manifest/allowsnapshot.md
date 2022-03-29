---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 723817557020f4ec3dbe5b3135877fe49bf67acb
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485871"
---
# <a name="allowsnapshot-element"></a>AllowSnapshot 要素

ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。

**アドインの種類:** コンテンツ

## <a name="syntax"></a>構文

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>解説

 > [!IMPORTANT]
 > **AllowSnapshot** の既定値は `true` です。 これにより、Office アドインをサポートしないバージョンの Office アプリケーションでドキュメントを開くユーザーに対してアドインのイメージが表示され、アプリケーションがアドインをホストするサーバーに接続できない場合は、アドインの静的イメージが提供されます。 However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.
