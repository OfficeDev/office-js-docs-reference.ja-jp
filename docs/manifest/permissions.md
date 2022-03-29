---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、アドインの API アクセス レベルOffice指定します。
ms.date: 06/26/2020
ms.localizationpriority: medium
ms.openlocfilehash: 48369c71b94c286d357868bc1f0620cb09b58878
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486026"
---
# <a name="permissions-element"></a>Permissions 要素

Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

コンテンツ アドインおよび作業ウィンドウ アドインの場合:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

詳細については、「コンテンツアドインと作業ウィンドウ アドインでの [API](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) 使用のアクセス許可の要求」および「Outlookについて」を[参照してください](/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)。
