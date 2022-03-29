---
title: マニフェスト ファイルの AppDomains 要素
description: '`SourceLocation` Office アドインが使用する要素で指定されたドメインに加えて、すべてのドメインを一覧表示し、Office。'
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: 6bf1785cf11e31648d9bc69e101cd5a5cf3ecb9f
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487688"
---
# <a name="appdomains-element"></a>AppDomains 要素

要素で指定された`SourceLocation`ドメインに加えて、Office アドインが使用し、Office によって信頼される必要があるドメインを一覧表示します。 これにより、ドメイン内のページでアドイン内の IFrame から Office.js API を呼び出し、その他の効果があります。 追加の各ドメインに、**AppDomain** 要素を指定します。

 **アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> **AppDomain** 要素の値に制限があります。 詳細については、「 [AppDomain」を参照してください](appdomain.md)。

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

[AppDomain](appdomain.md)

## <a name="remarks"></a>解説

アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。 この要素は空にできません。
