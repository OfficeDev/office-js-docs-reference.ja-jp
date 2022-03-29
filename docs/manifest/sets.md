---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、Office アドインが Office でアクティブ化したり、基本マニフェスト設定を上書きしたりするために必要な Office JavaScript API の最小セットを指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: a67e719136667529df0cd1066c850771291d0f75
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486142"
---
# <a name="sets-element"></a>Sets 要素

この要素の意味は、マニフェストで使用される場所によって異なります。

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェスト (つまり、親 **Requirements** 要素は [OfficeApp](officeapp.md) の直接の [子) で](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets)使用する場合、**Sets** 要素は、Office によってアクティブ化するために Office アドインが必要とする Office JavaScript API 要件 (要件セット) の最小サブセットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>VersionOverrides 要素の孫として

[VersionOverrides](versionoverrides.md) を有効にするために、Office バージョンと[プラットフォーム (Windows](/office/dev/add-ins/develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API 要件 (要件セット) の最小セットを指定します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 親の [Requirements 要素と同](requirements.md) じです。

**次の要件セットに関連付けられている**。

- 親の [Requirements 要素と同](requirements.md) じです。

## <a name="syntax"></a>構文

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>含まれる場所

[Requirements](requirements.md)

## <a name="can-contain"></a>含めることができるもの

[Set](set.md)

## <a name="attributes"></a>属性

|属性|種類|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|文字列|省略可能|すべての子 Set 要素 **の既定の MinVersion** 属性値を [指定](set.md) します。 既定値は "1.1" です。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「アドインをホストできる Office バージョンとプラットフォームを指定する」を [参照してください](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in)。

