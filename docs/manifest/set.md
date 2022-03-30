---
title: マニフェスト ファイルの Set 要素
description: Set 要素は、Office によってアクティブ化したり、基本マニフェスト設定を上書きしたりするために Office Office アドインに必要な javaScript API 要件セットを指定します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1584a8ddff2cb03b507954d57565907eea81069c
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64502498"
---
# <a name="set-element"></a>Set 要素

この要素の意味は、マニフェストで使用される場所によって異なります。

## <a name="in-the-base-manifest"></a>基本マニフェストで

基本マニフェストで使用する場合 (つまり、祖父母 **Requirements** 要素は [OfficeApp](officeapp.md) の直接の子)、**Set** 要素は Office によってアクティブ化 [](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets)するために Office アドインが必要とする Office JavaScript API からの要件セットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>VersionOverrides 要素の孫として

[versionOverrides](versionoverrides.md) を有効にするために、Office バージョンとプラットフォーム (Windows、Mac、Web、iOS、iPad など) でサポートする必要がある Office JavaScript API の要件セットを指定します。[](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-applications-and-requirement-sets)

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

**次の要件セットに関連付けられている**。

- 祖父母の [Requirements 要素と同](requirements.md) じです。

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>含まれる場所

[Sets](sets.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|名前|string|必須|[要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。 親 Sets 要素で **指定されている場合は、DefaultMinVersion** の値を [上書き](sets.md) します。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「アドインをホストできる Office バージョンとプラットフォームを指定する」を [参照してください](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in)。

