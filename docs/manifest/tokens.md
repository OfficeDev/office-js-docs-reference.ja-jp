---
title: マニフェスト ファイル内の Tokens 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 75512fbf230da6bf60074094446f25f79f01a2f6
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486139"
---
# <a name="tokens-element"></a>Tokens 要素

テンプレート URL で使用できるトークンを定義します。 この要素の使用の詳細については、「 [Work with extended overrides of the manifest」を参照してください](/office/dev/add-ins/develop/extended-overrides)。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>含まれる場所

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>含める必要があるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[トークン](token.md)|||x|

## <a name="example"></a>例

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```