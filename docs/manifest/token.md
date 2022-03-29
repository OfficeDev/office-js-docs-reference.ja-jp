---
title: マニフェスト ファイル内の Token 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: f2064cdd6b9d0a4319ef8c01d056531811c7a8a2
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486566"
---
# <a name="token-element"></a>Token 要素

個々の URL トークンを定義します。 この要素の使用の詳細については、「 [Work with extended overrides of the manifest」を参照してください](/office/dev/add-ins/develop/extended-overrides)。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>含まれる場所

[トークン](tokens.md)

## <a name="can-contain"></a>含めることができるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>属性

|属性|説明|
|:-----|:-----|
|DefaultValue|子要素に条件が一致する場合、このトークンの `<Override>` 既定値。|
|名前|トークン名。 この名前はユーザー定義です。 トークンの種類は type 属性によって決まります。|
|xsi:type|トークンの種類を定義します。 この属性は、次のいずれかの値に設定する必要があります`"LocaleToken"`。 `"RequirementsToken"`|

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