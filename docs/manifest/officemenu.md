---
title: マニフェスト ファイルの OfficeMenu 要素
description: OfficeMenu 要素は、コンテキスト メニューに追加するコントロールのコレクションOfficeします。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: fb496953286ef7b4e359b5f9271b5f28bce763ec
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486031"
---
# <a name="officemenu-element"></a>OfficeMenu 要素

Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md)

## <a name="attributes"></a>属性

| 属性            | 必須 | 説明                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | はい      | 定義する OfficeMenu の種類。|

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [メニューの種類の制御](control-menu.md)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

## <a name="xsitype"></a>xsi:type

この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。

- `ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。
- `ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。

## <a name="example"></a>例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.myMenu">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />
          </Action>
        </Item>
      </Items>
    </Control>
</OfficeMenu>
```
