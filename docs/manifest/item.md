---
title: マニフェスト ファイル内の Item 要素
description: メニュー内のアイテムを指定します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 940ecd91b22677c8476a05ab8b0bf033168b293b
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487658"
---
# <a name="item-element"></a>Item 要素

メニュー内のアイテムを指定します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)     | はい |  ボタンのテキストです。 |
|  [Supertip](supertip.md)  | はい |  このボタンのヒントです。    |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。 Item 要素の Action 子 **は** 1 つ **のみです** 。  |
|  [Enabled (有効)](enabled.md)    | いいえ |  アドインの起動時にコントロールを有効にするかどうかを指定します。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにボタンを表示するかどうかを指定します。 使用する場合は、最初の子 *要素である* 必要があります。 |

### <a name="label"></a>ラベル

ボタンのテキストを、32 文字以内にできる唯一の属性を使用して指定し、Resources 要素の **ShortStrings** 子の **String** 要素の **id** 属性の値に設定する [必要](resources.md)があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)。

## <a name="examples"></a>例

例については、「Control [of type Menu」を参照してください](control-menu.md)。