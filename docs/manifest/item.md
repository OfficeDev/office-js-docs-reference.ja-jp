---
title: マニフェスト ファイル内の Item 要素
description: メニュー内のアイテムを指定します。
ms.date: 03/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: a2ea0d1dd8d0b01f91340c61ebf2e21abf598b84
ms.sourcegitcommit: 07f76f08c7952f0777a7aa163fa734f88f6dee9b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/01/2022
ms.locfileid: "64604248"
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
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)     | はい |  メニュー項目のテキスト。 |
|  [Supertip](supertip.md)  | はい |  メニュー項目のスーパーヒント。    |
|  [Icon](icon.md)      | 不要 |  メニュー項目のイメージ。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。 Item 要素の Action 子 **は** 1 つ **のみです** 。  |
|  [Enabled (有効)](enabled.md)    | 不要 |  アドインの起動時にメニュー項目を有効にするかどうかを指定します。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにメニュー項目を表示するかどうかを指定します。 使用する場合は、最初の子 *要素である* 必要があります。 |

### <a name="label"></a>Label

ボタンのテキストを、32 文字以内にできる唯一の属性を使用して指定し、Resources 要素の **ShortStrings** 子の **String** 要素の **id** 属性の値に設定する [必要](resources.md)があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/common/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../requirement-sets/outlook/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="examples"></a>例

例については、「Control [of type Menu」を参照してください](control-menu.md)。