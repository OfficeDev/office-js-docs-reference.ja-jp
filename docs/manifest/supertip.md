---
title: マニフェスト ファイルの Supertip 要素
description: Supertip 要素は、リッチ ヒント (タイトルと説明の両方) を定義します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d4e59970af7c41f5303ec9cf94ef70e1373d1a59
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64501776"
---
# <a name="supertip"></a>Supertip

豊富なヒント (タイトルと説明の両方) を定義します。 これは、Button コントロールと [Menu コントロールの](control-button.md) 両方 [で使用されます](control-menu.md)。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0
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
| [Title](#title) | はい | ヒントのテキストです。 |
| [説明](#description) | はい | ヒントの説明です。<br>**注**: (Outlook) Windows Mac クライアントだけがサポートされています。 |

### <a name="title"></a>タイトル

必ず指定します。 ヒントのテキストです。 **resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

### <a name="description"></a>説明

必ず指定します。 ヒントの記述です。 **resid 属性** は 32 文字以内で、Resources 要素の **LongStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

> [!NOTE]
> たとえばOutlook、Windows Mac クライアントだけが Description 要素を **サポート** します。

## <a name="example"></a>例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
