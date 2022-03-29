---
title: マニフェスト ファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 154e69367b00a49c896d9fb108d621ba212c4580
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64485963"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 要素

同等の COM アドインまたは XLL の下位互換性を指定します。

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**アドインの種類:** 作業ウィンドウ、メール、カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

## <a name="syntax"></a>構文

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>含める必要があるもの

[Type](type.md)

## <a name="can-contain"></a>含めることができるもの

[ProgId](progid.md)
[FileName](filename.md)

## <a name="remarks"></a>注釈

COM アドインを同等のアドインとして指定するには、要素と要素の両方を `ProgId` 指定 `Type` します。 XLL を同等のアドインとして指定するには、要素と要素の両方を `FileName` 指定 `Type` します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Office アドインを既存の COM アドインと互換できるようにする](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)