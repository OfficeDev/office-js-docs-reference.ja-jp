---
title: マニフェスト ファイルの Type 要素
description: Type 要素は、同等のアドインが COM アドインか XLL かを指定します。
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38063472e403e118baa80ee854f49a1093c1dd3d
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486134"
---
# <a name="type-element"></a>Type 要素

同等のアドインが COM アドインか XLL かを指定します。

**アドインの種類:** 作業ウィンドウ、カスタム関数

## <a name="syntax"></a>構文

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>アドインの型の値

要素には、次のいずれかの値を指定する必要 `Type` があります。

- COM: COM アドインと同等のアドインを指定します。
- XLL: 同等のアドインが XLL のExcelします。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Office アドインを既存の COM アドインと互換できるようにする](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)