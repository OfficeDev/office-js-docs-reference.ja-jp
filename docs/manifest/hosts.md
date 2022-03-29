---
title: マニフェスト ファイルの Hosts 要素
description: アドインがOfficeするクライアント アプリケーションOfficeを指定します。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 060a98229b90d7f6a8c467af21dfef4c63c922d9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486605"
---
# <a name="hosts-element"></a>Hosts 要素

アドインがOfficeするクライアント アプリケーションOfficeを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

## <a name="as-child-of-versionoverrides-element"></a>VersionOverrides 要素の子として

このセクションの情報 *は、***Hosts** 要素が [VersionOverrides の子である場合にのみ適用されます](versionoverrides.md)。

この要素は、基本マニフェスト **の Hosts** 要素をオーバーライドします。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest)。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  はい   |  ホストとその設定について説明します。 |
