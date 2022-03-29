---
title: マニフェスト ファイルの AppDomain 要素
description: アドインで使用され、ユーザーが信頼する必要がある追加のドメインをOffice。
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: ec0dec56e75d25a2ba79eed9932aa0158be45b7e
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64486317"
---
# <a name="appdomain-element"></a>AppDomain 要素

SourceLocation 要素で指定Office、信頼する必要がある追加のドメインを[指定します](sourcelocation.md)。 ドメインを指定すると、次の効果があります。

- これにより、ドメイン内のページ、ルート、または他のリソースを、デスクトップ プラットフォーム上のアドインのルート作業ウィンドウで直接開Officeできます。 (**AppDomain** でドメインを指定する場合は、Office on the web または IFrame でリソースを開く必要はありません。また、Dialog [API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins) で開いたダイアログでリソースを開く場合も必要ありません)。
- これにより、ドメイン内のページがアドインOffice.js IFrames から API 呼び出しを実行できます。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain.com</AppDomain>`) が含まれている必要があります。
> 2. ドメインに明示的なポートがある場合は、それを含める (たとえば、`<AppDomain>https://myappdomain.com:9999</AppDomain>`)。
> 3. サブドメインを信頼する必要がある場合は、そのサブドメインを含める (例:)`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` 。 サブドメインであり、 `mysubdomain.mydomain.com` 異 `mydomain.com` なるドメインです。 両方を信頼する必要がある場合は、両方とも別々の **AppDomain 要素に含む必要** があります。
> 4. [SourceLocation](sourcelocation.md) 要素で指定されたドメインと同じドメインを一覧に表示すると、効果が得らなく、誤解を招く可能性があります。 特に、上で開発する場合 `localhost`は、 **AppDomain** 要素を作成する必要があります `localhost`。
> 5. ドメインを越える URL のセグメントを含めない。 たとえば、ページの完全な URL を含めない。
> 6. 値 *に* "/"というスラッシュを付け込む必要があります。

## <a name="contained-in"></a>含まれる場所

[AppDomains](appdomains.md)

## <a name="remarks"></a>注釈

詳細については、「[Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」をご覧ください。
