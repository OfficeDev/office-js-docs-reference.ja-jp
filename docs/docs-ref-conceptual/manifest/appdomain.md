# <a name="appdomain-element"></a><span data-ttu-id="57100-101">AppDomain 要素</span><span class="sxs-lookup"><span data-stu-id="57100-101">AppDomain element</span></span>

<span data-ttu-id="57100-102">アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。</span><span class="sxs-lookup"><span data-stu-id="57100-102">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="57100-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="57100-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="57100-104">構文</span><span class="sxs-lookup"><span data-stu-id="57100-104">Syntax</span></span>

```XML
<AppDomain>string </AppDomain>
```

## <a name="contained-in"></a><span data-ttu-id="57100-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="57100-105">Contained in</span></span>

[<span data-ttu-id="57100-106">AppDomains</span><span class="sxs-lookup"><span data-stu-id="57100-106">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="57100-107">備考</span><span class="sxs-lookup"><span data-stu-id="57100-107">Remarks</span></span>

<span data-ttu-id="57100-108">SourceLocation 要素で指定されたもの以外の他のドメインを指定するのには、**アプリケーション ドメイン**と**アプリケーション ドメイン**の要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="57100-108">The  **AppDomains** and **AppDomain** elements are used to specify any additional domains other than the one specified in the SourceLocation element.</span></span> <span data-ttu-id="57100-109">詳細については、「Office アドイン XML マニフェスト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="57100-109">For more information, see Office Add-ins XML manifest.</span></span>

