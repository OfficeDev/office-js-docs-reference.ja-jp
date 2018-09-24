# <a name="outlook-javascript-api-requirement-sets"></a><span data-ttu-id="a808c-101">Outlook の JavaScript API の要件の設定</span><span class="sxs-lookup"><span data-stu-id="a808c-101">Outlook JavaScript API requirement sets</span></span>

<span data-ttu-id="a808c-102">Outlook アドインの場合は、その[マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)で[要求](/javascript/office/manifest/requirements)要素を使用して、必要な API バージョンを宣言します。</span><span class="sxs-lookup"><span data-stu-id="a808c-102">Outlook add-ins declare what API versions they require by using the [Requirements](/javascript/office/manifest/requirements) element in their [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests).</span></span> <span data-ttu-id="a808c-103">Outlook アドインには、`Name` 属性が `Mailbox` に設定され、`MinVersion` 属性がアドインのシナリオをサポートする最小 API 要件セットに設定された [Set](/javascript/office/manifest/set) 要素が常に含まれます。</span><span class="sxs-lookup"><span data-stu-id="a808c-103">Outlook add-ins always include a [Set](/javascript/office/manifest/set) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.</span></span>

<span data-ttu-id="a808c-104">たとえば、次のマニフェストのスニペットは、最小要件セットの 1.1 を表します。</span><span class="sxs-lookup"><span data-stu-id="a808c-104">For example, the following manifest snippet indicates a minimum requirement set of 1.1:</span></span>

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

<span data-ttu-id="a808c-105">すべての Outlook API は `Mailbox` [要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)に属しています。</span><span class="sxs-lookup"><span data-stu-id="a808c-105">All Outlook APIs belong to the `Mailbox` [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).</span></span> <span data-ttu-id="a808c-106">`Mailbox` 要件のセットにはバージョンがあります。リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。</span><span class="sxs-lookup"><span data-stu-id="a808c-106">The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set.</span></span> <span data-ttu-id="a808c-107">すべての Outlook クライアントが最新の API セットをサポートしているわけではありません。しかし Outlook クライアントが要件セットのサポートを宣言する場合は、その要件セットの API すべてがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a808c-107">Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, it supports all of the APIs in that requirement set.</span></span>

<span data-ttu-id="a808c-p103">マニフェストに要件セットの最小バージョンを設定することで、アドインが表示される Outlook クライアントをコントロールできます。クライアントが最小要件セットをサポートしない場合、アドインはロードされません。たとえば、要件セットのバージョン 1.3 が指定されている場合、1.3 以上をサポートしていない Outlook クライアントには表示されません。</span><span class="sxs-lookup"><span data-stu-id="a808c-p103">Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.</span></span>

## <a name="using-apis-from-later-requirement-sets"></a><span data-ttu-id="a808c-111">後続の要件セットからの API の使用</span><span class="sxs-lookup"><span data-stu-id="a808c-111">Using APIs from later requirement sets</span></span>

<span data-ttu-id="a808c-112">要件のセットを設定しても、アドインを使用できる利用可能な Api は制限されません。</span><span class="sxs-lookup"><span data-stu-id="a808c-112">Setting a requirement set does not limit the available APIs that the add-in can use.</span></span> <span data-ttu-id="a808c-113">など、アドインの指定要件の設定、1.1、1.3 をサポートしている Outlook クライアントで実行されている場合、アドイン Api を使用して要件セット 1.3 から。</span><span class="sxs-lookup"><span data-stu-id="a808c-113">For example, if the add-in specifies requirement set 1.1, but it is running in an Outlook client which support 1.3, the add-in can use APIs from requirement set 1.3.</span></span>

<span data-ttu-id="a808c-114">新しい Api を使用するには、開発者だけをチェックできますの存在の JavaScript の標準的な方法を使用します。</span><span class="sxs-lookup"><span data-stu-id="a808c-114">To use newer APIs, developers can just check for their existence by using standard JavaScript technique:</span></span>

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

<span data-ttu-id="a808c-115">このようなチェックは、マニフェストで指定された要件セットバージョンに存在する API には必要ありません。</span><span class="sxs-lookup"><span data-stu-id="a808c-115">No such checks are necessary for any APIs which are present in the requirement set version specified in in the manifest.</span></span>

## <a name="choosing-a-minimum-requirement-set"></a><span data-ttu-id="a808c-116">最小要件セットの選択</span><span class="sxs-lookup"><span data-stu-id="a808c-116">Choosing a minimum requirement set</span></span>

<span data-ttu-id="a808c-117">開発者は、アドインを使用するために必要な、シナリオで必須の API のセットが含まれている初期の要件セットを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a808c-117">Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.</span></span>

## <a name="clients"></a><span data-ttu-id="a808c-118">クライアント</span><span class="sxs-lookup"><span data-stu-id="a808c-118">Clients</span></span>

<span data-ttu-id="a808c-119">以下のクライアントは、Outlook のアドインをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="a808c-119">The following clients support Outlook add-ins.</span></span>

| <span data-ttu-id="a808c-120">クライアント</span><span class="sxs-lookup"><span data-stu-id="a808c-120">Client</span></span> | <span data-ttu-id="a808c-121">サポートされる API の要件セット</span><span class="sxs-lookup"><span data-stu-id="a808c-121">Supported API requirement sets</span></span> |
| --- | --- |
| <span data-ttu-id="a808c-122">Windows 用の outlook 2019</span><span class="sxs-lookup"><span data-stu-id="a808c-122">Outlook 2019 for Windows</span></span> | <span data-ttu-id="a808c-123">1.1、1.2、1.3、1.4、1.5、1.6</span><span class="sxs-lookup"><span data-stu-id="a808c-123">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="a808c-124">For Mac outlook 2019</span><span class="sxs-lookup"><span data-stu-id="a808c-124">Outlook 2019 for Mac</span></span> | <span data-ttu-id="a808c-125">1.1、1.2、1.3、1.4、1.5、1.6</span><span class="sxs-lookup"><span data-stu-id="a808c-125">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="a808c-126">Windows 版 Outlook 2016 (クイック実行)</span><span class="sxs-lookup"><span data-stu-id="a808c-126">Outlook 2016 (Click-to-Run) for Windows</span></span> | <span data-ttu-id="a808c-127">1.1、1.2、1.3、1.4、1.5、1.6、1.7</span><span class="sxs-lookup"><span data-stu-id="a808c-127">1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7</span></span> |
| <span data-ttu-id="a808c-128">Windows 版 Outlook 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="a808c-128">Outlook 2016 (MSI) for Windows</span></span> | <span data-ttu-id="a808c-129">1.1、1.2、1.3、1.4</span><span class="sxs-lookup"><span data-stu-id="a808c-129">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="a808c-130">Outlook 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a808c-130">Outlook 2016 for Mac</span></span> | <span data-ttu-id="a808c-131">1.1、1.2、1.3、1.4、1.5、1.6</span><span class="sxs-lookup"><span data-stu-id="a808c-131">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="a808c-132">Windows 版 Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="a808c-132">Outlook 2013 for Windows</span></span> | <span data-ttu-id="a808c-133">1.1、1.2、1.3、1.4</span><span class="sxs-lookup"><span data-stu-id="a808c-133">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="a808c-134">Outlook for iPhone</span><span class="sxs-lookup"><span data-stu-id="a808c-134">Outlook for iPhone</span></span> | <span data-ttu-id="a808c-135">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="a808c-135">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="a808c-136">Outlook for Android</span><span class="sxs-lookup"><span data-stu-id="a808c-136">Outlook for Android</span></span> | <span data-ttu-id="a808c-137">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="a808c-137">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="a808c-138">Outlook on the web (Office 365 および Outlook.com)</span><span class="sxs-lookup"><span data-stu-id="a808c-138">Outlook on the web (Office 365 and Outlook.com)</span></span> | <span data-ttu-id="a808c-139">1.1、1.2、1.3、1.4、1.5、1.6</span><span class="sxs-lookup"><span data-stu-id="a808c-139">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="a808c-140">Outlook Web App (Exchange 2013 On-Premise)</span><span class="sxs-lookup"><span data-stu-id="a808c-140">Outlook Web App (Exchange 2013 On-Premise)</span></span> | <span data-ttu-id="a808c-141">1.1</span><span class="sxs-lookup"><span data-stu-id="a808c-141">1.1</span></span> |
| <span data-ttu-id="a808c-142">Outlook Web App (Exchange 2016 On-Premise)</span><span class="sxs-lookup"><span data-stu-id="a808c-142">Outlook Web App (Exchange 2016 On-Premise)</span></span> | <span data-ttu-id="a808c-p105">1.1, 1.2. 1.3</span><span class="sxs-lookup"><span data-stu-id="a808c-p105">1.1, 1.2. 1.3</span></span> |

> [!NOTE]
> <span data-ttu-id="a808c-145">Outlook 2013 で 1.3 のサポートは、 [2015年 12 月 8日、Outlook 2013 (KB3114349) の更新](https://support.microsoft.com/kb/3114349)の一部として追加されました。</span><span class="sxs-lookup"><span data-stu-id="a808c-145">Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349).</span></span> <span data-ttu-id="a808c-146">[2016年 9 月 13日、Outlook 2013 (KB3118280) の更新](https://support.microsoft.com/help/3118280)の一部として、Outlook 2013 で 1.4 のサポートが追加されました。</span><span class="sxs-lookup"><span data-stu-id="a808c-146">Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).</span></span>
