# <a name="getstarted-element"></a><span data-ttu-id="68bed-101">GetStarted 要素</span><span class="sxs-lookup"><span data-stu-id="68bed-101">GetStarted element</span></span>

<span data-ttu-id="68bed-p101">アドインが、Word、Excel、PowerPoint、OneNote のホストにインストールされているときに表示される吹き出しで使用される情報を提供します。**GetStarted** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="68bed-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="68bed-104">子要素</span><span class="sxs-lookup"><span data-stu-id="68bed-104">Child elements</span></span>

| <span data-ttu-id="68bed-105">要素</span><span class="sxs-lookup"><span data-stu-id="68bed-105">Element</span></span>                       | <span data-ttu-id="68bed-106">必須</span><span class="sxs-lookup"><span data-stu-id="68bed-106">Required</span></span> | <span data-ttu-id="68bed-107">説明</span><span class="sxs-lookup"><span data-stu-id="68bed-107">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="68bed-108">タイトル</span><span class="sxs-lookup"><span data-stu-id="68bed-108">Title</span></span>](#title)               | <span data-ttu-id="68bed-109">はい</span><span class="sxs-lookup"><span data-stu-id="68bed-109">Yes</span></span>      | <span data-ttu-id="68bed-110">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="68bed-110">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="68bed-111">説明</span><span class="sxs-lookup"><span data-stu-id="68bed-111">Description</span></span>](#description)   | <span data-ttu-id="68bed-112">はい</span><span class="sxs-lookup"><span data-stu-id="68bed-112">Yes</span></span>      | <span data-ttu-id="68bed-113">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="68bed-113">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="68bed-114">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="68bed-114">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="68bed-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="68bed-115">No</span></span>       | <span data-ttu-id="68bed-116">アドインの詳細を説明するページの URL。</span><span class="sxs-lookup"><span data-stu-id="68bed-116">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="68bed-117">Title</span><span class="sxs-lookup"><span data-stu-id="68bed-117">Title</span></span> 

<span data-ttu-id="68bed-p102">必須。吹き出しの一番上に使用するタイトル。**resid** 属性は [Resources](resources.md) セクションの **ShortStrings** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="68bed-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="68bed-121">説明</span><span class="sxs-lookup"><span data-stu-id="68bed-121">Description</span></span>

<span data-ttu-id="68bed-p103">必須。吹き出しの説明/本文の内容。**resid** 属性は [Resources](resources.md) セクションの **LongStrings** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="68bed-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="68bed-125">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="68bed-125">LearnMoreUrl</span></span>

<span data-ttu-id="68bed-p104">必須。ユーザーがアドインの詳細を参照できるページの URL。**resid** 属性は [Resources](resources.md) セクションの **Urls** 要素にある有効な ID を参照します。</span><span class="sxs-lookup"><span data-stu-id="68bed-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="68bed-129">**LearnMoreUrl**は、Word、Excel、または PowerPoint のクライアントで現在は表示されません。</span><span class="sxs-lookup"><span data-stu-id="68bed-129">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="68bed-130">これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="68bed-130">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="68bed-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="68bed-131">See also</span></span>

<span data-ttu-id="68bed-132">次のコード サンプルでは、**GetStarted** 要素を使用しています。</span><span class="sxs-lookup"><span data-stu-id="68bed-132">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="68bed-133">テーブルとグラフの書式設定を操作するための Excel Web アドイン</span><span class="sxs-lookup"><span data-stu-id="68bed-133">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="68bed-134">Word アドインの JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="68bed-134">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="68bed-135">PowerPoint アドインで Microsoft Graph を使用して Excel のグラフを挿入します。</span><span class="sxs-lookup"><span data-stu-id="68bed-135">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
