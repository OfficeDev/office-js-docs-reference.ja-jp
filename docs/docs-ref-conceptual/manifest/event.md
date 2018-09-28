# <a name="event-element"></a><span data-ttu-id="3b9ce-101">Event 要素</span><span class="sxs-lookup"><span data-stu-id="3b9ce-101">Event element</span></span>

<span data-ttu-id="3b9ce-102">アドインでイベント ハンドラーを定義します。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="3b9ce-103">`Event`の要素は現在、Outlook を Office 365 に web 上でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-103">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="3b9ce-104">属性</span><span class="sxs-lookup"><span data-stu-id="3b9ce-104">Attributes</span></span>

|  <span data-ttu-id="3b9ce-105">属性</span><span class="sxs-lookup"><span data-stu-id="3b9ce-105">Attribute</span></span>  |  <span data-ttu-id="3b9ce-106">必須</span><span class="sxs-lookup"><span data-stu-id="3b9ce-106">Required</span></span>  |  <span data-ttu-id="3b9ce-107">説明</span><span class="sxs-lookup"><span data-stu-id="3b9ce-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3b9ce-108">Type</span><span class="sxs-lookup"><span data-stu-id="3b9ce-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="3b9ce-109">はい</span><span class="sxs-lookup"><span data-stu-id="3b9ce-109">Yes</span></span>  | <span data-ttu-id="3b9ce-110">処理するイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="3b9ce-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="3b9ce-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="3b9ce-112">はい</span><span class="sxs-lookup"><span data-stu-id="3b9ce-112">Yes</span></span>  | <span data-ttu-id="3b9ce-p101">イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="3b9ce-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="3b9ce-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="3b9ce-116">はい</span><span class="sxs-lookup"><span data-stu-id="3b9ce-116">Yes</span></span>  | <span data-ttu-id="3b9ce-117">イベント ハンドラーの関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="3b9ce-118">Type 属性</span><span class="sxs-lookup"><span data-stu-id="3b9ce-118">Type attribute</span></span>

<span data-ttu-id="3b9ce-p102">必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="3b9ce-122">イベントの種類</span><span class="sxs-lookup"><span data-stu-id="3b9ce-122">Event type</span></span>  |  <span data-ttu-id="3b9ce-123">説明</span><span class="sxs-lookup"><span data-stu-id="3b9ce-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="3b9ce-124">ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="3b9ce-125">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="3b9ce-125">FunctionExecution attribute</span></span>

<span data-ttu-id="3b9ce-p103">必須です。`synchronous` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="3b9ce-128">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="3b9ce-128">FunctionName attribute</span></span>

<span data-ttu-id="3b9ce-p104">必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3b9ce-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```