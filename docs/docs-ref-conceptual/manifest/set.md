# <a name="set-element"></a><span data-ttu-id="17b69-101">Set 要素</span><span class="sxs-lookup"><span data-stu-id="17b69-101">Set element</span></span>

<span data-ttu-id="17b69-102">Office アドインをアクティブにするために必要な JavaScript API for Office の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="17b69-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="17b69-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="17b69-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="17b69-104">構文</span><span class="sxs-lookup"><span data-stu-id="17b69-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="17b69-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="17b69-105">Contained in</span></span>

[<span data-ttu-id="17b69-106">Sets</span><span class="sxs-lookup"><span data-stu-id="17b69-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="17b69-107">属性</span><span class="sxs-lookup"><span data-stu-id="17b69-107">Attributes</span></span>

|<span data-ttu-id="17b69-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="17b69-108">**Attribute**</span></span>|<span data-ttu-id="17b69-109">**型**</span><span class="sxs-lookup"><span data-stu-id="17b69-109">**Type**</span></span>|<span data-ttu-id="17b69-110">**必須**</span><span class="sxs-lookup"><span data-stu-id="17b69-110">**Required**</span></span>|<span data-ttu-id="17b69-111">**説明**</span><span class="sxs-lookup"><span data-stu-id="17b69-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="17b69-112">名前</span><span class="sxs-lookup"><span data-stu-id="17b69-112">Name</span></span>|<span data-ttu-id="17b69-113">文字列</span><span class="sxs-lookup"><span data-stu-id="17b69-113">string</span></span>|<span data-ttu-id="17b69-114">必須</span><span class="sxs-lookup"><span data-stu-id="17b69-114">required</span></span>|<span data-ttu-id="17b69-115">[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。</span><span class="sxs-lookup"><span data-stu-id="17b69-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="17b69-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="17b69-116">MinVersion</span></span>|<span data-ttu-id="17b69-117">文字列</span><span class="sxs-lookup"><span data-stu-id="17b69-117">string</span></span>|<span data-ttu-id="17b69-118">省略可能</span><span class="sxs-lookup"><span data-stu-id="17b69-118">optional</span></span>|<span data-ttu-id="17b69-p101">アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](sets.md) 要素に指定されている場合は、その値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="17b69-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="17b69-121">備考</span><span class="sxs-lookup"><span data-stu-id="17b69-121">Remarks</span></span>

<span data-ttu-id="17b69-122">要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="17b69-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="17b69-123">**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="17b69-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="17b69-124">メール アドインの場合、1 つだけです`"Mailbox"`利用可能な要件を設定します。</span><span class="sxs-lookup"><span data-stu-id="17b69-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="17b69-125">この要件のセットには、outlook でメールのアドインでサポートされている API の全体のサブセットが含まれているし、指定する必要があります、`"Mailbox"`要件は、メールでこのアドインのマニフェストの設定 (は省略可能な場合と同様の内容とタスクのウィンドウ - アドイン)。</span><span class="sxs-lookup"><span data-stu-id="17b69-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="17b69-126">また、メールのアドインの特定のメソッドのサポートを宣言できません。</span><span class="sxs-lookup"><span data-stu-id="17b69-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
