# <a name="override-element"></a><span data-ttu-id="55660-101">Override 要素</span><span class="sxs-lookup"><span data-stu-id="55660-101">Override element</span></span>

<span data-ttu-id="55660-102">追加ロケールの設定の値を指定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="55660-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="55660-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="55660-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="55660-104">構文</span><span class="sxs-lookup"><span data-stu-id="55660-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="55660-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="55660-105">Contained in</span></span>

|<span data-ttu-id="55660-106">**要素**</span><span class="sxs-lookup"><span data-stu-id="55660-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="55660-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="55660-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="55660-108">説明</span><span class="sxs-lookup"><span data-stu-id="55660-108">Description</span></span>](description.md)|
|[<span data-ttu-id="55660-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="55660-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="55660-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="55660-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="55660-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="55660-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="55660-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="55660-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="55660-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="55660-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="55660-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="55660-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="55660-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="55660-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="55660-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="55660-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="55660-117">属性</span><span class="sxs-lookup"><span data-stu-id="55660-117">Attributes</span></span>

|<span data-ttu-id="55660-118">**属性**</span><span class="sxs-lookup"><span data-stu-id="55660-118">**Attribute**</span></span>|<span data-ttu-id="55660-119">**型**</span><span class="sxs-lookup"><span data-stu-id="55660-119">**Type**</span></span>|<span data-ttu-id="55660-120">**必須**</span><span class="sxs-lookup"><span data-stu-id="55660-120">**Required**</span></span>|<span data-ttu-id="55660-121">**説明**</span><span class="sxs-lookup"><span data-stu-id="55660-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="55660-122">Locale</span><span class="sxs-lookup"><span data-stu-id="55660-122">Locale</span></span>|<span data-ttu-id="55660-123">文字列</span><span class="sxs-lookup"><span data-stu-id="55660-123">string</span></span>|<span data-ttu-id="55660-124">必須</span><span class="sxs-lookup"><span data-stu-id="55660-124">required</span></span>|<span data-ttu-id="55660-125">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="55660-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="55660-126">値</span><span class="sxs-lookup"><span data-stu-id="55660-126">Value</span></span>|<span data-ttu-id="55660-127">文字列</span><span class="sxs-lookup"><span data-stu-id="55660-127">string</span></span>|<span data-ttu-id="55660-128">必須</span><span class="sxs-lookup"><span data-stu-id="55660-128">required</span></span>|<span data-ttu-id="55660-129">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="55660-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="55660-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="55660-130">See also</span></span>

- [<span data-ttu-id="55660-131">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="55660-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
