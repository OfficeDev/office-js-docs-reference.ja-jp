# <a name="defaultsettings-element"></a><span data-ttu-id="c4fd4-101">DefaultSettings 要素</span><span class="sxs-lookup"><span data-stu-id="c4fd4-101">DefaultSettings element</span></span>

<span data-ttu-id="c4fd4-102">既定のソースの場所と、コンテンツの他の既定の設定を指定または作業ウィンドウ アドインです。</span><span class="sxs-lookup"><span data-stu-id="c4fd4-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="c4fd4-103">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c4fd4-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c4fd4-104">構文</span><span class="sxs-lookup"><span data-stu-id="c4fd4-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="c4fd4-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="c4fd4-105">Contained in</span></span>

[<span data-ttu-id="c4fd4-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c4fd4-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c4fd4-107">含めることができます。</span><span class="sxs-lookup"><span data-stu-id="c4fd4-107">Can contain</span></span>

|<span data-ttu-id="c4fd4-108">**要素**</span><span class="sxs-lookup"><span data-stu-id="c4fd4-108">**Element**</span></span>|<span data-ttu-id="c4fd4-109">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="c4fd4-109">**Content**</span></span>|<span data-ttu-id="c4fd4-110">**メール**</span><span class="sxs-lookup"><span data-stu-id="c4fd4-110">**Mail**</span></span>|<span data-ttu-id="c4fd4-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c4fd4-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c4fd4-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c4fd4-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="c4fd4-113">x</span><span class="sxs-lookup"><span data-stu-id="c4fd4-113">x</span></span>||<span data-ttu-id="c4fd4-114">x</span><span class="sxs-lookup"><span data-stu-id="c4fd4-114">x</span></span>|
|[<span data-ttu-id="c4fd4-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="c4fd4-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="c4fd4-116">x</span><span class="sxs-lookup"><span data-stu-id="c4fd4-116">x</span></span>|||
|[<span data-ttu-id="c4fd4-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="c4fd4-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="c4fd4-118">x</span><span class="sxs-lookup"><span data-stu-id="c4fd4-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="c4fd4-119">注釈</span><span class="sxs-lookup"><span data-stu-id="c4fd4-119">Remarks</span></span>

<span data-ttu-id="c4fd4-120">**DefaultSettings** 要素のソースの場所と他の設定が適用されるのは、コンテンツ アドインと作業ウィンドウ アドインのみです。メール アドインの場合は、ソース ファイルの既定の場所とその他の既定の設定を [FormSettings](formsettings.md) 要素に指定します。</span><span class="sxs-lookup"><span data-stu-id="c4fd4-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

