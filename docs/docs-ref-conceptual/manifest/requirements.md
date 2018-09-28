# <a name="requirements-element"></a><span data-ttu-id="0319f-101">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="0319f-101">Requirements element</span></span>

<span data-ttu-id="0319f-102">Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="0319f-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="0319f-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="0319f-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0319f-104">構文</span><span class="sxs-lookup"><span data-stu-id="0319f-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="0319f-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="0319f-105">Contained in</span></span>

[<span data-ttu-id="0319f-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0319f-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="0319f-107">含めることができます。</span><span class="sxs-lookup"><span data-stu-id="0319f-107">Can contain</span></span>

|<span data-ttu-id="0319f-108">**要素**</span><span class="sxs-lookup"><span data-stu-id="0319f-108">**Element**</span></span>|<span data-ttu-id="0319f-109">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="0319f-109">**Content**</span></span>|<span data-ttu-id="0319f-110">**メール**</span><span class="sxs-lookup"><span data-stu-id="0319f-110">**Mail**</span></span>|<span data-ttu-id="0319f-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0319f-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0319f-112">Sets</span><span class="sxs-lookup"><span data-stu-id="0319f-112">Sets</span></span>](sets.md)|<span data-ttu-id="0319f-113">x</span><span class="sxs-lookup"><span data-stu-id="0319f-113">x</span></span>|<span data-ttu-id="0319f-114">x</span><span class="sxs-lookup"><span data-stu-id="0319f-114">x</span></span>|<span data-ttu-id="0319f-115">x</span><span class="sxs-lookup"><span data-stu-id="0319f-115">x</span></span>|
|[<span data-ttu-id="0319f-116">メソッド</span><span class="sxs-lookup"><span data-stu-id="0319f-116">Methods</span></span>](methods.md)|<span data-ttu-id="0319f-117">x</span><span class="sxs-lookup"><span data-stu-id="0319f-117">x</span></span>||<span data-ttu-id="0319f-118">x</span><span class="sxs-lookup"><span data-stu-id="0319f-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="0319f-119">注釈</span><span class="sxs-lookup"><span data-stu-id="0319f-119">Remarks</span></span>

<span data-ttu-id="0319f-120">要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0319f-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

