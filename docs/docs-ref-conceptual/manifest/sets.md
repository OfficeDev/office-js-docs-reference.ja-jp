# <a name="sets-element"></a><span data-ttu-id="c1684-101">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="c1684-101">Sets element</span></span>

<span data-ttu-id="c1684-102">Office アドインをアクティブにするために必要な JavaScript API for Office の最小限のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="c1684-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="c1684-103">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="c1684-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c1684-104">構文</span><span class="sxs-lookup"><span data-stu-id="c1684-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="c1684-105">含まれています。</span><span class="sxs-lookup"><span data-stu-id="c1684-105">Contained in</span></span>

[<span data-ttu-id="c1684-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c1684-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="c1684-107">含めることができます。</span><span class="sxs-lookup"><span data-stu-id="c1684-107">Can contain</span></span>

[<span data-ttu-id="c1684-108">Set</span><span class="sxs-lookup"><span data-stu-id="c1684-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="c1684-109">属性</span><span class="sxs-lookup"><span data-stu-id="c1684-109">Attributes</span></span>

|<span data-ttu-id="c1684-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="c1684-110">**Attribute**</span></span>|<span data-ttu-id="c1684-111">**型**</span><span class="sxs-lookup"><span data-stu-id="c1684-111">**Type**</span></span>|<span data-ttu-id="c1684-112">**必須**</span><span class="sxs-lookup"><span data-stu-id="c1684-112">**Required**</span></span>|<span data-ttu-id="c1684-113">**説明**</span><span class="sxs-lookup"><span data-stu-id="c1684-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c1684-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="c1684-114">DefaultMinVersion</span></span>|<span data-ttu-id="c1684-115">文字列</span><span class="sxs-lookup"><span data-stu-id="c1684-115">string</span></span>|<span data-ttu-id="c1684-116">省略可能</span><span class="sxs-lookup"><span data-stu-id="c1684-116">optional</span></span>|<span data-ttu-id="c1684-p101">すべての子の **Set** 要素に対して、既定の [MinVersion](set.md) 属性値を指定します。既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="c1684-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="c1684-119">備考</span><span class="sxs-lookup"><span data-stu-id="c1684-119">Remarks</span></span>

<span data-ttu-id="c1684-120">要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1684-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="c1684-121">**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c1684-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

