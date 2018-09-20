# <a name="group-element"></a><span data-ttu-id="f3f3d-101">Group 要素</span><span class="sxs-lookup"><span data-stu-id="f3f3d-101">Group element</span></span>

<span data-ttu-id="f3f3d-p101">タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="f3f3d-105">属性</span><span class="sxs-lookup"><span data-stu-id="f3f3d-105">Attributes</span></span>

|  <span data-ttu-id="f3f3d-106">属性</span><span class="sxs-lookup"><span data-stu-id="f3f3d-106">Attribute</span></span>  |  <span data-ttu-id="f3f3d-107">必須</span><span class="sxs-lookup"><span data-stu-id="f3f3d-107">Required</span></span>  |  <span data-ttu-id="f3f3d-108">説明</span><span class="sxs-lookup"><span data-stu-id="f3f3d-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f3f3d-109">id</span><span class="sxs-lookup"><span data-stu-id="f3f3d-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="f3f3d-110">はい</span><span class="sxs-lookup"><span data-stu-id="f3f3d-110">Yes</span></span>  | <span data-ttu-id="f3f3d-111">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="f3f3d-112">id 属性</span><span class="sxs-lookup"><span data-stu-id="f3f3d-112">id attribute</span></span>

<span data-ttu-id="f3f3d-p102">必須。グループの一意識別子。最大 125 文字の文字列です。マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f3f3d-117">子要素</span><span class="sxs-lookup"><span data-stu-id="f3f3d-117">Child elements</span></span>
|  <span data-ttu-id="f3f3d-118">要素</span><span class="sxs-lookup"><span data-stu-id="f3f3d-118">Element</span></span> |  <span data-ttu-id="f3f3d-119">必須</span><span class="sxs-lookup"><span data-stu-id="f3f3d-119">Required</span></span>  |  <span data-ttu-id="f3f3d-120">説明</span><span class="sxs-lookup"><span data-stu-id="f3f3d-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f3f3d-121">Label</span><span class="sxs-lookup"><span data-stu-id="f3f3d-121">Label</span></span>](#label)      | <span data-ttu-id="f3f3d-122">はい</span><span class="sxs-lookup"><span data-stu-id="f3f3d-122">Yes</span></span> |  <span data-ttu-id="f3f3d-123">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="f3f3d-124">Control</span><span class="sxs-lookup"><span data-stu-id="f3f3d-124">Control</span></span>](#control)    | <span data-ttu-id="f3f3d-125">はい</span><span class="sxs-lookup"><span data-stu-id="f3f3d-125">Yes</span></span> |  <span data-ttu-id="f3f3d-126">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="f3f3d-127">Label</span><span class="sxs-lookup"><span data-stu-id="f3f3d-127">Label</span></span> 

<span data-ttu-id="f3f3d-p103">必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="f3f3d-131">Control</span><span class="sxs-lookup"><span data-stu-id="f3f3d-131">Control</span></span>
<span data-ttu-id="f3f3d-132">1 つのグループに少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="f3f3d-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```