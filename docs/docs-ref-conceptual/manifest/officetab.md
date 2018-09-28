# <a name="officetab-element"></a><span data-ttu-id="e07a9-101">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="e07a9-101">OfficeTab element</span></span>

<span data-ttu-id="e07a9-p101">アドイン コマンドを表示するリボン タブを定義します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="e07a9-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e07a9-105">子要素</span><span class="sxs-lookup"><span data-stu-id="e07a9-105">Child elements</span></span>

|  <span data-ttu-id="e07a9-106">要素</span><span class="sxs-lookup"><span data-stu-id="e07a9-106">Element</span></span> |  <span data-ttu-id="e07a9-107">必須</span><span class="sxs-lookup"><span data-stu-id="e07a9-107">Required</span></span>  |  <span data-ttu-id="e07a9-108">Description</span><span class="sxs-lookup"><span data-stu-id="e07a9-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e07a9-109">Group</span><span class="sxs-lookup"><span data-stu-id="e07a9-109">Group</span></span>      | <span data-ttu-id="e07a9-110">はい</span><span class="sxs-lookup"><span data-stu-id="e07a9-110">Yes</span></span> |  <span data-ttu-id="e07a9-p102">コマンドのグループを定義します。既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="e07a9-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="e07a9-113">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e07a9-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="e07a9-114">**太字**の値は、デスクトップと (2016 以降 Windows 用の Word と Word のオンラインなどのオンラインの両方でサポートされます。</span><span class="sxs-lookup"><span data-stu-id="e07a9-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="e07a9-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="e07a9-115">Outlook</span></span>

- <span data-ttu-id="e07a9-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="e07a9-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="e07a9-117">Word</span><span class="sxs-lookup"><span data-stu-id="e07a9-117">Word</span></span>

- <span data-ttu-id="e07a9-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e07a9-118">**TabHome**</span></span>
- <span data-ttu-id="e07a9-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e07a9-119">**TabInsert**</span></span>
- <span data-ttu-id="e07a9-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="e07a9-120">TabWordDesign</span></span>
- <span data-ttu-id="e07a9-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="e07a9-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="e07a9-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="e07a9-122">TabReferences</span></span>
- <span data-ttu-id="e07a9-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="e07a9-123">TabMailings</span></span>
- <span data-ttu-id="e07a9-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="e07a9-124">TabReviewWord</span></span>
- <span data-ttu-id="e07a9-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e07a9-125">**TabView**</span></span>
- <span data-ttu-id="e07a9-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e07a9-126">TabDeveloper</span></span>
- <span data-ttu-id="e07a9-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e07a9-127">TabAddIns</span></span>
- <span data-ttu-id="e07a9-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="e07a9-128">TabBlogPost</span></span>
- <span data-ttu-id="e07a9-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="e07a9-129">TabBlogInsert</span></span>
- <span data-ttu-id="e07a9-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e07a9-130">TabPrintPreview</span></span>
- <span data-ttu-id="e07a9-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="e07a9-131">TabOutlining</span></span>
- <span data-ttu-id="e07a9-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="e07a9-132">TabConflicts</span></span>
- <span data-ttu-id="e07a9-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e07a9-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="e07a9-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="e07a9-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="e07a9-135">Excel</span><span class="sxs-lookup"><span data-stu-id="e07a9-135">Excel</span></span>

- <span data-ttu-id="e07a9-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e07a9-136">**TabHome**</span></span>
- <span data-ttu-id="e07a9-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e07a9-137">**TabInsert**</span></span>
- <span data-ttu-id="e07a9-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="e07a9-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="e07a9-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="e07a9-139">TabFormulas</span></span>
- <span data-ttu-id="e07a9-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="e07a9-140">**TabData**</span></span>
- <span data-ttu-id="e07a9-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="e07a9-141">**TabReview**</span></span>
- <span data-ttu-id="e07a9-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e07a9-142">**TabView**</span></span>
- <span data-ttu-id="e07a9-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e07a9-143">TabDeveloper</span></span>
- <span data-ttu-id="e07a9-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e07a9-144">TabAddIns</span></span>
- <span data-ttu-id="e07a9-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e07a9-145">TabPrintPreview</span></span>
- <span data-ttu-id="e07a9-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e07a9-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="e07a9-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e07a9-147">PowerPoint</span></span>

- <span data-ttu-id="e07a9-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e07a9-148">**TabHome**</span></span>
- <span data-ttu-id="e07a9-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e07a9-149">**TabInsert**</span></span>
- <span data-ttu-id="e07a9-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="e07a9-150">**TabDesign**</span></span>
- <span data-ttu-id="e07a9-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="e07a9-151">**TabTransitions**</span></span>
- <span data-ttu-id="e07a9-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="e07a9-152">**TabAnimations**</span></span>
- <span data-ttu-id="e07a9-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="e07a9-153">TabSlideShow</span></span>
- <span data-ttu-id="e07a9-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="e07a9-154">TabReview</span></span>
- <span data-ttu-id="e07a9-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e07a9-155">**TabView**</span></span>
- <span data-ttu-id="e07a9-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e07a9-156">TabDeveloper</span></span>
- <span data-ttu-id="e07a9-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e07a9-157">TabAddIns</span></span>
- <span data-ttu-id="e07a9-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="e07a9-158">TabPrintPreview</span></span>
- <span data-ttu-id="e07a9-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="e07a9-159">TabMerge</span></span>
- <span data-ttu-id="e07a9-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="e07a9-160">TabGrayscale</span></span>
- <span data-ttu-id="e07a9-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="e07a9-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="e07a9-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="e07a9-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="e07a9-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="e07a9-163">TabSlideMaster</span></span>
- <span data-ttu-id="e07a9-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="e07a9-164">TabHandoutMaster</span></span>
- <span data-ttu-id="e07a9-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="e07a9-165">TabNotesMaster</span></span>
- <span data-ttu-id="e07a9-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="e07a9-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="e07a9-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="e07a9-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="e07a9-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="e07a9-168">OneNote</span></span>

- <span data-ttu-id="e07a9-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="e07a9-169">**TabHome**</span></span>
- <span data-ttu-id="e07a9-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="e07a9-170">**TabInsert**</span></span>
- <span data-ttu-id="e07a9-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="e07a9-171">**TabView**</span></span>
- <span data-ttu-id="e07a9-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="e07a9-172">TabDeveloper</span></span>
- <span data-ttu-id="e07a9-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="e07a9-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="e07a9-174">Group</span><span class="sxs-lookup"><span data-stu-id="e07a9-174">Group</span></span>

<span data-ttu-id="e07a9-p104">タブの UI 拡張ポイントのグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。**id** は最大 125 文字の文字列です。「[Group 要素](group.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e07a9-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="e07a9-179">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="e07a9-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
