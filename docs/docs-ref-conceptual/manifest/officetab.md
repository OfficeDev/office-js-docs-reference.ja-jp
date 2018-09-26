# <a name="officetab-element"></a><span data-ttu-id="dd96e-101">OfficeTab 要素</span><span class="sxs-lookup"><span data-stu-id="dd96e-101">OfficeTab element</span></span>

<span data-ttu-id="dd96e-p101">アドイン コマンドを表示するリボン タブを定義します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="dd96e-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dd96e-105">子要素</span><span class="sxs-lookup"><span data-stu-id="dd96e-105">Child elements</span></span>

|  <span data-ttu-id="dd96e-106">要素</span><span class="sxs-lookup"><span data-stu-id="dd96e-106">Element</span></span> |  <span data-ttu-id="dd96e-107">必須</span><span class="sxs-lookup"><span data-stu-id="dd96e-107">Required</span></span>  |  <span data-ttu-id="dd96e-108">Description</span><span class="sxs-lookup"><span data-stu-id="dd96e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dd96e-109">Group</span><span class="sxs-lookup"><span data-stu-id="dd96e-109">Group</span></span>      | <span data-ttu-id="dd96e-110">はい</span><span class="sxs-lookup"><span data-stu-id="dd96e-110">Yes</span></span> |  <span data-ttu-id="dd96e-p102">コマンドのグループを定義します。既定のタブには、アドインごとに 1 つのグループのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="dd96e-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="dd96e-113">ホストごとの有効なタブ `id` 値は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="dd96e-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="dd96e-114">**太字**の値は、デスクトップと (2016 以降 Windows 用の Word と Word のオンラインなどのオンラインの両方でサポートされます。</span><span class="sxs-lookup"><span data-stu-id="dd96e-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="dd96e-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="dd96e-115">Outlook</span></span>

- <span data-ttu-id="dd96e-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="dd96e-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="dd96e-117">Word</span><span class="sxs-lookup"><span data-stu-id="dd96e-117">Word</span></span>

- <span data-ttu-id="dd96e-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dd96e-118">**TabHome**</span></span>
- <span data-ttu-id="dd96e-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dd96e-119">**TabInsert**</span></span>
- <span data-ttu-id="dd96e-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="dd96e-120">TabWordDesign</span></span>
- <span data-ttu-id="dd96e-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="dd96e-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="dd96e-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="dd96e-122">TabReferences</span></span>
- <span data-ttu-id="dd96e-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="dd96e-123">TabMailings</span></span>
- <span data-ttu-id="dd96e-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="dd96e-124">TabReviewWord</span></span>
- <span data-ttu-id="dd96e-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dd96e-125">**TabView**</span></span>
- <span data-ttu-id="dd96e-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dd96e-126">TabDeveloper</span></span>
- <span data-ttu-id="dd96e-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dd96e-127">TabAddIns</span></span>
- <span data-ttu-id="dd96e-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="dd96e-128">TabBlogPost</span></span>
- <span data-ttu-id="dd96e-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="dd96e-129">TabBlogInsert</span></span>
- <span data-ttu-id="dd96e-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dd96e-130">TabPrintPreview</span></span>
- <span data-ttu-id="dd96e-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="dd96e-131">TabOutlining</span></span>
- <span data-ttu-id="dd96e-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="dd96e-132">TabConflicts</span></span>
- <span data-ttu-id="dd96e-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dd96e-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dd96e-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dd96e-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="dd96e-135">Excel</span><span class="sxs-lookup"><span data-stu-id="dd96e-135">Excel</span></span>

- <span data-ttu-id="dd96e-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dd96e-136">**TabHome**</span></span>
- <span data-ttu-id="dd96e-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dd96e-137">**TabInsert**</span></span>
- <span data-ttu-id="dd96e-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="dd96e-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="dd96e-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="dd96e-139">TabFormulas</span></span>
- <span data-ttu-id="dd96e-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="dd96e-140">**TabData**</span></span>
- <span data-ttu-id="dd96e-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="dd96e-141">**TabReview**</span></span>
- <span data-ttu-id="dd96e-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dd96e-142">**TabView**</span></span>
- <span data-ttu-id="dd96e-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dd96e-143">TabDeveloper</span></span>
- <span data-ttu-id="dd96e-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dd96e-144">TabAddIns</span></span>
- <span data-ttu-id="dd96e-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dd96e-145">TabPrintPreview</span></span>
- <span data-ttu-id="dd96e-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dd96e-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="dd96e-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dd96e-147">PowerPoint</span></span>

- <span data-ttu-id="dd96e-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dd96e-148">**TabHome**</span></span>
- <span data-ttu-id="dd96e-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dd96e-149">**TabInsert**</span></span>
- <span data-ttu-id="dd96e-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="dd96e-150">**TabDesign**</span></span>
- <span data-ttu-id="dd96e-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="dd96e-151">**TabTransitions**</span></span>
- <span data-ttu-id="dd96e-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="dd96e-152">**TabAnimations**</span></span>
- <span data-ttu-id="dd96e-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="dd96e-153">TabSlideShow</span></span>
- <span data-ttu-id="dd96e-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="dd96e-154">TabReview</span></span>
- <span data-ttu-id="dd96e-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dd96e-155">**TabView**</span></span>
- <span data-ttu-id="dd96e-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dd96e-156">TabDeveloper</span></span>
- <span data-ttu-id="dd96e-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dd96e-157">TabAddIns</span></span>
- <span data-ttu-id="dd96e-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="dd96e-158">TabPrintPreview</span></span>
- <span data-ttu-id="dd96e-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="dd96e-159">TabMerge</span></span>
- <span data-ttu-id="dd96e-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="dd96e-160">TabGrayscale</span></span>
- <span data-ttu-id="dd96e-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="dd96e-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="dd96e-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="dd96e-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="dd96e-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="dd96e-163">TabSlideMaster</span></span>
- <span data-ttu-id="dd96e-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="dd96e-164">TabHandoutMaster</span></span>
- <span data-ttu-id="dd96e-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="dd96e-165">TabNotesMaster</span></span>
- <span data-ttu-id="dd96e-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="dd96e-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="dd96e-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="dd96e-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="dd96e-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="dd96e-168">OneNote</span></span>

- <span data-ttu-id="dd96e-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="dd96e-169">**TabHome**</span></span>
- <span data-ttu-id="dd96e-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="dd96e-170">**TabInsert**</span></span>
- <span data-ttu-id="dd96e-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="dd96e-171">**TabView**</span></span>
- <span data-ttu-id="dd96e-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="dd96e-172">TabDeveloper</span></span>
- <span data-ttu-id="dd96e-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="dd96e-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="dd96e-174">Group</span><span class="sxs-lookup"><span data-stu-id="dd96e-174">Group</span></span>

<span data-ttu-id="dd96e-p104">タブの UI 拡張ポイントのグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性は必須であり、各 **id** 属性はマニフェスト内で一意でなければなりません。**id** は最大 125 文字の文字列です。「[Group 要素](group.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd96e-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="dd96e-179">OfficeTab の例</span><span class="sxs-lookup"><span data-stu-id="dd96e-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
