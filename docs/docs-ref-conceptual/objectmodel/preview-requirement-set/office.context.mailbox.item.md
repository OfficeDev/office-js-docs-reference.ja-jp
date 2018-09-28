
# <a name="item"></a><span data-ttu-id="b325e-101">item</span><span class="sxs-lookup"><span data-stu-id="b325e-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b325e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b325e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b325e-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-105">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-105">Requirements</span></span>

|<span data-ttu-id="b325e-106">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-106">Requirement</span></span>|<span data-ttu-id="b325e-107">値</span><span class="sxs-lookup"><span data-stu-id="b325e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-109">1.0</span></span>|
|[<span data-ttu-id="b325e-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="b325e-111">Restricted</span></span>|
|[<span data-ttu-id="b325e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b325e-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-114">Members and methods</span></span>

| <span data-ttu-id="b325e-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-115">Member</span></span> | <span data-ttu-id="b325e-116">種類</span><span class="sxs-lookup"><span data-stu-id="b325e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b325e-117">attachments</span><span class="sxs-lookup"><span data-stu-id="b325e-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="b325e-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-118">Member</span></span> |
| [<span data-ttu-id="b325e-119">bcc</span><span class="sxs-lookup"><span data-stu-id="b325e-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b325e-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-120">Member</span></span> |
| [<span data-ttu-id="b325e-121">body</span><span class="sxs-lookup"><span data-stu-id="b325e-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="b325e-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-122">Member</span></span> |
| [<span data-ttu-id="b325e-123">cc</span><span class="sxs-lookup"><span data-stu-id="b325e-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b325e-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-124">Member</span></span> |
| [<span data-ttu-id="b325e-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="b325e-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="b325e-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-126">Member</span></span> |
| [<span data-ttu-id="b325e-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="b325e-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="b325e-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-128">Member</span></span> |
| [<span data-ttu-id="b325e-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b325e-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="b325e-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-130">Member</span></span> |
| [<span data-ttu-id="b325e-131">end</span><span class="sxs-lookup"><span data-stu-id="b325e-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="b325e-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-132">Member</span></span> |
| [<span data-ttu-id="b325e-133">from</span><span class="sxs-lookup"><span data-stu-id="b325e-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="b325e-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-134">Member</span></span> |
| [<span data-ttu-id="b325e-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="b325e-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="b325e-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-136">Member</span></span> |
| [<span data-ttu-id="b325e-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="b325e-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="b325e-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-138">Member</span></span> |
| [<span data-ttu-id="b325e-139">itemId</span><span class="sxs-lookup"><span data-stu-id="b325e-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="b325e-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-140">Member</span></span> |
| [<span data-ttu-id="b325e-141">itemType</span><span class="sxs-lookup"><span data-stu-id="b325e-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="b325e-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-142">Member</span></span> |
| [<span data-ttu-id="b325e-143">location</span><span class="sxs-lookup"><span data-stu-id="b325e-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="b325e-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-144">Member</span></span> |
| [<span data-ttu-id="b325e-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="b325e-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="b325e-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-146">Member</span></span> |
| [<span data-ttu-id="b325e-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="b325e-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="b325e-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-148">Member</span></span> |
| [<span data-ttu-id="b325e-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b325e-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b325e-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-150">Member</span></span> |
| [<span data-ttu-id="b325e-151">organizer</span><span class="sxs-lookup"><span data-stu-id="b325e-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="b325e-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-152">Member</span></span> |
| [<span data-ttu-id="b325e-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="b325e-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="b325e-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-154">Member</span></span> |
| [<span data-ttu-id="b325e-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b325e-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b325e-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-156">Member</span></span> |
| [<span data-ttu-id="b325e-157">sender</span><span class="sxs-lookup"><span data-stu-id="b325e-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="b325e-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-158">Member</span></span> |
| [<span data-ttu-id="b325e-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="b325e-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="b325e-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-160">Member</span></span> |
| [<span data-ttu-id="b325e-161">start</span><span class="sxs-lookup"><span data-stu-id="b325e-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="b325e-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-162">Member</span></span> |
| [<span data-ttu-id="b325e-163">subject</span><span class="sxs-lookup"><span data-stu-id="b325e-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="b325e-164">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-164">Member</span></span> |
| [<span data-ttu-id="b325e-165">to</span><span class="sxs-lookup"><span data-stu-id="b325e-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b325e-166">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-166">Member</span></span> |
| [<span data-ttu-id="b325e-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="b325e-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-168">Method</span></span> |
| [<span data-ttu-id="b325e-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="b325e-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="b325e-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-170">Method</span></span> |
| [<span data-ttu-id="b325e-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b325e-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-172">Method</span></span> |
| [<span data-ttu-id="b325e-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="b325e-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-174">Method</span></span> |
| [<span data-ttu-id="b325e-175">close</span><span class="sxs-lookup"><span data-stu-id="b325e-175">close</span></span>](#close) | <span data-ttu-id="b325e-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-176">Method</span></span> |
| [<span data-ttu-id="b325e-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b325e-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="b325e-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-178">Method</span></span> |
| [<span data-ttu-id="b325e-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b325e-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="b325e-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-180">Method</span></span> |
| [<span data-ttu-id="b325e-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="b325e-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="b325e-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-182">Method</span></span> |
| [<span data-ttu-id="b325e-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b325e-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="b325e-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-184">Method</span></span> |
| [<span data-ttu-id="b325e-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b325e-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="b325e-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-186">Method</span></span> |
| [<span data-ttu-id="b325e-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="b325e-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-188">Method</span></span> |
| [<span data-ttu-id="b325e-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b325e-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="b325e-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-190">Method</span></span> |
| [<span data-ttu-id="b325e-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b325e-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="b325e-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-192">Method</span></span> |
| [<span data-ttu-id="b325e-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="b325e-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-194">Method</span></span> |
| [<span data-ttu-id="b325e-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="b325e-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="b325e-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-196">Method</span></span> |
| [<span data-ttu-id="b325e-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b325e-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="b325e-198">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-198">Method</span></span> |
| [<span data-ttu-id="b325e-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="b325e-200">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-200">Method</span></span> |
| [<span data-ttu-id="b325e-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="b325e-202">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-202">Method</span></span> |
| [<span data-ttu-id="b325e-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="b325e-204">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-204">Method</span></span> |
| [<span data-ttu-id="b325e-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b325e-206">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-206">Method</span></span> |
| [<span data-ttu-id="b325e-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="b325e-208">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-208">Method</span></span> |
| [<span data-ttu-id="b325e-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b325e-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="b325e-210">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="b325e-211">例</span><span class="sxs-lookup"><span data-stu-id="b325e-211">Example</span></span>

<span data-ttu-id="b325e-212">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b325e-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="b325e-213">メンバー</span><span class="sxs-lookup"><span data-stu-id="b325e-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="b325e-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b325e-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="b325e-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-217">ファイルの特定の種類は、潜在的なセキュリティの問題により、Outlook によってブロックされは返されません。</span><span class="sxs-lookup"><span data-stu-id="b325e-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b325e-218">詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b325e-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-219">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-219">Type:</span></span>

*   <span data-ttu-id="b325e-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b325e-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-221">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-221">Requirements</span></span>

|<span data-ttu-id="b325e-222">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-222">Requirement</span></span>|<span data-ttu-id="b325e-223">値</span><span class="sxs-lookup"><span data-stu-id="b325e-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-225">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-225">1.0</span></span>|
|[<span data-ttu-id="b325e-226">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-227">ReadItem</span></span>|
|[<span data-ttu-id="b325e-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-229">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-230">例</span><span class="sxs-lookup"><span data-stu-id="b325e-230">Example</span></span>

<span data-ttu-id="b325e-231">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="b325e-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b325e-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b325e-233">取得またはメッセージの bcc (ブラインド カーボン コピー) 受信者を更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b325e-234">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-235">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-235">Type:</span></span>

*   [<span data-ttu-id="b325e-236">Recipients</span><span class="sxs-lookup"><span data-stu-id="b325e-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b325e-237">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-237">Requirements</span></span>

|<span data-ttu-id="b325e-238">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-238">Requirement</span></span>|<span data-ttu-id="b325e-239">値</span><span class="sxs-lookup"><span data-stu-id="b325e-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-241">1.1</span><span class="sxs-lookup"><span data-stu-id="b325e-241">1.1</span></span>|
|[<span data-ttu-id="b325e-242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-243">ReadItem</span></span>|
|[<span data-ttu-id="b325e-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-245">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-246">例</span><span class="sxs-lookup"><span data-stu-id="b325e-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="b325e-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="b325e-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="b325e-248">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-249">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-249">Type:</span></span>

*   [<span data-ttu-id="b325e-250">Body</span><span class="sxs-lookup"><span data-stu-id="b325e-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="b325e-251">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-251">Requirements</span></span>

|<span data-ttu-id="b325e-252">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-252">Requirement</span></span>|<span data-ttu-id="b325e-253">値</span><span class="sxs-lookup"><span data-stu-id="b325e-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-254">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-255">1.1</span><span class="sxs-lookup"><span data-stu-id="b325e-255">1.1</span></span>|
|[<span data-ttu-id="b325e-256">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-257">ReadItem</span></span>|
|[<span data-ttu-id="b325e-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-259">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b325e-260">[cc]: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="b325e-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b325e-261">メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b325e-262">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b325e-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-263">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-263">Read mode</span></span>

<span data-ttu-id="b325e-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-266">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-266">Compose mode</span></span>

<span data-ttu-id="b325e-267">`cc`を`Recipients`オブジェクトを取得または、メッセージの**Cc**行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-268">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-268">Type:</span></span>

*   <span data-ttu-id="b325e-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-270">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-270">Requirements</span></span>

|<span data-ttu-id="b325e-271">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-271">Requirement</span></span>|<span data-ttu-id="b325e-272">値</span><span class="sxs-lookup"><span data-stu-id="b325e-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-274">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-274">1.0</span></span>|
|[<span data-ttu-id="b325e-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-276">ReadItem</span></span>|
|[<span data-ttu-id="b325e-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-278">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-279">例</span><span class="sxs-lookup"><span data-stu-id="b325e-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b325e-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b325e-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="b325e-281">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b325e-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="b325e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b325e-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-286">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-286">Type:</span></span>

*   <span data-ttu-id="b325e-287">String</span><span class="sxs-lookup"><span data-stu-id="b325e-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-288">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-288">Requirements</span></span>

|<span data-ttu-id="b325e-289">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-289">Requirement</span></span>|<span data-ttu-id="b325e-290">値</span><span class="sxs-lookup"><span data-stu-id="b325e-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-292">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-292">1.0</span></span>|
|[<span data-ttu-id="b325e-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-294">ReadItem</span></span>|
|[<span data-ttu-id="b325e-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-296">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b325e-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b325e-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="b325e-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-300">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-300">Type:</span></span>

*   <span data-ttu-id="b325e-301">日付</span><span class="sxs-lookup"><span data-stu-id="b325e-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-302">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-302">Requirements</span></span>

|<span data-ttu-id="b325e-303">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-303">Requirement</span></span>|<span data-ttu-id="b325e-304">値</span><span class="sxs-lookup"><span data-stu-id="b325e-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-305">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-306">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-306">1.0</span></span>|
|[<span data-ttu-id="b325e-307">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-308">ReadItem</span></span>|
|[<span data-ttu-id="b325e-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-310">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-311">例</span><span class="sxs-lookup"><span data-stu-id="b325e-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b325e-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b325e-312">dateTimeModified :Date</span></span>

<span data-ttu-id="b325e-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-315">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-316">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-316">Type:</span></span>

*   <span data-ttu-id="b325e-317">日付</span><span class="sxs-lookup"><span data-stu-id="b325e-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-318">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-318">Requirements</span></span>

|<span data-ttu-id="b325e-319">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-319">Requirement</span></span>|<span data-ttu-id="b325e-320">値</span><span class="sxs-lookup"><span data-stu-id="b325e-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-321">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-322">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-322">1.0</span></span>|
|[<span data-ttu-id="b325e-323">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-324">ReadItem</span></span>|
|[<span data-ttu-id="b325e-325">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-326">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-327">例</span><span class="sxs-lookup"><span data-stu-id="b325e-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="b325e-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b325e-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="b325e-329">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b325e-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-332">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-332">Read mode</span></span>

<span data-ttu-id="b325e-333">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-334">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-334">Compose mode</span></span>

<span data-ttu-id="b325e-335">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b325e-336">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b325e-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-337">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-337">Type:</span></span>

*   <span data-ttu-id="b325e-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b325e-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-339">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-339">Requirements</span></span>

|<span data-ttu-id="b325e-340">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-340">Requirement</span></span>|<span data-ttu-id="b325e-341">値</span><span class="sxs-lookup"><span data-stu-id="b325e-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-342">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-343">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-343">1.0</span></span>|
|[<span data-ttu-id="b325e-344">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-345">ReadItem</span></span>|
|[<span data-ttu-id="b325e-346">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-347">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-348">例</span><span class="sxs-lookup"><span data-stu-id="b325e-348">Example</span></span>

<span data-ttu-id="b325e-349">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="b325e-350">:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="b325e-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="b325e-351">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="b325e-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-354">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`from`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b325e-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-355">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-355">Read mode</span></span>

<span data-ttu-id="b325e-356">`from`を`EmailAddressDetails`オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="b325e-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="b325e-357">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-357">Compose mode</span></span>

<span data-ttu-id="b325e-358">`from`を`From`を取得するメソッドを提供するオブジェクト、値からです。</span><span class="sxs-lookup"><span data-stu-id="b325e-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b325e-359">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-359">Type:</span></span>

*   <span data-ttu-id="b325e-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="b325e-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-361">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-361">Requirements</span></span>

|<span data-ttu-id="b325e-362">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="b325e-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-364">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-364">1.0</span></span>|<span data-ttu-id="b325e-365">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-365">1.7</span></span>|
|[<span data-ttu-id="b325e-366">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-367">ReadItem</span></span>|<span data-ttu-id="b325e-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-370">Read</span><span class="sxs-lookup"><span data-stu-id="b325e-370">Read</span></span>|<span data-ttu-id="b325e-371">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b325e-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b325e-372">internetMessageId :String</span></span>

<span data-ttu-id="b325e-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-375">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-375">Type:</span></span>

*   <span data-ttu-id="b325e-376">String</span><span class="sxs-lookup"><span data-stu-id="b325e-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-377">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-377">Requirements</span></span>

|<span data-ttu-id="b325e-378">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-378">Requirement</span></span>|<span data-ttu-id="b325e-379">値</span><span class="sxs-lookup"><span data-stu-id="b325e-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-380">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-381">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-381">1.0</span></span>|
|[<span data-ttu-id="b325e-382">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-383">ReadItem</span></span>|
|[<span data-ttu-id="b325e-384">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-385">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-386">例</span><span class="sxs-lookup"><span data-stu-id="b325e-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b325e-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b325e-387">itemClass :String</span></span>

<span data-ttu-id="b325e-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b325e-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="b325e-392">種類</span><span class="sxs-lookup"><span data-stu-id="b325e-392">Type</span></span>|<span data-ttu-id="b325e-393">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-393">Description</span></span>|<span data-ttu-id="b325e-394">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="b325e-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="b325e-395">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="b325e-395">Appointment items</span></span>|<span data-ttu-id="b325e-396">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b325e-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="b325e-397">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="b325e-397">Message items</span></span>|<span data-ttu-id="b325e-398">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="b325e-399">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-400">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-400">Type:</span></span>

*   <span data-ttu-id="b325e-401">String</span><span class="sxs-lookup"><span data-stu-id="b325e-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-402">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-402">Requirements</span></span>

|<span data-ttu-id="b325e-403">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-403">Requirement</span></span>|<span data-ttu-id="b325e-404">値</span><span class="sxs-lookup"><span data-stu-id="b325e-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-405">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-406">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-406">1.0</span></span>|
|[<span data-ttu-id="b325e-407">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-408">ReadItem</span></span>|
|[<span data-ttu-id="b325e-409">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-410">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-411">例</span><span class="sxs-lookup"><span data-stu-id="b325e-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b325e-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b325e-412">(nullable) itemId :String</span></span>

<span data-ttu-id="b325e-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-415">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="b325e-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b325e-416">`itemId`プロパティは、Outlook のエントリ ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="b325e-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b325e-417">この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b325e-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b325e-418">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b325e-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b325e-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-421">種類:</span><span class="sxs-lookup"><span data-stu-id="b325e-421">Type:</span></span>

*   <span data-ttu-id="b325e-422">String</span><span class="sxs-lookup"><span data-stu-id="b325e-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-423">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-423">Requirements</span></span>

|<span data-ttu-id="b325e-424">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-424">Requirement</span></span>|<span data-ttu-id="b325e-425">値</span><span class="sxs-lookup"><span data-stu-id="b325e-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-427">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-427">1.0</span></span>|
|[<span data-ttu-id="b325e-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-429">ReadItem</span></span>|
|[<span data-ttu-id="b325e-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-431">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-432">例</span><span class="sxs-lookup"><span data-stu-id="b325e-432">Example</span></span>

<span data-ttu-id="b325e-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="b325e-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b325e-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b325e-436">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b325e-437">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="b325e-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-438">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-438">Type:</span></span>

*   [<span data-ttu-id="b325e-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b325e-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b325e-440">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-440">Requirements</span></span>

|<span data-ttu-id="b325e-441">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-441">Requirement</span></span>|<span data-ttu-id="b325e-442">値</span><span class="sxs-lookup"><span data-stu-id="b325e-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-443">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-444">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-444">1.0</span></span>|
|[<span data-ttu-id="b325e-445">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-446">ReadItem</span></span>|
|[<span data-ttu-id="b325e-447">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-448">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-449">例</span><span class="sxs-lookup"><span data-stu-id="b325e-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="b325e-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="b325e-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="b325e-451">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-452">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-452">Read mode</span></span>

<span data-ttu-id="b325e-453">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-454">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-454">Compose mode</span></span>

<span data-ttu-id="b325e-455">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-456">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-456">Type:</span></span>

*   <span data-ttu-id="b325e-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="b325e-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-458">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-458">Requirements</span></span>

|<span data-ttu-id="b325e-459">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-459">Requirement</span></span>|<span data-ttu-id="b325e-460">値</span><span class="sxs-lookup"><span data-stu-id="b325e-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-461">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-462">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-462">1.0</span></span>|
|[<span data-ttu-id="b325e-463">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-464">ReadItem</span></span>|
|[<span data-ttu-id="b325e-465">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-466">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-467">例</span><span class="sxs-lookup"><span data-stu-id="b325e-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b325e-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b325e-468">normalizedSubject :String</span></span>

<span data-ttu-id="b325e-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b325e-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-473">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-473">Type:</span></span>

*   <span data-ttu-id="b325e-474">String</span><span class="sxs-lookup"><span data-stu-id="b325e-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-475">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-475">Requirements</span></span>

|<span data-ttu-id="b325e-476">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-476">Requirement</span></span>|<span data-ttu-id="b325e-477">値</span><span class="sxs-lookup"><span data-stu-id="b325e-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-478">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-479">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-479">1.0</span></span>|
|[<span data-ttu-id="b325e-480">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-481">ReadItem</span></span>|
|[<span data-ttu-id="b325e-482">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-483">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-484">例</span><span class="sxs-lookup"><span data-stu-id="b325e-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="b325e-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b325e-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="b325e-486">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-487">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-487">Type:</span></span>

*   [<span data-ttu-id="b325e-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b325e-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b325e-489">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-489">Requirements</span></span>

|<span data-ttu-id="b325e-490">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-490">Requirement</span></span>|<span data-ttu-id="b325e-491">値</span><span class="sxs-lookup"><span data-stu-id="b325e-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-492">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-493">1.3</span><span class="sxs-lookup"><span data-stu-id="b325e-493">1.3</span></span>|
|[<span data-ttu-id="b325e-494">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-495">ReadItem</span></span>|
|[<span data-ttu-id="b325e-496">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-497">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b325e-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b325e-499">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b325e-500">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b325e-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-501">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-501">Read mode</span></span>

<span data-ttu-id="b325e-502">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-503">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-503">Compose mode</span></span>

<span data-ttu-id="b325e-504">`optionalAttendees`を`Recipients`オブジェクトを取得または省略可能な会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-505">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-505">Type:</span></span>

*   <span data-ttu-id="b325e-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-507">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-507">Requirements</span></span>

|<span data-ttu-id="b325e-508">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-508">Requirement</span></span>|<span data-ttu-id="b325e-509">値</span><span class="sxs-lookup"><span data-stu-id="b325e-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-510">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-511">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-511">1.0</span></span>|
|[<span data-ttu-id="b325e-512">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-513">ReadItem</span></span>|
|[<span data-ttu-id="b325e-514">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-515">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-516">例</span><span class="sxs-lookup"><span data-stu-id="b325e-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="b325e-517">オーガナイザー:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[オーガナイザー](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="b325e-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="b325e-518">指定した会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-519">Read mode</span></span>

<span data-ttu-id="b325e-520">`organizer`プロパティは、会議の開催者を表す[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-521">Compose mode</span></span>

<span data-ttu-id="b325e-522">`organizer`プロパティが開催者の値を取得するメソッドを提供する[構成内容変更](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-523">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-523">Type:</span></span>

*   <span data-ttu-id="b325e-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [オーガナイザー](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="b325e-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-525">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-525">Requirements</span></span>

|<span data-ttu-id="b325e-526">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="b325e-527">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-528">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-528">1.0</span></span>|<span data-ttu-id="b325e-529">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-529">1.7</span></span>|
|[<span data-ttu-id="b325e-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-531">ReadItem</span></span>|<span data-ttu-id="b325e-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-533">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-534">Read</span><span class="sxs-lookup"><span data-stu-id="b325e-534">Read</span></span>|<span data-ttu-id="b325e-535">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-536">例</span><span class="sxs-lookup"><span data-stu-id="b325e-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="b325e-537">(許容) 定期的:[定期的なアイテム](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="b325e-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="b325e-538">取得または予定の定期的なパターンを設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="b325e-539">定期的な会議出席依頼を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="b325e-540">モードの予定表アイテムを読んだり作成したりします。</span><span class="sxs-lookup"><span data-stu-id="b325e-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="b325e-541">会議出席依頼アイテムの読み取りモードです。</span><span class="sxs-lookup"><span data-stu-id="b325e-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="b325e-542">`recurrence`プロパティは、アイテムが系列または系列のインスタンスである場合に定期的な予定または会議出席依頼に[定期的なアイテム](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="b325e-543">`null`単独の予定および会議出席依頼を単独の予定が返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="b325e-544">`undefined`会議出席依頼ではないメッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="b325e-545">注: 会議出席依頼がある、 `itemClass` IPM の値です。Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="b325e-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="b325e-546">注: 定期的なアイテム オブジェクトがある場合`null`、これは、オブジェクトが 1 つの予定または会議出席依頼、単独の予定および一連の一部ではないのであることを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-547">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-547">Type:</span></span>

* [<span data-ttu-id="b325e-548">定期的なアイテム</span><span class="sxs-lookup"><span data-stu-id="b325e-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="b325e-549">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-549">Requirement</span></span>|<span data-ttu-id="b325e-550">値</span><span class="sxs-lookup"><span data-stu-id="b325e-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-551">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-552">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-552">1.7</span></span>|
|[<span data-ttu-id="b325e-553">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-554">ReadItem</span></span>|
|[<span data-ttu-id="b325e-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-556">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b325e-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b325e-558">イベントの出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b325e-559">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b325e-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-560">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-560">Read mode</span></span>

<span data-ttu-id="b325e-561">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-562">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-562">Compose mode</span></span>

<span data-ttu-id="b325e-563">`requiredAttendees`を`Recipients`オブジェクトを取得または会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-564">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-564">Type:</span></span>

*   <span data-ttu-id="b325e-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-566">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-566">Requirements</span></span>

|<span data-ttu-id="b325e-567">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-567">Requirement</span></span>|<span data-ttu-id="b325e-568">値</span><span class="sxs-lookup"><span data-stu-id="b325e-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-569">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-570">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-570">1.0</span></span>|
|[<span data-ttu-id="b325e-571">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-572">ReadItem</span></span>|
|[<span data-ttu-id="b325e-573">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-574">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-575">例</span><span class="sxs-lookup"><span data-stu-id="b325e-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="b325e-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b325e-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="b325e-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b325e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b325e-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-581">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`sender`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b325e-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-582">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-582">Type:</span></span>

*   [<span data-ttu-id="b325e-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b325e-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b325e-584">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-584">Requirements</span></span>

|<span data-ttu-id="b325e-585">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-585">Requirement</span></span>|<span data-ttu-id="b325e-586">値</span><span class="sxs-lookup"><span data-stu-id="b325e-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-587">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-588">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-588">1.0</span></span>|
|[<span data-ttu-id="b325e-589">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-590">ReadItem</span></span>|
|[<span data-ttu-id="b325e-591">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-592">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-593">例</span><span class="sxs-lookup"><span data-stu-id="b325e-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="b325e-594">(許容) seriesId: 文字列</span><span class="sxs-lookup"><span data-stu-id="b325e-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="b325e-595">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="b325e-596">OWA と outlook 2002 で、`seriesId`は、この項目が属する親 (系列) アイテムの Exchange Web サービス (EWS) の ID を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="b325e-597">IOS および Android で、 `seriesId` 、親項目の残りの部分 ID を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-598">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="b325e-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b325e-599">`seriesId`プロパティは Outlook の REST API で使用される Outlook の Id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="b325e-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="b325e-600">この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b325e-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b325e-601">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b325e-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="b325e-602">`seriesId`プロパティを返します。`null`アイテムの親アイテムを次のようにされていない単一の関連するアイテム、予定または会議を要求し、返しますの`undefined`、その他の項目の要求を満たしていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-603">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-603">Type:</span></span>

* <span data-ttu-id="b325e-604">String</span><span class="sxs-lookup"><span data-stu-id="b325e-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-605">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-605">Requirements</span></span>

|<span data-ttu-id="b325e-606">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-606">Requirement</span></span>|<span data-ttu-id="b325e-607">値</span><span class="sxs-lookup"><span data-stu-id="b325e-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-609">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-609">1.7</span></span>|
|[<span data-ttu-id="b325e-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-611">ReadItem</span></span>|
|[<span data-ttu-id="b325e-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-613">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-614">例</span><span class="sxs-lookup"><span data-stu-id="b325e-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="b325e-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b325e-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="b325e-616">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b325e-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-619">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-619">Read mode</span></span>

<span data-ttu-id="b325e-620">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-621">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-621">Compose mode</span></span>

<span data-ttu-id="b325e-622">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b325e-623">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b325e-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-624">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-624">Type:</span></span>

*   <span data-ttu-id="b325e-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b325e-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-626">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-626">Requirements</span></span>

|<span data-ttu-id="b325e-627">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-627">Requirement</span></span>|<span data-ttu-id="b325e-628">値</span><span class="sxs-lookup"><span data-stu-id="b325e-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-629">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-630">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-630">1.0</span></span>|
|[<span data-ttu-id="b325e-631">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-632">ReadItem</span></span>|
|[<span data-ttu-id="b325e-633">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-634">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-635">例</span><span class="sxs-lookup"><span data-stu-id="b325e-635">Example</span></span>

<span data-ttu-id="b325e-636">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="b325e-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b325e-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="b325e-638">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b325e-639">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b325e-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-640">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-640">Read mode</span></span>

<span data-ttu-id="b325e-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b325e-643">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-643">Compose mode</span></span>

<span data-ttu-id="b325e-644">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b325e-645">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-645">Type:</span></span>

*   <span data-ttu-id="b325e-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b325e-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-647">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-647">Requirements</span></span>

|<span data-ttu-id="b325e-648">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-648">Requirement</span></span>|<span data-ttu-id="b325e-649">値</span><span class="sxs-lookup"><span data-stu-id="b325e-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-651">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-651">1.0</span></span>|
|[<span data-ttu-id="b325e-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-653">ReadItem</span></span>|
|[<span data-ttu-id="b325e-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-655">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b325e-656">: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="b325e-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b325e-657">[メッセージの [**宛先**] 行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b325e-658">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b325e-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b325e-659">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b325e-659">Read mode</span></span>

<span data-ttu-id="b325e-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b325e-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b325e-662">Compose mode</span></span>

<span data-ttu-id="b325e-663">`to`を`Recipients`オブジェクトを取得または、メッセージの [**宛先**] 行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b325e-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b325e-664">型:</span><span class="sxs-lookup"><span data-stu-id="b325e-664">Type:</span></span>

*   <span data-ttu-id="b325e-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b325e-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-666">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-666">Requirements</span></span>

|<span data-ttu-id="b325e-667">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-667">Requirement</span></span>|<span data-ttu-id="b325e-668">値</span><span class="sxs-lookup"><span data-stu-id="b325e-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-670">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-670">1.0</span></span>|
|[<span data-ttu-id="b325e-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-672">ReadItem</span></span>|
|[<span data-ttu-id="b325e-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-674">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-675">例</span><span class="sxs-lookup"><span data-stu-id="b325e-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b325e-676">メソッド</span><span class="sxs-lookup"><span data-stu-id="b325e-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b325e-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b325e-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b325e-678">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="b325e-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b325e-679">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="b325e-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b325e-680">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-681">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-681">Parameters:</span></span>
|<span data-ttu-id="b325e-682">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-682">Name</span></span>|<span data-ttu-id="b325e-683">型</span><span class="sxs-lookup"><span data-stu-id="b325e-683">Type</span></span>|<span data-ttu-id="b325e-684">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-684">Attributes</span></span>|<span data-ttu-id="b325e-685">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="b325e-686">String</span><span class="sxs-lookup"><span data-stu-id="b325e-686">String</span></span>||<span data-ttu-id="b325e-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="b325e-689">String</span><span class="sxs-lookup"><span data-stu-id="b325e-689">String</span></span>||<span data-ttu-id="b325e-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b325e-692">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-692">Object</span></span>|<span data-ttu-id="b325e-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-693">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-694">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-695">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-695">Object</span></span>|<span data-ttu-id="b325e-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-696">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-697">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="b325e-698">Boolean</span><span class="sxs-lookup"><span data-stu-id="b325e-698">Boolean</span></span>|<span data-ttu-id="b325e-699">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-699">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-700">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="b325e-701">function</span><span class="sxs-lookup"><span data-stu-id="b325e-701">function</span></span>|<span data-ttu-id="b325e-702">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-702">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-703">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b325e-704">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b325e-705">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b325e-706">エラー</span><span class="sxs-lookup"><span data-stu-id="b325e-706">Errors</span></span>

|<span data-ttu-id="b325e-707">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b325e-707">Error code</span></span>|<span data-ttu-id="b325e-708">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="b325e-709">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="b325e-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="b325e-710">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="b325e-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b325e-711">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b325e-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-712">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-712">Requirements</span></span>

|<span data-ttu-id="b325e-713">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-713">Requirement</span></span>|<span data-ttu-id="b325e-714">値</span><span class="sxs-lookup"><span data-stu-id="b325e-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-715">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-716">1.1</span><span class="sxs-lookup"><span data-stu-id="b325e-716">1.1</span></span>|
|[<span data-ttu-id="b325e-717">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-719">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-720">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b325e-721">例</span><span class="sxs-lookup"><span data-stu-id="b325e-721">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="b325e-722">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="b325e-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="b325e-723">addFileAttachmentFromBase64Async (base64File、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="b325e-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b325e-724">メッセージまたは予定を添付ファイルとしてエンコード base64 からファイルを追加します。</span><span class="sxs-lookup"><span data-stu-id="b325e-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b325e-725">`addFileAttachmentFromBase64Async`メソッドは、base64 エンコーディングからファイルをアップロードし、作成フォーム内の項目にアタッチします。</span><span class="sxs-lookup"><span data-stu-id="b325e-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="b325e-726">このメソッドは、AsyncResult.value オブジェクトの添付ファイルの識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="b325e-727">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-728">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-728">Parameters:</span></span>
|<span data-ttu-id="b325e-729">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-729">Name</span></span>|<span data-ttu-id="b325e-730">型</span><span class="sxs-lookup"><span data-stu-id="b325e-730">Type</span></span>|<span data-ttu-id="b325e-731">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-731">Attributes</span></span>|<span data-ttu-id="b325e-732">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="b325e-733">String</span><span class="sxs-lookup"><span data-stu-id="b325e-733">String</span></span>||<span data-ttu-id="b325e-734">イメージや、電子メール、またはイベントに追加するファイルのコンテンツを base64 にエンコードされます。</span><span class="sxs-lookup"><span data-stu-id="b325e-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="b325e-735">String</span><span class="sxs-lookup"><span data-stu-id="b325e-735">String</span></span>||<span data-ttu-id="b325e-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b325e-738">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-738">Object</span></span>|<span data-ttu-id="b325e-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-739">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-740">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-741">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-741">Object</span></span>|<span data-ttu-id="b325e-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-742">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-743">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="b325e-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="b325e-744">Boolean</span></span>|<span data-ttu-id="b325e-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-745">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-746">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="b325e-747">function</span><span class="sxs-lookup"><span data-stu-id="b325e-747">function</span></span>|<span data-ttu-id="b325e-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-748">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-749">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b325e-750">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b325e-751">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b325e-752">エラー</span><span class="sxs-lookup"><span data-stu-id="b325e-752">Errors</span></span>

|<span data-ttu-id="b325e-753">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b325e-753">Error code</span></span>|<span data-ttu-id="b325e-754">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="b325e-755">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="b325e-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="b325e-756">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="b325e-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b325e-757">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b325e-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-758">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-758">Requirements</span></span>

|<span data-ttu-id="b325e-759">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-759">Requirement</span></span>|<span data-ttu-id="b325e-760">値</span><span class="sxs-lookup"><span data-stu-id="b325e-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-761">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-762">プレビュー</span><span class="sxs-lookup"><span data-stu-id="b325e-762">Preview</span></span>|
|[<span data-ttu-id="b325e-763">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-765">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-766">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b325e-767">例</span><span class="sxs-lookup"><span data-stu-id="b325e-767">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b325e-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b325e-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b325e-769">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="b325e-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="b325e-770">現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="b325e-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-771">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-771">Parameters:</span></span>

| <span data-ttu-id="b325e-772">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-772">Name</span></span> | <span data-ttu-id="b325e-773">型</span><span class="sxs-lookup"><span data-stu-id="b325e-773">Type</span></span> | <span data-ttu-id="b325e-774">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-774">Attributes</span></span> | <span data-ttu-id="b325e-775">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b325e-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b325e-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b325e-777">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="b325e-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b325e-778">Function</span><span class="sxs-lookup"><span data-stu-id="b325e-778">Function</span></span> || <span data-ttu-id="b325e-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b325e-782">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-782">Object</span></span> | <span data-ttu-id="b325e-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-783">&lt;optional&gt;</span></span> | <span data-ttu-id="b325e-784">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b325e-785">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-785">Object</span></span> | <span data-ttu-id="b325e-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-786">&lt;optional&gt;</span></span> | <span data-ttu-id="b325e-787">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b325e-788">function</span><span class="sxs-lookup"><span data-stu-id="b325e-788">function</span></span>| <span data-ttu-id="b325e-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-789">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-790">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-791">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-791">Requirements</span></span>

|<span data-ttu-id="b325e-792">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-792">Requirement</span></span>| <span data-ttu-id="b325e-793">値</span><span class="sxs-lookup"><span data-stu-id="b325e-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-794">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b325e-795">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-795">1.7</span></span> |
|[<span data-ttu-id="b325e-796">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b325e-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-797">ReadItem</span></span> |
|[<span data-ttu-id="b325e-798">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b325e-799">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b325e-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b325e-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b325e-801">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="b325e-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b325e-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b325e-805">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b325e-806">Office アドインは、Outlook Web App で実行されている場合、`addItemAttachmentAsync`メソッドが項目を編集しているアイテム以外のアイテムに関連付けることができますただし、これはサポートされていません、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="b325e-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-807">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-807">Parameters:</span></span>

|<span data-ttu-id="b325e-808">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-808">Name</span></span>|<span data-ttu-id="b325e-809">型</span><span class="sxs-lookup"><span data-stu-id="b325e-809">Type</span></span>|<span data-ttu-id="b325e-810">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-810">Attributes</span></span>|<span data-ttu-id="b325e-811">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="b325e-812">String</span><span class="sxs-lookup"><span data-stu-id="b325e-812">String</span></span>||<span data-ttu-id="b325e-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="b325e-815">String</span><span class="sxs-lookup"><span data-stu-id="b325e-815">String</span></span>||<span data-ttu-id="b325e-p141">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b325e-818">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-818">Object</span></span>|<span data-ttu-id="b325e-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-819">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-820">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-821">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-821">Object</span></span>|<span data-ttu-id="b325e-822">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-822">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-823">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-824">function</span><span class="sxs-lookup"><span data-stu-id="b325e-824">function</span></span>|<span data-ttu-id="b325e-825">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-825">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-826">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b325e-827">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b325e-828">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b325e-829">エラー</span><span class="sxs-lookup"><span data-stu-id="b325e-829">Errors</span></span>

|<span data-ttu-id="b325e-830">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b325e-830">Error code</span></span>|<span data-ttu-id="b325e-831">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b325e-832">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b325e-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-833">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-833">Requirements</span></span>

|<span data-ttu-id="b325e-834">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-834">Requirement</span></span>|<span data-ttu-id="b325e-835">値</span><span class="sxs-lookup"><span data-stu-id="b325e-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-836">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-837">1.1</span><span class="sxs-lookup"><span data-stu-id="b325e-837">1.1</span></span>|
|[<span data-ttu-id="b325e-838">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-840">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-841">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-842">例</span><span class="sxs-lookup"><span data-stu-id="b325e-842">Example</span></span>

<span data-ttu-id="b325e-843">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="b325e-844">close()</span><span class="sxs-lookup"><span data-stu-id="b325e-844">close()</span></span>

<span data-ttu-id="b325e-845">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="b325e-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b325e-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-848">アイテム予定は、以前保存されたを使用する場合は、web 上の Outlook で`saveAsync`を求めるメッセージを保存、破棄、または、キャンセル場合でも、変更が発生していないから、項目を保存します。</span><span class="sxs-lookup"><span data-stu-id="b325e-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b325e-849">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="b325e-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-850">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-850">Requirements</span></span>

|<span data-ttu-id="b325e-851">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-851">Requirement</span></span>|<span data-ttu-id="b325e-852">値</span><span class="sxs-lookup"><span data-stu-id="b325e-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-853">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-854">1.3</span><span class="sxs-lookup"><span data-stu-id="b325e-854">1.3</span></span>|
|[<span data-ttu-id="b325e-855">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-856">制限あり</span><span class="sxs-lookup"><span data-stu-id="b325e-856">Restricted</span></span>|
|[<span data-ttu-id="b325e-857">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-858">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b325e-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b325e-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b325e-860">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-861">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-862">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b325e-863">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="b325e-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b325e-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="b325e-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-867">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-867">Parameters:</span></span>

|<span data-ttu-id="b325e-868">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-868">Name</span></span>|<span data-ttu-id="b325e-869">型</span><span class="sxs-lookup"><span data-stu-id="b325e-869">Type</span></span>|<span data-ttu-id="b325e-870">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-870">Attributes</span></span>|<span data-ttu-id="b325e-871">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="b325e-872">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b325e-872">String &#124; Object</span></span>||<span data-ttu-id="b325e-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b325e-875">**または**</span><span class="sxs-lookup"><span data-stu-id="b325e-875">**OR**</span></span><br/><span data-ttu-id="b325e-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b325e-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="b325e-878">String</span><span class="sxs-lookup"><span data-stu-id="b325e-878">String</span></span>|<span data-ttu-id="b325e-879">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-879">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="b325e-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="b325e-883">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-883">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-884">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="b325e-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="b325e-885">String</span><span class="sxs-lookup"><span data-stu-id="b325e-885">String</span></span>||<span data-ttu-id="b325e-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="b325e-888">String</span><span class="sxs-lookup"><span data-stu-id="b325e-888">String</span></span>||<span data-ttu-id="b325e-889">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="b325e-890">String</span><span class="sxs-lookup"><span data-stu-id="b325e-890">String</span></span>||<span data-ttu-id="b325e-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="b325e-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="b325e-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="b325e-893">Boolean</span></span>||<span data-ttu-id="b325e-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="b325e-896">String</span><span class="sxs-lookup"><span data-stu-id="b325e-896">String</span></span>||<span data-ttu-id="b325e-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="b325e-900">function</span><span class="sxs-lookup"><span data-stu-id="b325e-900">function</span></span>|<span data-ttu-id="b325e-901">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-901">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-902">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-903">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-903">Requirements</span></span>

|<span data-ttu-id="b325e-904">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-904">Requirement</span></span>|<span data-ttu-id="b325e-905">値</span><span class="sxs-lookup"><span data-stu-id="b325e-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-907">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-907">1.0</span></span>|
|[<span data-ttu-id="b325e-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-909">ReadItem</span></span>|
|[<span data-ttu-id="b325e-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b325e-912">例</span><span class="sxs-lookup"><span data-stu-id="b325e-912">Examples</span></span>

<span data-ttu-id="b325e-913">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="b325e-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b325e-914">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b325e-915">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b325e-916">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-916">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="b325e-917">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-917">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="b325e-918">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b325e-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b325e-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="b325e-920">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-921">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-922">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b325e-923">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="b325e-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b325e-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="b325e-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-927">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-927">Parameters:</span></span>

|<span data-ttu-id="b325e-928">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-928">Name</span></span>|<span data-ttu-id="b325e-929">型</span><span class="sxs-lookup"><span data-stu-id="b325e-929">Type</span></span>|<span data-ttu-id="b325e-930">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-930">Attributes</span></span>|<span data-ttu-id="b325e-931">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="b325e-932">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b325e-932">String &#124; Object</span></span>||<span data-ttu-id="b325e-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b325e-935">**または**</span><span class="sxs-lookup"><span data-stu-id="b325e-935">**OR**</span></span><br/><span data-ttu-id="b325e-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b325e-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="b325e-938">String</span><span class="sxs-lookup"><span data-stu-id="b325e-938">String</span></span>|<span data-ttu-id="b325e-939">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-939">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b325e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="b325e-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="b325e-943">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-943">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-944">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="b325e-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="b325e-945">String</span><span class="sxs-lookup"><span data-stu-id="b325e-945">String</span></span>||<span data-ttu-id="b325e-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="b325e-948">String</span><span class="sxs-lookup"><span data-stu-id="b325e-948">String</span></span>||<span data-ttu-id="b325e-949">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="b325e-950">String</span><span class="sxs-lookup"><span data-stu-id="b325e-950">String</span></span>||<span data-ttu-id="b325e-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="b325e-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="b325e-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="b325e-953">Boolean</span></span>||<span data-ttu-id="b325e-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="b325e-956">String</span><span class="sxs-lookup"><span data-stu-id="b325e-956">String</span></span>||<span data-ttu-id="b325e-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="b325e-960">function</span><span class="sxs-lookup"><span data-stu-id="b325e-960">function</span></span>|<span data-ttu-id="b325e-961">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-961">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-962">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-963">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-963">Requirements</span></span>

|<span data-ttu-id="b325e-964">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-964">Requirement</span></span>|<span data-ttu-id="b325e-965">値</span><span class="sxs-lookup"><span data-stu-id="b325e-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-966">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-967">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-967">1.0</span></span>|
|[<span data-ttu-id="b325e-968">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-969">ReadItem</span></span>|
|[<span data-ttu-id="b325e-970">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-971">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b325e-972">例</span><span class="sxs-lookup"><span data-stu-id="b325e-972">Examples</span></span>

<span data-ttu-id="b325e-973">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="b325e-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b325e-974">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b325e-975">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b325e-976">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-976">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="b325e-977">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-977">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="b325e-978">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="b325e-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="b325e-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b325e-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="b325e-980">選択したアイテムの本文内のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-981">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-982">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-982">Requirements</span></span>

|<span data-ttu-id="b325e-983">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-983">Requirement</span></span>|<span data-ttu-id="b325e-984">値</span><span class="sxs-lookup"><span data-stu-id="b325e-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-985">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-986">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-986">1.0</span></span>|
|[<span data-ttu-id="b325e-987">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-988">ReadItem</span></span>|
|[<span data-ttu-id="b325e-989">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-990">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-991">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-991">Returns:</span></span>

<span data-ttu-id="b325e-992">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b325e-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b325e-993">例</span><span class="sxs-lookup"><span data-stu-id="b325e-993">Example</span></span>

<span data-ttu-id="b325e-994">次の使用例は、現在の項目の本文に連絡先のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="b325e-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b325e-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b325e-996">選択したアイテムの本文に指定されたエンティティ型のすべてのエンティティの配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-997">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-998">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-998">Parameters:</span></span>

|<span data-ttu-id="b325e-999">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-999">Name</span></span>|<span data-ttu-id="b325e-1000">種類</span><span class="sxs-lookup"><span data-stu-id="b325e-1000">Type</span></span>|<span data-ttu-id="b325e-1001">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="b325e-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b325e-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="b325e-1003">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="b325e-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1004">Requirements</span><span class="sxs-lookup"><span data-stu-id="b325e-1004">Requirements</span></span>

|<span data-ttu-id="b325e-1005">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1005">Requirement</span></span>|<span data-ttu-id="b325e-1006">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1007">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-1008">1.0</span></span>|
|[<span data-ttu-id="b325e-1009">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1010">制限あり</span><span class="sxs-lookup"><span data-stu-id="b325e-1010">Restricted</span></span>|
|[<span data-ttu-id="b325e-1011">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1012">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1013">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1013">Returns:</span></span>

<span data-ttu-id="b325e-1014">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b325e-1015">アイテムの本文に指定した型のエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b325e-1016">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="b325e-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b325e-1017">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="b325e-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="b325e-1018">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="b325e-1018">Value of `entityType`</span></span>|<span data-ttu-id="b325e-1019">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="b325e-1019">Type of objects in returned array</span></span>|<span data-ttu-id="b325e-1020">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="b325e-1021">文字列</span><span class="sxs-lookup"><span data-stu-id="b325e-1021">String</span></span>|<span data-ttu-id="b325e-1022">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b325e-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="b325e-1023">連絡先</span><span class="sxs-lookup"><span data-stu-id="b325e-1023">Contact</span></span>|<span data-ttu-id="b325e-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b325e-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="b325e-1025">文字列</span><span class="sxs-lookup"><span data-stu-id="b325e-1025">String</span></span>|<span data-ttu-id="b325e-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b325e-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="b325e-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b325e-1027">MeetingSuggestion</span></span>|<span data-ttu-id="b325e-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b325e-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="b325e-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b325e-1029">PhoneNumber</span></span>|<span data-ttu-id="b325e-1030">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b325e-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="b325e-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b325e-1031">TaskSuggestion</span></span>|<span data-ttu-id="b325e-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b325e-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="b325e-1033">文字列</span><span class="sxs-lookup"><span data-stu-id="b325e-1033">String</span></span>|<span data-ttu-id="b325e-1034">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b325e-1034">**Restricted**</span></span>|

<span data-ttu-id="b325e-1035">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b325e-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b325e-1036">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1036">Example</span></span>

<span data-ttu-id="b325e-1037">次の例では、現在の項目の本文に郵便番号のアドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="b325e-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b325e-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b325e-1039">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1040">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-1041">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1042">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1042">Parameters:</span></span>

|<span data-ttu-id="b325e-1043">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1043">Name</span></span>|<span data-ttu-id="b325e-1044">種類</span><span class="sxs-lookup"><span data-stu-id="b325e-1044">Type</span></span>|<span data-ttu-id="b325e-1045">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="b325e-1046">String</span><span class="sxs-lookup"><span data-stu-id="b325e-1046">String</span></span>|<span data-ttu-id="b325e-1047">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="b325e-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1048">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1048">Requirements</span></span>

|<span data-ttu-id="b325e-1049">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1049">Requirement</span></span>|<span data-ttu-id="b325e-1050">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1051">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-1052">1.0</span></span>|
|[<span data-ttu-id="b325e-1053">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1054">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1055">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1056">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1057">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1057">Returns:</span></span>

<span data-ttu-id="b325e-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b325e-1060">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b325e-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="b325e-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b325e-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="b325e-1062">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1063">このメソッドは Outlook 2016 または Windows (クイック実行バージョン 16.0.8413.1000 以降) と、web 上で Outlook を後で Office 365 のです。</span><span class="sxs-lookup"><span data-stu-id="b325e-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1064">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1064">Parameters:</span></span>
|<span data-ttu-id="b325e-1065">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1065">Name</span></span>|<span data-ttu-id="b325e-1066">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1066">Type</span></span>|<span data-ttu-id="b325e-1067">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1067">Attributes</span></span>|<span data-ttu-id="b325e-1068">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b325e-1069">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b325e-1069">Object</span></span>|<span data-ttu-id="b325e-1070">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1071">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1072">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1072">Object</span></span>|<span data-ttu-id="b325e-1073">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1074">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-1075">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1075">function</span></span>|<span data-ttu-id="b325e-1076">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1077">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b325e-1078">成功した場合、初期化データが提供されている、`asyncResult.value`文字列としてのプロパティです。</span><span class="sxs-lookup"><span data-stu-id="b325e-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="b325e-1079">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1080">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1080">Requirements</span></span>

|<span data-ttu-id="b325e-1081">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1081">Requirement</span></span>|<span data-ttu-id="b325e-1082">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1083">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1084">プレビュー</span><span class="sxs-lookup"><span data-stu-id="b325e-1084">Preview</span></span>|
|[<span data-ttu-id="b325e-1085">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1086">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1087">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1088">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-1089">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1089">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="b325e-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b325e-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b325e-1091">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1092">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b325e-1096">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="b325e-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b325e-1097">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b325e-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b325e-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-1101">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1101">Requirements</span></span>

|<span data-ttu-id="b325e-1102">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1102">Requirement</span></span>|<span data-ttu-id="b325e-1103">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1104">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-1105">1.0</span></span>|
|[<span data-ttu-id="b325e-1106">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1107">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1108">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1109">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1110">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1110">Returns:</span></span>

<span data-ttu-id="b325e-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="b325e-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b325e-1113">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b325e-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b325e-1114">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b325e-1115">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1115">Example</span></span>

<span data-ttu-id="b325e-1116">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b325e-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b325e-1117">getRegExMatchesByName(name)] → [(許容) {配列。 < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="b325e-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b325e-1118">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1119">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-1120">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="b325e-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b325e-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="b325e-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1123">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1123">Parameters:</span></span>

|<span data-ttu-id="b325e-1124">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1124">Name</span></span>|<span data-ttu-id="b325e-1125">種類</span><span class="sxs-lookup"><span data-stu-id="b325e-1125">Type</span></span>|<span data-ttu-id="b325e-1126">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="b325e-1127">String</span><span class="sxs-lookup"><span data-stu-id="b325e-1127">String</span></span>|<span data-ttu-id="b325e-1128">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="b325e-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1129">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1129">Requirements</span></span>

|<span data-ttu-id="b325e-1130">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1130">Requirement</span></span>|<span data-ttu-id="b325e-1131">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-1133">1.0</span></span>|
|[<span data-ttu-id="b325e-1134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1135">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1137">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1138">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1138">Returns:</span></span>

<span data-ttu-id="b325e-1139">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="b325e-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b325e-1140">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b325e-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b325e-1141">配列。 < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="b325e-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b325e-1142">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b325e-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b325e-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b325e-1144">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b325e-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1147">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1147">Parameters:</span></span>

|<span data-ttu-id="b325e-1148">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1148">Name</span></span>|<span data-ttu-id="b325e-1149">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1149">Type</span></span>|<span data-ttu-id="b325e-1150">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1150">Attributes</span></span>|<span data-ttu-id="b325e-1151">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="b325e-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b325e-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b325e-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="b325e-1156">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1156">Object</span></span>|<span data-ttu-id="b325e-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1158">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1159">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1159">Object</span></span>|<span data-ttu-id="b325e-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1161">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-1162">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1162">function</span></span>||<span data-ttu-id="b325e-1163">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b325e-1164">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b325e-1165">選択範囲は、source プロパティにアクセスするには、呼び出す`asyncResult.value.sourceProperty`、いずれかの方法となる`body`または`subject`。</span><span class="sxs-lookup"><span data-stu-id="b325e-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1166">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1166">Requirements</span></span>

|<span data-ttu-id="b325e-1167">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1167">Requirement</span></span>|<span data-ttu-id="b325e-1168">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="b325e-1170">1.2</span></span>|
|[<span data-ttu-id="b325e-1171">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-1173">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1174">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1175">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1175">Returns:</span></span>

<span data-ttu-id="b325e-1176">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="b325e-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b325e-1177">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b325e-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b325e-1178">String</span><span class="sxs-lookup"><span data-stu-id="b325e-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b325e-1179">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1179">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="b325e-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b325e-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="b325e-p168">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1183">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-1184">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1184">Requirements</span></span>

|<span data-ttu-id="b325e-1185">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1185">Requirement</span></span>|<span data-ttu-id="b325e-1186">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="b325e-1188">1.6</span></span>|
|[<span data-ttu-id="b325e-1189">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1190">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1192">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1193">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1193">Returns:</span></span>

<span data-ttu-id="b325e-1194">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b325e-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b325e-1195">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1195">Example</span></span>

<span data-ttu-id="b325e-1196">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="b325e-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="b325e-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b325e-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="b325e-p169">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1200">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b325e-p170">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b325e-1204">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="b325e-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b325e-1205">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b325e-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b325e-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b325e-1209">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1209">Requirements</span></span>

|<span data-ttu-id="b325e-1210">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1210">Requirement</span></span>|<span data-ttu-id="b325e-1211">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="b325e-1213">1.6</span></span>|
|[<span data-ttu-id="b325e-1214">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1215">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1217">読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b325e-1218">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b325e-1218">Returns:</span></span>

<span data-ttu-id="b325e-p172">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="b325e-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="b325e-1221">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1221">Example</span></span>

<span data-ttu-id="b325e-1222">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b325e-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="b325e-1223">getSharedPropertiesAsync ([オプション] では、コールバック)</span><span class="sxs-lookup"><span data-stu-id="b325e-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="b325e-1224">共有フォルダー、予定表、またはメールボックス内の選択されている予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1225">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1225">Parameters:</span></span>

|<span data-ttu-id="b325e-1226">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1226">Name</span></span>|<span data-ttu-id="b325e-1227">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1227">Type</span></span>|<span data-ttu-id="b325e-1228">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1228">Attributes</span></span>|<span data-ttu-id="b325e-1229">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b325e-1230">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b325e-1230">Object</span></span>|<span data-ttu-id="b325e-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1232">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1233">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1233">Object</span></span>|<span data-ttu-id="b325e-1234">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1235">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-1236">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1236">function</span></span>||<span data-ttu-id="b325e-1237">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b325e-1238">共有のプロパティはそのまま、[`SharedProperties`](/javascript/api/outlook/office.sharedproperties)オブジェクトで、`asyncResult.value`プロパティ。</span><span class="sxs-lookup"><span data-stu-id="b325e-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b325e-1239">このオブジェクトは、アイテムの共有のプロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1240">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1240">Requirements</span></span>

|<span data-ttu-id="b325e-1241">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1241">Requirement</span></span>|<span data-ttu-id="b325e-1242">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1244">プレビュー</span><span class="sxs-lookup"><span data-stu-id="b325e-1244">Preview</span></span>|
|[<span data-ttu-id="b325e-1245">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1246">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1247">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1248">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-1249">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b325e-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b325e-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b325e-1251">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b325e-p174">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="b325e-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1255">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1255">Parameters:</span></span>

|<span data-ttu-id="b325e-1256">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1256">Name</span></span>|<span data-ttu-id="b325e-1257">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1257">Type</span></span>|<span data-ttu-id="b325e-1258">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1258">Attributes</span></span>|<span data-ttu-id="b325e-1259">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="b325e-1260">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1260">function</span></span>||<span data-ttu-id="b325e-1261">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b325e-1262">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b325e-1263">取得し、アイテムのカスタム プロパティを削除してサーバーにバックアップを設定するカスタム プロパティに対する変更を保存するのには、このオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="b325e-1264">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b325e-1264">Object</span></span>|<span data-ttu-id="b325e-1265">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1266">開発者は、コールバック関数にアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b325e-1267">によってこのオブジェクトにアクセスできる、`asyncResult.asyncContext`コールバック関数のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="b325e-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1268">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1268">Requirements</span></span>

|<span data-ttu-id="b325e-1269">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1269">Requirement</span></span>|<span data-ttu-id="b325e-1270">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="b325e-1272">1.0</span></span>|
|[<span data-ttu-id="b325e-1273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1274">ReadItem</span></span>|
|[<span data-ttu-id="b325e-1275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1276">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-1277">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1277">Example</span></span>

<span data-ttu-id="b325e-p177">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b325e-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b325e-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b325e-1282">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b325e-p178">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1287">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1287">Parameters:</span></span>

|<span data-ttu-id="b325e-1288">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1288">Name</span></span>|<span data-ttu-id="b325e-1289">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1289">Type</span></span>|<span data-ttu-id="b325e-1290">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1290">Attributes</span></span>|<span data-ttu-id="b325e-1291">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="b325e-1292">String</span><span class="sxs-lookup"><span data-stu-id="b325e-1292">String</span></span>||<span data-ttu-id="b325e-p179">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="b325e-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="b325e-1295">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1295">Object</span></span>|<span data-ttu-id="b325e-1296">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1297">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1298">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1298">Object</span></span>|<span data-ttu-id="b325e-1299">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1300">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-1301">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1301">function</span></span>|<span data-ttu-id="b325e-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1303">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b325e-1304">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b325e-1305">エラー</span><span class="sxs-lookup"><span data-stu-id="b325e-1305">Errors</span></span>

|<span data-ttu-id="b325e-1306">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b325e-1306">Error code</span></span>|<span data-ttu-id="b325e-1307">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="b325e-1308">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="b325e-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1309">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1309">Requirements</span></span>

|<span data-ttu-id="b325e-1310">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1310">Requirement</span></span>|<span data-ttu-id="b325e-1311">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="b325e-1313">1.1</span></span>|
|[<span data-ttu-id="b325e-1314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-1316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1317">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-1318">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1318">Example</span></span>

<span data-ttu-id="b325e-1319">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1319">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b325e-1320">removeHandlerAsync (イベントの種類、ハンドラー、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="b325e-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b325e-1321">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="b325e-1322">現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="b325e-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1323">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1323">Parameters:</span></span>

| <span data-ttu-id="b325e-1324">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1324">Name</span></span> | <span data-ttu-id="b325e-1325">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1325">Type</span></span> | <span data-ttu-id="b325e-1326">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1326">Attributes</span></span> | <span data-ttu-id="b325e-1327">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b325e-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b325e-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b325e-1329">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="b325e-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b325e-1330">Function</span><span class="sxs-lookup"><span data-stu-id="b325e-1330">Function</span></span> || <span data-ttu-id="b325e-p180">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="b325e-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b325e-1334">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1334">Object</span></span> | <span data-ttu-id="b325e-1335">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="b325e-1336">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b325e-1337">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1337">Object</span></span> | <span data-ttu-id="b325e-1338">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="b325e-1339">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b325e-1340">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1340">function</span></span>| <span data-ttu-id="b325e-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1342">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1343">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1343">Requirements</span></span>

|<span data-ttu-id="b325e-1344">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1344">Requirement</span></span>| <span data-ttu-id="b325e-1345">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b325e-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="b325e-1347">1.7</span></span> |
|[<span data-ttu-id="b325e-1348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b325e-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1349">ReadItem</span></span> |
|[<span data-ttu-id="b325e-1350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b325e-1351">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b325e-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b325e-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b325e-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="b325e-1353">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="b325e-p181">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1357">アドインを呼び出す場合は、`saveAsync`内のアイテムの作成モードを取得するのには、 `itemId` EWS または REST API を使用するにすると、Outlook キャッシュ モードでは、かかる場合がある項目が実際には、サーバーと同期をとる前にいくつかの時間に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b325e-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b325e-1358">使用して、項目が同期されるまで、`itemId`エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b325e-p183">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b325e-1362">次のクライアントのさまざまな問題のある`saveAsync`の予定の作成モード。</span><span class="sxs-lookup"><span data-stu-id="b325e-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b325e-1363">Mac の Outlook をサポートしていない`saveAsync`での会議では、作成モードです。</span><span class="sxs-lookup"><span data-stu-id="b325e-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b325e-1364">呼び出す`saveAsync`Mac の Outlook で会議のエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b325e-1365">Web 上で outlook が常に招待状を送信または更新する場合`saveAsync`予定で作成モードです。</span><span class="sxs-lookup"><span data-stu-id="b325e-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1366">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1366">Parameters:</span></span>

|<span data-ttu-id="b325e-1367">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1367">Name</span></span>|<span data-ttu-id="b325e-1368">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1368">Type</span></span>|<span data-ttu-id="b325e-1369">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1369">Attributes</span></span>|<span data-ttu-id="b325e-1370">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b325e-1371">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b325e-1371">Object</span></span>|<span data-ttu-id="b325e-1372">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1373">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1374">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1374">Object</span></span>|<span data-ttu-id="b325e-1375">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1376">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b325e-1377">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1377">function</span></span>||<span data-ttu-id="b325e-1378">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b325e-1379">成功した場合、項目の識別子が提供されている、`asyncResult.value`プロパティ。</span><span class="sxs-lookup"><span data-stu-id="b325e-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1380">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1380">Requirements</span></span>

|<span data-ttu-id="b325e-1381">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1381">Requirement</span></span>|<span data-ttu-id="b325e-1382">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1383">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="b325e-1384">1.3</span></span>|
|[<span data-ttu-id="b325e-1385">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-1387">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1388">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b325e-1389">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b325e-p185">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b325e-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b325e-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b325e-1393">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="b325e-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b325e-p186">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b325e-1397">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b325e-1397">Parameters:</span></span>

|<span data-ttu-id="b325e-1398">名前</span><span class="sxs-lookup"><span data-stu-id="b325e-1398">Name</span></span>|<span data-ttu-id="b325e-1399">型</span><span class="sxs-lookup"><span data-stu-id="b325e-1399">Type</span></span>|<span data-ttu-id="b325e-1400">属性</span><span class="sxs-lookup"><span data-stu-id="b325e-1400">Attributes</span></span>|<span data-ttu-id="b325e-1401">説明</span><span class="sxs-lookup"><span data-stu-id="b325e-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="b325e-1402">String</span><span class="sxs-lookup"><span data-stu-id="b325e-1402">String</span></span>||<span data-ttu-id="b325e-p187">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="b325e-1406">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1406">Object</span></span>|<span data-ttu-id="b325e-1407">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1408">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b325e-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b325e-1409">Object</span><span class="sxs-lookup"><span data-stu-id="b325e-1409">Object</span></span>|<span data-ttu-id="b325e-1410">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-1411">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="b325e-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b325e-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="b325e-1413">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b325e-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="b325e-p188">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b325e-p189">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b325e-1418">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="b325e-1419">function</span><span class="sxs-lookup"><span data-stu-id="b325e-1419">function</span></span>||<span data-ttu-id="b325e-1420">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b325e-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b325e-1421">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1421">Requirements</span></span>

|<span data-ttu-id="b325e-1422">要件</span><span class="sxs-lookup"><span data-stu-id="b325e-1422">Requirement</span></span>|<span data-ttu-id="b325e-1423">値</span><span class="sxs-lookup"><span data-stu-id="b325e-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="b325e-1424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b325e-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b325e-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="b325e-1425">1.2</span></span>|
|[<span data-ttu-id="b325e-1426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b325e-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b325e-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b325e-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="b325e-1428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b325e-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b325e-1429">作成</span><span class="sxs-lookup"><span data-stu-id="b325e-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b325e-1430">例</span><span class="sxs-lookup"><span data-stu-id="b325e-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```