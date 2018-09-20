
# <a name="item"></a><span data-ttu-id="e6778-101">item</span><span class="sxs-lookup"><span data-stu-id="e6778-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e6778-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e6778-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e6778-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-105">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-105">Requirements</span></span>

|<span data-ttu-id="e6778-106">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-106">Requirement</span></span>|<span data-ttu-id="e6778-107">値</span><span class="sxs-lookup"><span data-stu-id="e6778-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-109">1.0</span></span>|
|[<span data-ttu-id="e6778-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="e6778-111">Restricted</span></span>|
|[<span data-ttu-id="e6778-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e6778-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-114">Members and methods</span></span>

| <span data-ttu-id="e6778-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-115">Member</span></span> | <span data-ttu-id="e6778-116">種類</span><span class="sxs-lookup"><span data-stu-id="e6778-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e6778-117">attachments</span><span class="sxs-lookup"><span data-stu-id="e6778-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="e6778-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-118">Member</span></span> |
| [<span data-ttu-id="e6778-119">bcc</span><span class="sxs-lookup"><span data-stu-id="e6778-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e6778-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-120">Member</span></span> |
| [<span data-ttu-id="e6778-121">body</span><span class="sxs-lookup"><span data-stu-id="e6778-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="e6778-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-122">Member</span></span> |
| [<span data-ttu-id="e6778-123">cc</span><span class="sxs-lookup"><span data-stu-id="e6778-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e6778-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-124">Member</span></span> |
| [<span data-ttu-id="e6778-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="e6778-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e6778-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-126">Member</span></span> |
| [<span data-ttu-id="e6778-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e6778-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e6778-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-128">Member</span></span> |
| [<span data-ttu-id="e6778-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e6778-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e6778-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-130">Member</span></span> |
| [<span data-ttu-id="e6778-131">end</span><span class="sxs-lookup"><span data-stu-id="e6778-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="e6778-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-132">Member</span></span> |
| [<span data-ttu-id="e6778-133">from</span><span class="sxs-lookup"><span data-stu-id="e6778-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="e6778-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-134">Member</span></span> |
| [<span data-ttu-id="e6778-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e6778-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e6778-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-136">Member</span></span> |
| [<span data-ttu-id="e6778-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="e6778-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e6778-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-138">Member</span></span> |
| [<span data-ttu-id="e6778-139">itemId</span><span class="sxs-lookup"><span data-stu-id="e6778-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e6778-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-140">Member</span></span> |
| [<span data-ttu-id="e6778-141">itemType</span><span class="sxs-lookup"><span data-stu-id="e6778-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="e6778-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-142">Member</span></span> |
| [<span data-ttu-id="e6778-143">location</span><span class="sxs-lookup"><span data-stu-id="e6778-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="e6778-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-144">Member</span></span> |
| [<span data-ttu-id="e6778-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e6778-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e6778-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-146">Member</span></span> |
| [<span data-ttu-id="e6778-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e6778-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="e6778-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-148">Member</span></span> |
| [<span data-ttu-id="e6778-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e6778-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e6778-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-150">Member</span></span> |
| [<span data-ttu-id="e6778-151">organizer</span><span class="sxs-lookup"><span data-stu-id="e6778-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="e6778-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-152">Member</span></span> |
| [<span data-ttu-id="e6778-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="e6778-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="e6778-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-154">Member</span></span> |
| [<span data-ttu-id="e6778-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e6778-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e6778-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-156">Member</span></span> |
| [<span data-ttu-id="e6778-157">sender</span><span class="sxs-lookup"><span data-stu-id="e6778-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="e6778-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-158">Member</span></span> |
| [<span data-ttu-id="e6778-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="e6778-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="e6778-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-160">Member</span></span> |
| [<span data-ttu-id="e6778-161">start</span><span class="sxs-lookup"><span data-stu-id="e6778-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="e6778-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-162">Member</span></span> |
| [<span data-ttu-id="e6778-163">subject</span><span class="sxs-lookup"><span data-stu-id="e6778-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="e6778-164">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-164">Member</span></span> |
| [<span data-ttu-id="e6778-165">to</span><span class="sxs-lookup"><span data-stu-id="e6778-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e6778-166">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-166">Member</span></span> |
| [<span data-ttu-id="e6778-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e6778-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-168">Method</span></span> |
| [<span data-ttu-id="e6778-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e6778-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="e6778-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-170">Method</span></span> |
| [<span data-ttu-id="e6778-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e6778-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-172">Method</span></span> |
| [<span data-ttu-id="e6778-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e6778-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-174">Method</span></span> |
| [<span data-ttu-id="e6778-175">close</span><span class="sxs-lookup"><span data-stu-id="e6778-175">close</span></span>](#close) | <span data-ttu-id="e6778-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-176">Method</span></span> |
| [<span data-ttu-id="e6778-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e6778-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="e6778-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-178">Method</span></span> |
| [<span data-ttu-id="e6778-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e6778-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="e6778-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-180">Method</span></span> |
| [<span data-ttu-id="e6778-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="e6778-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="e6778-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-182">Method</span></span> |
| [<span data-ttu-id="e6778-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e6778-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="e6778-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-184">Method</span></span> |
| [<span data-ttu-id="e6778-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e6778-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="e6778-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-186">Method</span></span> |
| [<span data-ttu-id="e6778-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="e6778-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-188">Method</span></span> |
| [<span data-ttu-id="e6778-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e6778-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e6778-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-190">Method</span></span> |
| [<span data-ttu-id="e6778-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e6778-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e6778-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-192">Method</span></span> |
| [<span data-ttu-id="e6778-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e6778-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-194">Method</span></span> |
| [<span data-ttu-id="e6778-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="e6778-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="e6778-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-196">Method</span></span> |
| [<span data-ttu-id="e6778-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e6778-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="e6778-198">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-198">Method</span></span> |
| [<span data-ttu-id="e6778-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e6778-200">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-200">Method</span></span> |
| [<span data-ttu-id="e6778-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e6778-202">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-202">Method</span></span> |
| [<span data-ttu-id="e6778-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e6778-204">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-204">Method</span></span> |
| [<span data-ttu-id="e6778-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e6778-206">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-206">Method</span></span> |
| [<span data-ttu-id="e6778-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e6778-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e6778-208">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e6778-209">例</span><span class="sxs-lookup"><span data-stu-id="e6778-209">Example</span></span>

<span data-ttu-id="e6778-210">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e6778-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e6778-211">メンバー</span><span class="sxs-lookup"><span data-stu-id="e6778-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e6778-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e6778-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e6778-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-215">ファイルの特定の種類は、潜在的なセキュリティの問題により、Outlook によってブロックされは返されません。</span><span class="sxs-lookup"><span data-stu-id="e6778-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e6778-216">詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e6778-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-217">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-217">Type:</span></span>

*   <span data-ttu-id="e6778-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e6778-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-219">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-219">Requirements</span></span>

|<span data-ttu-id="e6778-220">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-220">Requirement</span></span>|<span data-ttu-id="e6778-221">値</span><span class="sxs-lookup"><span data-stu-id="e6778-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-223">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-223">1.0</span></span>|
|[<span data-ttu-id="e6778-224">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-225">ReadItem</span></span>|
|[<span data-ttu-id="e6778-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-227">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-228">例</span><span class="sxs-lookup"><span data-stu-id="e6778-228">Example</span></span>

<span data-ttu-id="e6778-229">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="e6778-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e6778-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e6778-231">取得またはメッセージの bcc (ブラインド カーボン コピー) 受信者を更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e6778-232">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-233">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-233">Type:</span></span>

*   [<span data-ttu-id="e6778-234">Recipients</span><span class="sxs-lookup"><span data-stu-id="e6778-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e6778-235">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-235">Requirements</span></span>

|<span data-ttu-id="e6778-236">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-236">Requirement</span></span>|<span data-ttu-id="e6778-237">値</span><span class="sxs-lookup"><span data-stu-id="e6778-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-238">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-239">1.1</span><span class="sxs-lookup"><span data-stu-id="e6778-239">1.1</span></span>|
|[<span data-ttu-id="e6778-240">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-241">ReadItem</span></span>|
|[<span data-ttu-id="e6778-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-243">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-244">例</span><span class="sxs-lookup"><span data-stu-id="e6778-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="e6778-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="e6778-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="e6778-246">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-247">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-247">Type:</span></span>

*   [<span data-ttu-id="e6778-248">Body</span><span class="sxs-lookup"><span data-stu-id="e6778-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="e6778-249">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-249">Requirements</span></span>

|<span data-ttu-id="e6778-250">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-250">Requirement</span></span>|<span data-ttu-id="e6778-251">値</span><span class="sxs-lookup"><span data-stu-id="e6778-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-252">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-253">1.1</span><span class="sxs-lookup"><span data-stu-id="e6778-253">1.1</span></span>|
|[<span data-ttu-id="e6778-254">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-255">ReadItem</span></span>|
|[<span data-ttu-id="e6778-256">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-257">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e6778-258">[cc]: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="e6778-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e6778-259">メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e6778-260">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e6778-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-261">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-261">Read mode</span></span>

<span data-ttu-id="e6778-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-264">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-264">Compose mode</span></span>

<span data-ttu-id="e6778-265">`cc`を`Recipients`オブジェクトを取得または、メッセージの**Cc**行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-266">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-266">Type:</span></span>

*   <span data-ttu-id="e6778-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-268">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-268">Requirements</span></span>

|<span data-ttu-id="e6778-269">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-269">Requirement</span></span>|<span data-ttu-id="e6778-270">値</span><span class="sxs-lookup"><span data-stu-id="e6778-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-272">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-272">1.0</span></span>|
|[<span data-ttu-id="e6778-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-274">ReadItem</span></span>|
|[<span data-ttu-id="e6778-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-276">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-277">例</span><span class="sxs-lookup"><span data-stu-id="e6778-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="e6778-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="e6778-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="e6778-279">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e6778-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="e6778-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e6778-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-284">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-284">Type:</span></span>

*   <span data-ttu-id="e6778-285">String</span><span class="sxs-lookup"><span data-stu-id="e6778-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-286">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-286">Requirements</span></span>

|<span data-ttu-id="e6778-287">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-287">Requirement</span></span>|<span data-ttu-id="e6778-288">値</span><span class="sxs-lookup"><span data-stu-id="e6778-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-289">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-290">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-290">1.0</span></span>|
|[<span data-ttu-id="e6778-291">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-292">ReadItem</span></span>|
|[<span data-ttu-id="e6778-293">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-294">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="e6778-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="e6778-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="e6778-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-298">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-298">Type:</span></span>

*   <span data-ttu-id="e6778-299">日付</span><span class="sxs-lookup"><span data-stu-id="e6778-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-300">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-300">Requirements</span></span>

|<span data-ttu-id="e6778-301">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-301">Requirement</span></span>|<span data-ttu-id="e6778-302">値</span><span class="sxs-lookup"><span data-stu-id="e6778-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-304">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-304">1.0</span></span>|
|[<span data-ttu-id="e6778-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-306">ReadItem</span></span>|
|[<span data-ttu-id="e6778-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-309">例</span><span class="sxs-lookup"><span data-stu-id="e6778-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="e6778-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="e6778-310">dateTimeModified :Date</span></span>

<span data-ttu-id="e6778-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-313">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-314">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-314">Type:</span></span>

*   <span data-ttu-id="e6778-315">日付</span><span class="sxs-lookup"><span data-stu-id="e6778-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-316">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-316">Requirements</span></span>

|<span data-ttu-id="e6778-317">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-317">Requirement</span></span>|<span data-ttu-id="e6778-318">値</span><span class="sxs-lookup"><span data-stu-id="e6778-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-319">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-320">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-320">1.0</span></span>|
|[<span data-ttu-id="e6778-321">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-322">ReadItem</span></span>|
|[<span data-ttu-id="e6778-323">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-324">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-325">例</span><span class="sxs-lookup"><span data-stu-id="e6778-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e6778-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e6778-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e6778-327">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e6778-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-330">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-330">Read mode</span></span>

<span data-ttu-id="e6778-331">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-332">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-332">Compose mode</span></span>

<span data-ttu-id="e6778-333">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e6778-334">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e6778-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-335">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-335">Type:</span></span>

*   <span data-ttu-id="e6778-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e6778-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-337">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-337">Requirements</span></span>

|<span data-ttu-id="e6778-338">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-338">Requirement</span></span>|<span data-ttu-id="e6778-339">値</span><span class="sxs-lookup"><span data-stu-id="e6778-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-340">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-341">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-341">1.0</span></span>|
|[<span data-ttu-id="e6778-342">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-343">ReadItem</span></span>|
|[<span data-ttu-id="e6778-344">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-345">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-346">例</span><span class="sxs-lookup"><span data-stu-id="e6778-346">Example</span></span>

<span data-ttu-id="e6778-347">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="e6778-348">:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e6778-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="e6778-349">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="e6778-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-352">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`from`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e6778-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-353">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-353">Read mode</span></span>

<span data-ttu-id="e6778-354">`from`を`EmailAddressDetails`オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e6778-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="e6778-355">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-355">Compose mode</span></span>

<span data-ttu-id="e6778-356">`from`を`From`を取得するメソッドを提供するオブジェクト、値からです。</span><span class="sxs-lookup"><span data-stu-id="e6778-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e6778-357">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-357">Type:</span></span>

*   <span data-ttu-id="e6778-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e6778-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-359">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-359">Requirements</span></span>

|<span data-ttu-id="e6778-360">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e6778-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-362">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-362">1.0</span></span>|<span data-ttu-id="e6778-363">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-363">Preview</span></span>|
|[<span data-ttu-id="e6778-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-365">ReadItem</span></span>|<span data-ttu-id="e6778-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-368">Read</span><span class="sxs-lookup"><span data-stu-id="e6778-368">Read</span></span>|<span data-ttu-id="e6778-369">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="e6778-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="e6778-370">internetMessageId :String</span></span>

<span data-ttu-id="e6778-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-373">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-373">Type:</span></span>

*   <span data-ttu-id="e6778-374">String</span><span class="sxs-lookup"><span data-stu-id="e6778-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-375">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-375">Requirements</span></span>

|<span data-ttu-id="e6778-376">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-376">Requirement</span></span>|<span data-ttu-id="e6778-377">値</span><span class="sxs-lookup"><span data-stu-id="e6778-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-378">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-379">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-379">1.0</span></span>|
|[<span data-ttu-id="e6778-380">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-381">ReadItem</span></span>|
|[<span data-ttu-id="e6778-382">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-383">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-384">例</span><span class="sxs-lookup"><span data-stu-id="e6778-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="e6778-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="e6778-385">itemClass :String</span></span>

<span data-ttu-id="e6778-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e6778-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="e6778-390">種類</span><span class="sxs-lookup"><span data-stu-id="e6778-390">Type</span></span>|<span data-ttu-id="e6778-391">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-391">Description</span></span>|<span data-ttu-id="e6778-392">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="e6778-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="e6778-393">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="e6778-393">Appointment items</span></span>|<span data-ttu-id="e6778-394">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e6778-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="e6778-395">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="e6778-395">Message items</span></span>|<span data-ttu-id="e6778-396">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="e6778-397">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-398">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-398">Type:</span></span>

*   <span data-ttu-id="e6778-399">String</span><span class="sxs-lookup"><span data-stu-id="e6778-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-400">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-400">Requirements</span></span>

|<span data-ttu-id="e6778-401">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-401">Requirement</span></span>|<span data-ttu-id="e6778-402">値</span><span class="sxs-lookup"><span data-stu-id="e6778-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-403">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-404">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-404">1.0</span></span>|
|[<span data-ttu-id="e6778-405">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-406">ReadItem</span></span>|
|[<span data-ttu-id="e6778-407">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-408">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-409">例</span><span class="sxs-lookup"><span data-stu-id="e6778-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e6778-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="e6778-410">(nullable) itemId :String</span></span>

<span data-ttu-id="e6778-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-413">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="e6778-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e6778-414">`itemId`プロパティは、Outlook のエントリ ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="e6778-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e6778-415">この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e6778-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e6778-416">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e6778-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e6778-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-419">種類:</span><span class="sxs-lookup"><span data-stu-id="e6778-419">Type:</span></span>

*   <span data-ttu-id="e6778-420">String</span><span class="sxs-lookup"><span data-stu-id="e6778-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-421">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-421">Requirements</span></span>

|<span data-ttu-id="e6778-422">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-422">Requirement</span></span>|<span data-ttu-id="e6778-423">値</span><span class="sxs-lookup"><span data-stu-id="e6778-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-425">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-425">1.0</span></span>|
|[<span data-ttu-id="e6778-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-427">ReadItem</span></span>|
|[<span data-ttu-id="e6778-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-429">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-430">例</span><span class="sxs-lookup"><span data-stu-id="e6778-430">Example</span></span>

<span data-ttu-id="e6778-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="e6778-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e6778-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e6778-434">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e6778-435">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="e6778-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-436">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-436">Type:</span></span>

*   [<span data-ttu-id="e6778-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e6778-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e6778-438">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-438">Requirements</span></span>

|<span data-ttu-id="e6778-439">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-439">Requirement</span></span>|<span data-ttu-id="e6778-440">値</span><span class="sxs-lookup"><span data-stu-id="e6778-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-441">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-442">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-442">1.0</span></span>|
|[<span data-ttu-id="e6778-443">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-444">ReadItem</span></span>|
|[<span data-ttu-id="e6778-445">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-446">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-447">例</span><span class="sxs-lookup"><span data-stu-id="e6778-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="e6778-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e6778-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="e6778-449">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-450">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-450">Read mode</span></span>

<span data-ttu-id="e6778-451">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-452">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-452">Compose mode</span></span>

<span data-ttu-id="e6778-453">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-454">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-454">Type:</span></span>

*   <span data-ttu-id="e6778-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e6778-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-456">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-456">Requirements</span></span>

|<span data-ttu-id="e6778-457">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-457">Requirement</span></span>|<span data-ttu-id="e6778-458">値</span><span class="sxs-lookup"><span data-stu-id="e6778-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-459">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-460">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-460">1.0</span></span>|
|[<span data-ttu-id="e6778-461">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-462">ReadItem</span></span>|
|[<span data-ttu-id="e6778-463">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-464">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-465">例</span><span class="sxs-lookup"><span data-stu-id="e6778-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e6778-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="e6778-466">normalizedSubject :String</span></span>

<span data-ttu-id="e6778-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e6778-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-471">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-471">Type:</span></span>

*   <span data-ttu-id="e6778-472">String</span><span class="sxs-lookup"><span data-stu-id="e6778-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-473">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-473">Requirements</span></span>

|<span data-ttu-id="e6778-474">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-474">Requirement</span></span>|<span data-ttu-id="e6778-475">値</span><span class="sxs-lookup"><span data-stu-id="e6778-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-477">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-477">1.0</span></span>|
|[<span data-ttu-id="e6778-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-479">ReadItem</span></span>|
|[<span data-ttu-id="e6778-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-481">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-482">例</span><span class="sxs-lookup"><span data-stu-id="e6778-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="e6778-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="e6778-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="e6778-484">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-485">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-485">Type:</span></span>

*   [<span data-ttu-id="e6778-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e6778-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="e6778-487">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-487">Requirements</span></span>

|<span data-ttu-id="e6778-488">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-488">Requirement</span></span>|<span data-ttu-id="e6778-489">値</span><span class="sxs-lookup"><span data-stu-id="e6778-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-490">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-491">1.3</span><span class="sxs-lookup"><span data-stu-id="e6778-491">1.3</span></span>|
|[<span data-ttu-id="e6778-492">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-493">ReadItem</span></span>|
|[<span data-ttu-id="e6778-494">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-495">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e6778-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e6778-497">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e6778-498">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e6778-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-499">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-499">Read mode</span></span>

<span data-ttu-id="e6778-500">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-501">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-501">Compose mode</span></span>

<span data-ttu-id="e6778-502">`optionalAttendees`を`Recipients`オブジェクトを取得または省略可能な会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-503">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-503">Type:</span></span>

*   <span data-ttu-id="e6778-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-505">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-505">Requirements</span></span>

|<span data-ttu-id="e6778-506">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-506">Requirement</span></span>|<span data-ttu-id="e6778-507">値</span><span class="sxs-lookup"><span data-stu-id="e6778-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-509">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-509">1.0</span></span>|
|[<span data-ttu-id="e6778-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-511">ReadItem</span></span>|
|[<span data-ttu-id="e6778-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-513">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-514">例</span><span class="sxs-lookup"><span data-stu-id="e6778-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="e6778-515">オーガナイザー:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[オーガナイザー](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="e6778-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="e6778-516">指定した会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-517">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-517">Read mode</span></span>

<span data-ttu-id="e6778-518">`organizer`プロパティは、会議の開催者を表す[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-519">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-519">Compose mode</span></span>

<span data-ttu-id="e6778-520">`organizer`プロパティが開催者の値を取得するメソッドを提供する[構成内容変更](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-521">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-521">Type:</span></span>

*   <span data-ttu-id="e6778-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [オーガナイザー](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="e6778-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-523">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-523">Requirements</span></span>

|<span data-ttu-id="e6778-524">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e6778-525">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-526">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-526">1.0</span></span>|<span data-ttu-id="e6778-527">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-527">Preview</span></span>|
|[<span data-ttu-id="e6778-528">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-529">ReadItem</span></span>|<span data-ttu-id="e6778-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-531">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-532">Read</span><span class="sxs-lookup"><span data-stu-id="e6778-532">Read</span></span>|<span data-ttu-id="e6778-533">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-534">例</span><span class="sxs-lookup"><span data-stu-id="e6778-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="e6778-535">(許容) 定期的:[定期的なアイテム](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="e6778-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="e6778-536">取得または予定の定期的なパターンを設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="e6778-537">定期的な会議出席依頼を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="e6778-538">モードの予定表アイテムを読んだり作成したりします。</span><span class="sxs-lookup"><span data-stu-id="e6778-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="e6778-539">会議出席依頼アイテムの読み取りモードです。</span><span class="sxs-lookup"><span data-stu-id="e6778-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="e6778-540">`recurrence`プロパティは、アイテムが系列または系列のインスタンスである場合に定期的な予定または会議出席依頼に[定期的なアイテム](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="e6778-541">`null`単独の予定および会議出席依頼を単独の予定が返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="e6778-542">`undefined`会議出席依頼ではないメッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="e6778-543">注: 会議出席依頼がある、 `itemClass` IPM の値です。Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="e6778-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="e6778-544">注: 定期的なアイテム オブジェクトがある場合`null`、これは、オブジェクトが 1 つの予定または会議出席依頼、単独の予定および一連の一部ではないのであることを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-545">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-545">Type:</span></span>

* [<span data-ttu-id="e6778-546">定期的なアイテム</span><span class="sxs-lookup"><span data-stu-id="e6778-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="e6778-547">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-547">Requirement</span></span>|<span data-ttu-id="e6778-548">値</span><span class="sxs-lookup"><span data-stu-id="e6778-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-550">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-550">Preview</span></span>|
|[<span data-ttu-id="e6778-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-552">ReadItem</span></span>|
|[<span data-ttu-id="e6778-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-554">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e6778-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e6778-556">イベントの出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e6778-557">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e6778-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-558">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-558">Read mode</span></span>

<span data-ttu-id="e6778-559">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-560">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-560">Compose mode</span></span>

<span data-ttu-id="e6778-561">`requiredAttendees`を`Recipients`オブジェクトを取得または会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-562">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-562">Type:</span></span>

*   <span data-ttu-id="e6778-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-564">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-564">Requirements</span></span>

|<span data-ttu-id="e6778-565">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-565">Requirement</span></span>|<span data-ttu-id="e6778-566">値</span><span class="sxs-lookup"><span data-stu-id="e6778-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-568">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-568">1.0</span></span>|
|[<span data-ttu-id="e6778-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-570">ReadItem</span></span>|
|[<span data-ttu-id="e6778-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-572">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-573">例</span><span class="sxs-lookup"><span data-stu-id="e6778-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="e6778-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e6778-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="e6778-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e6778-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e6778-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-579">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`sender`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e6778-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-580">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-580">Type:</span></span>

*   [<span data-ttu-id="e6778-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e6778-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e6778-582">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-582">Requirements</span></span>

|<span data-ttu-id="e6778-583">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-583">Requirement</span></span>|<span data-ttu-id="e6778-584">値</span><span class="sxs-lookup"><span data-stu-id="e6778-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-585">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-586">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-586">1.0</span></span>|
|[<span data-ttu-id="e6778-587">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-588">ReadItem</span></span>|
|[<span data-ttu-id="e6778-589">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-590">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-591">例</span><span class="sxs-lookup"><span data-stu-id="e6778-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="e6778-592">(許容) seriesId: 文字列</span><span class="sxs-lookup"><span data-stu-id="e6778-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="e6778-593">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="e6778-594">OWA と outlook 2002 で、`seriesId`は、この項目が属する親 (系列) アイテムの Exchange Web サービス (EWS) の ID を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="e6778-595">IOS および Android で、 `seriesId` 、親項目の残りの部分 ID を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-596">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="e6778-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e6778-597">`seriesId`プロパティは Outlook の REST API で使用される Outlook の Id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="e6778-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="e6778-598">この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e6778-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e6778-599">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e6778-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="e6778-600">`seriesId`プロパティを返します。`null`アイテムの親アイテムを次のようにされていない単一の関連するアイテム、予定または会議を要求し、返しますの`undefined`、その他の項目の要求を満たしていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-601">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-601">Type:</span></span>

* <span data-ttu-id="e6778-602">String</span><span class="sxs-lookup"><span data-stu-id="e6778-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-603">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-603">Requirements</span></span>

|<span data-ttu-id="e6778-604">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-604">Requirement</span></span>|<span data-ttu-id="e6778-605">値</span><span class="sxs-lookup"><span data-stu-id="e6778-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-607">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-607">Preview</span></span>|
|[<span data-ttu-id="e6778-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-609">ReadItem</span></span>|
|[<span data-ttu-id="e6778-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-611">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-612">例</span><span class="sxs-lookup"><span data-stu-id="e6778-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e6778-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e6778-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e6778-614">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e6778-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-617">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-617">Read mode</span></span>

<span data-ttu-id="e6778-618">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-619">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-619">Compose mode</span></span>

<span data-ttu-id="e6778-620">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e6778-621">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e6778-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-622">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-622">Type:</span></span>

*   <span data-ttu-id="e6778-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e6778-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-624">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-624">Requirements</span></span>

|<span data-ttu-id="e6778-625">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-625">Requirement</span></span>|<span data-ttu-id="e6778-626">値</span><span class="sxs-lookup"><span data-stu-id="e6778-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-627">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-628">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-628">1.0</span></span>|
|[<span data-ttu-id="e6778-629">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-630">ReadItem</span></span>|
|[<span data-ttu-id="e6778-631">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-632">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-633">例</span><span class="sxs-lookup"><span data-stu-id="e6778-633">Example</span></span>

<span data-ttu-id="e6778-634">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="e6778-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e6778-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="e6778-636">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e6778-637">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e6778-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-638">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-638">Read mode</span></span>

<span data-ttu-id="e6778-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="e6778-641">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-641">Compose mode</span></span>

<span data-ttu-id="e6778-642">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e6778-643">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-643">Type:</span></span>

*   <span data-ttu-id="e6778-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e6778-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-645">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-645">Requirements</span></span>

|<span data-ttu-id="e6778-646">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-646">Requirement</span></span>|<span data-ttu-id="e6778-647">値</span><span class="sxs-lookup"><span data-stu-id="e6778-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-648">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-649">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-649">1.0</span></span>|
|[<span data-ttu-id="e6778-650">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-651">ReadItem</span></span>|
|[<span data-ttu-id="e6778-652">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-653">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e6778-654">: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="e6778-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e6778-655">[メッセージの [**宛先**] 行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e6778-656">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e6778-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e6778-657">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e6778-657">Read mode</span></span>

<span data-ttu-id="e6778-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e6778-660">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e6778-660">Compose mode</span></span>

<span data-ttu-id="e6778-661">`to`を`Recipients`オブジェクトを取得または、メッセージの [**宛先**] 行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e6778-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e6778-662">型:</span><span class="sxs-lookup"><span data-stu-id="e6778-662">Type:</span></span>

*   <span data-ttu-id="e6778-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e6778-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-664">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-664">Requirements</span></span>

|<span data-ttu-id="e6778-665">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-665">Requirement</span></span>|<span data-ttu-id="e6778-666">値</span><span class="sxs-lookup"><span data-stu-id="e6778-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-667">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-668">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-668">1.0</span></span>|
|[<span data-ttu-id="e6778-669">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-670">ReadItem</span></span>|
|[<span data-ttu-id="e6778-671">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-672">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-673">例</span><span class="sxs-lookup"><span data-stu-id="e6778-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="e6778-674">メソッド</span><span class="sxs-lookup"><span data-stu-id="e6778-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e6778-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e6778-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e6778-676">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="e6778-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e6778-677">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="e6778-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e6778-678">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-679">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-679">Parameters:</span></span>
|<span data-ttu-id="e6778-680">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-680">Name</span></span>|<span data-ttu-id="e6778-681">型</span><span class="sxs-lookup"><span data-stu-id="e6778-681">Type</span></span>|<span data-ttu-id="e6778-682">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-682">Attributes</span></span>|<span data-ttu-id="e6778-683">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="e6778-684">String</span><span class="sxs-lookup"><span data-stu-id="e6778-684">String</span></span>||<span data-ttu-id="e6778-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e6778-687">String</span><span class="sxs-lookup"><span data-stu-id="e6778-687">String</span></span>||<span data-ttu-id="e6778-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e6778-690">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-690">Object</span></span>|<span data-ttu-id="e6778-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-691">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-692">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-693">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-693">Object</span></span>|<span data-ttu-id="e6778-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-694">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-695">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e6778-696">Boolean</span><span class="sxs-lookup"><span data-stu-id="e6778-696">Boolean</span></span>|<span data-ttu-id="e6778-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-697">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-698">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e6778-699">function</span><span class="sxs-lookup"><span data-stu-id="e6778-699">function</span></span>|<span data-ttu-id="e6778-700">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-700">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-701">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e6778-702">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e6778-703">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e6778-704">エラー</span><span class="sxs-lookup"><span data-stu-id="e6778-704">Errors</span></span>

|<span data-ttu-id="e6778-705">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e6778-705">Error code</span></span>|<span data-ttu-id="e6778-706">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e6778-707">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="e6778-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e6778-708">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="e6778-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e6778-709">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="e6778-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-710">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-710">Requirements</span></span>

|<span data-ttu-id="e6778-711">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-711">Requirement</span></span>|<span data-ttu-id="e6778-712">値</span><span class="sxs-lookup"><span data-stu-id="e6778-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-713">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-714">1.1</span><span class="sxs-lookup"><span data-stu-id="e6778-714">1.1</span></span>|
|[<span data-ttu-id="e6778-715">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-717">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-718">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e6778-719">例</span><span class="sxs-lookup"><span data-stu-id="e6778-719">Examples</span></span>

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

<span data-ttu-id="e6778-720">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="e6778-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="e6778-721">addFileAttachmentFromBase64Async (base64File、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="e6778-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e6778-722">メッセージまたは予定を添付ファイルとしてエンコード base64 からファイルを追加します。</span><span class="sxs-lookup"><span data-stu-id="e6778-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e6778-723">`addFileAttachmentFromBase64Async`メソッドは、base64 エンコーディングからファイルをアップロードし、作成フォーム内の項目にアタッチします。</span><span class="sxs-lookup"><span data-stu-id="e6778-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="e6778-724">このメソッドは、AsyncResult.value オブジェクトの添付ファイルの識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="e6778-725">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-726">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-726">Parameters:</span></span>
|<span data-ttu-id="e6778-727">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-727">Name</span></span>|<span data-ttu-id="e6778-728">型</span><span class="sxs-lookup"><span data-stu-id="e6778-728">Type</span></span>|<span data-ttu-id="e6778-729">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-729">Attributes</span></span>|<span data-ttu-id="e6778-730">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="e6778-731">String</span><span class="sxs-lookup"><span data-stu-id="e6778-731">String</span></span>||<span data-ttu-id="e6778-732">イメージや、電子メール、またはイベントに追加するファイルのコンテンツを base64 にエンコードされます。</span><span class="sxs-lookup"><span data-stu-id="e6778-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="e6778-733">String</span><span class="sxs-lookup"><span data-stu-id="e6778-733">String</span></span>||<span data-ttu-id="e6778-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e6778-736">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-736">Object</span></span>|<span data-ttu-id="e6778-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-737">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-738">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-739">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-739">Object</span></span>|<span data-ttu-id="e6778-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-740">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-741">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e6778-742">Boolean</span><span class="sxs-lookup"><span data-stu-id="e6778-742">Boolean</span></span>|<span data-ttu-id="e6778-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-743">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-744">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e6778-745">function</span><span class="sxs-lookup"><span data-stu-id="e6778-745">function</span></span>|<span data-ttu-id="e6778-746">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-746">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-747">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e6778-748">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e6778-749">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e6778-750">エラー</span><span class="sxs-lookup"><span data-stu-id="e6778-750">Errors</span></span>

|<span data-ttu-id="e6778-751">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e6778-751">Error code</span></span>|<span data-ttu-id="e6778-752">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e6778-753">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="e6778-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e6778-754">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="e6778-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e6778-755">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="e6778-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-756">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-756">Requirements</span></span>

|<span data-ttu-id="e6778-757">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-757">Requirement</span></span>|<span data-ttu-id="e6778-758">値</span><span class="sxs-lookup"><span data-stu-id="e6778-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-759">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-760">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-760">Preview</span></span>|
|[<span data-ttu-id="e6778-761">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-763">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-764">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e6778-765">例</span><span class="sxs-lookup"><span data-stu-id="e6778-765">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e6778-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e6778-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e6778-767">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="e6778-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e6778-768">現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="e6778-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-769">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-769">Parameters:</span></span>

| <span data-ttu-id="e6778-770">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-770">Name</span></span> | <span data-ttu-id="e6778-771">型</span><span class="sxs-lookup"><span data-stu-id="e6778-771">Type</span></span> | <span data-ttu-id="e6778-772">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-772">Attributes</span></span> | <span data-ttu-id="e6778-773">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e6778-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e6778-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e6778-775">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="e6778-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e6778-776">Function</span><span class="sxs-lookup"><span data-stu-id="e6778-776">Function</span></span> || <span data-ttu-id="e6778-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e6778-780">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-780">Object</span></span> | <span data-ttu-id="e6778-781">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-781">&lt;optional&gt;</span></span> | <span data-ttu-id="e6778-782">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e6778-783">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-783">Object</span></span> | <span data-ttu-id="e6778-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-784">&lt;optional&gt;</span></span> | <span data-ttu-id="e6778-785">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e6778-786">function</span><span class="sxs-lookup"><span data-stu-id="e6778-786">function</span></span>| <span data-ttu-id="e6778-787">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-787">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-788">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-789">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-789">Requirements</span></span>

|<span data-ttu-id="e6778-790">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-790">Requirement</span></span>| <span data-ttu-id="e6778-791">値</span><span class="sxs-lookup"><span data-stu-id="e6778-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-792">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6778-793">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-793">Preview</span></span> |
|[<span data-ttu-id="e6778-794">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6778-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-795">ReadItem</span></span> |
|[<span data-ttu-id="e6778-796">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6778-797">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="e6778-798">例</span><span class="sxs-lookup"><span data-stu-id="e6778-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e6778-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e6778-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e6778-800">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="e6778-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e6778-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e6778-804">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e6778-805">Office アドインは、Outlook Web App で実行されている場合、`addItemAttachmentAsync`メソッドが項目を編集しているアイテム以外のアイテムに関連付けることができますただし、これはサポートされていません、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="e6778-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-806">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-806">Parameters:</span></span>

|<span data-ttu-id="e6778-807">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-807">Name</span></span>|<span data-ttu-id="e6778-808">型</span><span class="sxs-lookup"><span data-stu-id="e6778-808">Type</span></span>|<span data-ttu-id="e6778-809">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-809">Attributes</span></span>|<span data-ttu-id="e6778-810">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="e6778-811">String</span><span class="sxs-lookup"><span data-stu-id="e6778-811">String</span></span>||<span data-ttu-id="e6778-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e6778-814">String</span><span class="sxs-lookup"><span data-stu-id="e6778-814">String</span></span>||<span data-ttu-id="e6778-p141">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e6778-817">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-817">Object</span></span>|<span data-ttu-id="e6778-818">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-818">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-819">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-820">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-820">Object</span></span>|<span data-ttu-id="e6778-821">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-821">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-822">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e6778-823">function</span><span class="sxs-lookup"><span data-stu-id="e6778-823">function</span></span>|<span data-ttu-id="e6778-824">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-824">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-825">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e6778-826">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e6778-827">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e6778-828">エラー</span><span class="sxs-lookup"><span data-stu-id="e6778-828">Errors</span></span>

|<span data-ttu-id="e6778-829">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e6778-829">Error code</span></span>|<span data-ttu-id="e6778-830">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e6778-831">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="e6778-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-832">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-832">Requirements</span></span>

|<span data-ttu-id="e6778-833">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-833">Requirement</span></span>|<span data-ttu-id="e6778-834">値</span><span class="sxs-lookup"><span data-stu-id="e6778-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-835">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-836">1.1</span><span class="sxs-lookup"><span data-stu-id="e6778-836">1.1</span></span>|
|[<span data-ttu-id="e6778-837">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-839">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-840">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-841">例</span><span class="sxs-lookup"><span data-stu-id="e6778-841">Example</span></span>

<span data-ttu-id="e6778-842">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="e6778-843">close()</span><span class="sxs-lookup"><span data-stu-id="e6778-843">close()</span></span>

<span data-ttu-id="e6778-844">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="e6778-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e6778-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-847">アイテム予定は、以前保存されたを使用する場合は、web 上の Outlook で`saveAsync`を求めるメッセージを保存、破棄、または、キャンセル場合でも、変更が発生していないから、項目を保存します。</span><span class="sxs-lookup"><span data-stu-id="e6778-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e6778-848">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="e6778-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-849">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-849">Requirements</span></span>

|<span data-ttu-id="e6778-850">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-850">Requirement</span></span>|<span data-ttu-id="e6778-851">値</span><span class="sxs-lookup"><span data-stu-id="e6778-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-852">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-853">1.3</span><span class="sxs-lookup"><span data-stu-id="e6778-853">1.3</span></span>|
|[<span data-ttu-id="e6778-854">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-855">制限あり</span><span class="sxs-lookup"><span data-stu-id="e6778-855">Restricted</span></span>|
|[<span data-ttu-id="e6778-856">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-857">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="e6778-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e6778-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="e6778-859">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-860">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-861">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e6778-862">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="e6778-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e6778-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="e6778-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-866">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-866">Parameters:</span></span>

|<span data-ttu-id="e6778-867">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-867">Name</span></span>|<span data-ttu-id="e6778-868">型</span><span class="sxs-lookup"><span data-stu-id="e6778-868">Type</span></span>|<span data-ttu-id="e6778-869">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-869">Attributes</span></span>|<span data-ttu-id="e6778-870">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e6778-871">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e6778-871">String &#124; Object</span></span>||<span data-ttu-id="e6778-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e6778-874">**または**</span><span class="sxs-lookup"><span data-stu-id="e6778-874">**OR**</span></span><br/><span data-ttu-id="e6778-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e6778-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e6778-877">String</span><span class="sxs-lookup"><span data-stu-id="e6778-877">String</span></span>|<span data-ttu-id="e6778-878">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-878">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e6778-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e6778-882">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-882">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-883">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="e6778-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e6778-884">String</span><span class="sxs-lookup"><span data-stu-id="e6778-884">String</span></span>||<span data-ttu-id="e6778-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e6778-887">String</span><span class="sxs-lookup"><span data-stu-id="e6778-887">String</span></span>||<span data-ttu-id="e6778-888">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e6778-889">String</span><span class="sxs-lookup"><span data-stu-id="e6778-889">String</span></span>||<span data-ttu-id="e6778-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="e6778-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e6778-892">Boolean</span><span class="sxs-lookup"><span data-stu-id="e6778-892">Boolean</span></span>||<span data-ttu-id="e6778-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e6778-895">String</span><span class="sxs-lookup"><span data-stu-id="e6778-895">String</span></span>||<span data-ttu-id="e6778-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e6778-899">function</span><span class="sxs-lookup"><span data-stu-id="e6778-899">function</span></span>|<span data-ttu-id="e6778-900">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-900">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-901">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-902">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-902">Requirements</span></span>

|<span data-ttu-id="e6778-903">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-903">Requirement</span></span>|<span data-ttu-id="e6778-904">値</span><span class="sxs-lookup"><span data-stu-id="e6778-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-906">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-906">1.0</span></span>|
|[<span data-ttu-id="e6778-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-908">ReadItem</span></span>|
|[<span data-ttu-id="e6778-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-910">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e6778-911">例</span><span class="sxs-lookup"><span data-stu-id="e6778-911">Examples</span></span>

<span data-ttu-id="e6778-912">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="e6778-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e6778-913">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e6778-914">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e6778-915">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-915">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e6778-916">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-916">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e6778-917">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="e6778-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e6778-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="e6778-919">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-920">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-921">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e6778-922">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="e6778-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e6778-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="e6778-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-926">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-926">Parameters:</span></span>

|<span data-ttu-id="e6778-927">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-927">Name</span></span>|<span data-ttu-id="e6778-928">型</span><span class="sxs-lookup"><span data-stu-id="e6778-928">Type</span></span>|<span data-ttu-id="e6778-929">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-929">Attributes</span></span>|<span data-ttu-id="e6778-930">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e6778-931">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e6778-931">String &#124; Object</span></span>||<span data-ttu-id="e6778-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e6778-934">**または**</span><span class="sxs-lookup"><span data-stu-id="e6778-934">**OR**</span></span><br/><span data-ttu-id="e6778-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e6778-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e6778-937">String</span><span class="sxs-lookup"><span data-stu-id="e6778-937">String</span></span>|<span data-ttu-id="e6778-938">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-938">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e6778-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e6778-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e6778-942">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-942">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-943">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="e6778-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e6778-944">String</span><span class="sxs-lookup"><span data-stu-id="e6778-944">String</span></span>||<span data-ttu-id="e6778-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e6778-947">String</span><span class="sxs-lookup"><span data-stu-id="e6778-947">String</span></span>||<span data-ttu-id="e6778-948">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e6778-949">String</span><span class="sxs-lookup"><span data-stu-id="e6778-949">String</span></span>||<span data-ttu-id="e6778-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="e6778-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e6778-952">Boolean</span><span class="sxs-lookup"><span data-stu-id="e6778-952">Boolean</span></span>||<span data-ttu-id="e6778-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e6778-955">String</span><span class="sxs-lookup"><span data-stu-id="e6778-955">String</span></span>||<span data-ttu-id="e6778-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e6778-959">function</span><span class="sxs-lookup"><span data-stu-id="e6778-959">function</span></span>|<span data-ttu-id="e6778-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-960">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-961">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-962">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-962">Requirements</span></span>

|<span data-ttu-id="e6778-963">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-963">Requirement</span></span>|<span data-ttu-id="e6778-964">値</span><span class="sxs-lookup"><span data-stu-id="e6778-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-965">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-966">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-966">1.0</span></span>|
|[<span data-ttu-id="e6778-967">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-968">ReadItem</span></span>|
|[<span data-ttu-id="e6778-969">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-970">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e6778-971">例</span><span class="sxs-lookup"><span data-stu-id="e6778-971">Examples</span></span>

<span data-ttu-id="e6778-972">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="e6778-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e6778-973">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e6778-974">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e6778-975">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-975">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e6778-976">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-976">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e6778-977">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="e6778-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e6778-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e6778-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e6778-979">選択したアイテムの本文内のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-980">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-981">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-981">Requirements</span></span>

|<span data-ttu-id="e6778-982">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-982">Requirement</span></span>|<span data-ttu-id="e6778-983">値</span><span class="sxs-lookup"><span data-stu-id="e6778-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-984">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-985">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-985">1.0</span></span>|
|[<span data-ttu-id="e6778-986">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-987">ReadItem</span></span>|
|[<span data-ttu-id="e6778-988">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-989">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-990">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-990">Returns:</span></span>

<span data-ttu-id="e6778-991">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e6778-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e6778-992">例</span><span class="sxs-lookup"><span data-stu-id="e6778-992">Example</span></span>

<span data-ttu-id="e6778-993">次の使用例は、現在の項目の本文に連絡先のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e6778-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e6778-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e6778-995">選択したアイテムの本文に指定されたエンティティ型のすべてのエンティティの配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-996">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-997">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-997">Parameters:</span></span>

|<span data-ttu-id="e6778-998">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-998">Name</span></span>|<span data-ttu-id="e6778-999">種類</span><span class="sxs-lookup"><span data-stu-id="e6778-999">Type</span></span>|<span data-ttu-id="e6778-1000">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="e6778-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e6778-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="e6778-1002">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="e6778-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1003">Requirements</span><span class="sxs-lookup"><span data-stu-id="e6778-1003">Requirements</span></span>

|<span data-ttu-id="e6778-1004">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1004">Requirement</span></span>|<span data-ttu-id="e6778-1005">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1006">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-1007">1.0</span></span>|
|[<span data-ttu-id="e6778-1008">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1009">制限あり</span><span class="sxs-lookup"><span data-stu-id="e6778-1009">Restricted</span></span>|
|[<span data-ttu-id="e6778-1010">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1011">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1012">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1012">Returns:</span></span>

<span data-ttu-id="e6778-1013">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e6778-1014">アイテムの本文に指定した型のエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e6778-1015">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="e6778-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e6778-1016">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="e6778-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="e6778-1017">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="e6778-1017">Value of `entityType`</span></span>|<span data-ttu-id="e6778-1018">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="e6778-1018">Type of objects in returned array</span></span>|<span data-ttu-id="e6778-1019">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="e6778-1020">文字列</span><span class="sxs-lookup"><span data-stu-id="e6778-1020">String</span></span>|<span data-ttu-id="e6778-1021">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e6778-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="e6778-1022">連絡先</span><span class="sxs-lookup"><span data-stu-id="e6778-1022">Contact</span></span>|<span data-ttu-id="e6778-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e6778-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="e6778-1024">文字列</span><span class="sxs-lookup"><span data-stu-id="e6778-1024">String</span></span>|<span data-ttu-id="e6778-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e6778-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="e6778-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e6778-1026">MeetingSuggestion</span></span>|<span data-ttu-id="e6778-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e6778-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="e6778-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e6778-1028">PhoneNumber</span></span>|<span data-ttu-id="e6778-1029">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e6778-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="e6778-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e6778-1030">TaskSuggestion</span></span>|<span data-ttu-id="e6778-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e6778-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="e6778-1032">文字列</span><span class="sxs-lookup"><span data-stu-id="e6778-1032">String</span></span>|<span data-ttu-id="e6778-1033">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e6778-1033">**Restricted**</span></span>|

<span data-ttu-id="e6778-1034">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e6778-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="e6778-1035">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1035">Example</span></span>

<span data-ttu-id="e6778-1036">次の例では、現在の項目の本文に郵便番号のアドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e6778-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e6778-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e6778-1038">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1039">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-1040">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1041">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1041">Parameters:</span></span>

|<span data-ttu-id="e6778-1042">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1042">Name</span></span>|<span data-ttu-id="e6778-1043">種類</span><span class="sxs-lookup"><span data-stu-id="e6778-1043">Type</span></span>|<span data-ttu-id="e6778-1044">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e6778-1045">String</span><span class="sxs-lookup"><span data-stu-id="e6778-1045">String</span></span>|<span data-ttu-id="e6778-1046">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="e6778-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1047">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1047">Requirements</span></span>

|<span data-ttu-id="e6778-1048">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1048">Requirement</span></span>|<span data-ttu-id="e6778-1049">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1050">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-1051">1.0</span></span>|
|[<span data-ttu-id="e6778-1052">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1053">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1054">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1055">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1056">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1056">Returns:</span></span>

<span data-ttu-id="e6778-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e6778-1059">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e6778-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="e6778-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e6778-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="e6778-1061">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1062">このメソッドのみでサポートされて Outlook 2016 (クイック実行バージョンが 16.0.8413.1000 より大きい値) を Windows および web 上で Outlook を Office 365 の。</span><span class="sxs-lookup"><span data-stu-id="e6778-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1063">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1063">Parameters:</span></span>
|<span data-ttu-id="e6778-1064">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1064">Name</span></span>|<span data-ttu-id="e6778-1065">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1065">Type</span></span>|<span data-ttu-id="e6778-1066">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1066">Attributes</span></span>|<span data-ttu-id="e6778-1067">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e6778-1068">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e6778-1068">Object</span></span>|<span data-ttu-id="e6778-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1070">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-1071">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1071">Object</span></span>|<span data-ttu-id="e6778-1072">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1073">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e6778-1074">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1074">function</span></span>|<span data-ttu-id="e6778-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1076">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e6778-1077">成功した場合、初期化データが提供されている、`asyncResult.value`文字列としてのプロパティです。</span><span class="sxs-lookup"><span data-stu-id="e6778-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="e6778-1078">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1079">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1079">Requirements</span></span>

|<span data-ttu-id="e6778-1080">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1080">Requirement</span></span>|<span data-ttu-id="e6778-1081">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1082">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1083">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-1083">Preview</span></span>|
|[<span data-ttu-id="e6778-1084">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1085">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1086">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1087">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-1088">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1088">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="e6778-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e6778-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e6778-1090">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1091">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e6778-1095">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="e6778-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e6778-1096">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="e6778-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e6778-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-1100">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1100">Requirements</span></span>

|<span data-ttu-id="e6778-1101">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1101">Requirement</span></span>|<span data-ttu-id="e6778-1102">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1103">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-1104">1.0</span></span>|
|[<span data-ttu-id="e6778-1105">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1106">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1107">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1108">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1109">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1109">Returns:</span></span>

<span data-ttu-id="e6778-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="e6778-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e6778-1112">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e6778-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e6778-1113">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e6778-1114">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1114">Example</span></span>

<span data-ttu-id="e6778-1115">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e6778-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e6778-1116">getRegExMatchesByName(name)] → [(許容) {配列。 < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="e6778-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e6778-1117">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1118">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-1119">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="e6778-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e6778-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="e6778-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1122">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1122">Parameters:</span></span>

|<span data-ttu-id="e6778-1123">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1123">Name</span></span>|<span data-ttu-id="e6778-1124">種類</span><span class="sxs-lookup"><span data-stu-id="e6778-1124">Type</span></span>|<span data-ttu-id="e6778-1125">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e6778-1126">String</span><span class="sxs-lookup"><span data-stu-id="e6778-1126">String</span></span>|<span data-ttu-id="e6778-1127">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="e6778-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1128">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1128">Requirements</span></span>

|<span data-ttu-id="e6778-1129">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1129">Requirement</span></span>|<span data-ttu-id="e6778-1130">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-1132">1.0</span></span>|
|[<span data-ttu-id="e6778-1133">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1134">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1135">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1136">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1137">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1137">Returns:</span></span>

<span data-ttu-id="e6778-1138">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="e6778-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="e6778-1139">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e6778-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e6778-1140">配列。 < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="e6778-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e6778-1141">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e6778-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e6778-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e6778-1143">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e6778-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1146">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1146">Parameters:</span></span>

|<span data-ttu-id="e6778-1147">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1147">Name</span></span>|<span data-ttu-id="e6778-1148">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1148">Type</span></span>|<span data-ttu-id="e6778-1149">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1149">Attributes</span></span>|<span data-ttu-id="e6778-1150">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="e6778-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e6778-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e6778-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="e6778-1155">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1155">Object</span></span>|<span data-ttu-id="e6778-1156">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1157">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-1158">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1158">Object</span></span>|<span data-ttu-id="e6778-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1160">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e6778-1161">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1161">function</span></span>||<span data-ttu-id="e6778-1162">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e6778-1163">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e6778-1164">選択範囲は、source プロパティにアクセスするには、呼び出す`asyncResult.value.sourceProperty`、いずれかの方法となる`body`または`subject`。</span><span class="sxs-lookup"><span data-stu-id="e6778-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1165">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1165">Requirements</span></span>

|<span data-ttu-id="e6778-1166">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1166">Requirement</span></span>|<span data-ttu-id="e6778-1167">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="e6778-1169">1.2</span></span>|
|[<span data-ttu-id="e6778-1170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-1172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1173">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1174">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1174">Returns:</span></span>

<span data-ttu-id="e6778-1175">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="e6778-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="e6778-1176">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e6778-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e6778-1177">String</span><span class="sxs-lookup"><span data-stu-id="e6778-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e6778-1178">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1178">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e6778-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e6778-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e6778-p168">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1182">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-1183">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1183">Requirements</span></span>

|<span data-ttu-id="e6778-1184">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1184">Requirement</span></span>|<span data-ttu-id="e6778-1185">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="e6778-1187">1.6</span></span>|
|[<span data-ttu-id="e6778-1188">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1189">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1190">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1191">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1192">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1192">Returns:</span></span>

<span data-ttu-id="e6778-1193">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e6778-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e6778-1194">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1194">Example</span></span>

<span data-ttu-id="e6778-1195">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="e6778-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="e6778-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e6778-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="e6778-p169">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1199">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e6778-p170">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e6778-1203">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="e6778-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e6778-1204">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="e6778-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e6778-p171">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e6778-1208">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1208">Requirements</span></span>

|<span data-ttu-id="e6778-1209">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1209">Requirement</span></span>|<span data-ttu-id="e6778-1210">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="e6778-1212">1.6</span></span>|
|[<span data-ttu-id="e6778-1213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1214">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1216">読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e6778-1217">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e6778-1217">Returns:</span></span>

<span data-ttu-id="e6778-p172">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="e6778-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="e6778-1220">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1220">Example</span></span>

<span data-ttu-id="e6778-1221">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e6778-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e6778-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e6778-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e6778-1223">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e6778-p173">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="e6778-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1227">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1227">Parameters:</span></span>

|<span data-ttu-id="e6778-1228">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1228">Name</span></span>|<span data-ttu-id="e6778-1229">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1229">Type</span></span>|<span data-ttu-id="e6778-1230">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1230">Attributes</span></span>|<span data-ttu-id="e6778-1231">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="e6778-1232">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1232">function</span></span>||<span data-ttu-id="e6778-1233">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e6778-1234">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e6778-1235">取得し、アイテムのカスタム プロパティを削除してサーバーにバックアップを設定するカスタム プロパティに対する変更を保存するのには、このオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="e6778-1236">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e6778-1236">Object</span></span>|<span data-ttu-id="e6778-1237">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1238">開発者は、コールバック関数にアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e6778-1239">によってこのオブジェクトにアクセスできる、`asyncResult.asyncContext`コールバック関数のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="e6778-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1240">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1240">Requirements</span></span>

|<span data-ttu-id="e6778-1241">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1241">Requirement</span></span>|<span data-ttu-id="e6778-1242">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="e6778-1244">1.0</span></span>|
|[<span data-ttu-id="e6778-1245">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1246">ReadItem</span></span>|
|[<span data-ttu-id="e6778-1247">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1248">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-1249">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1249">Example</span></span>

<span data-ttu-id="e6778-p176">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e6778-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e6778-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e6778-1254">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e6778-p177">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1259">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1259">Parameters:</span></span>

|<span data-ttu-id="e6778-1260">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1260">Name</span></span>|<span data-ttu-id="e6778-1261">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1261">Type</span></span>|<span data-ttu-id="e6778-1262">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1262">Attributes</span></span>|<span data-ttu-id="e6778-1263">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e6778-1264">String</span><span class="sxs-lookup"><span data-stu-id="e6778-1264">String</span></span>||<span data-ttu-id="e6778-p178">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="e6778-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="e6778-1267">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1267">Object</span></span>|<span data-ttu-id="e6778-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1269">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-1270">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1270">Object</span></span>|<span data-ttu-id="e6778-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1272">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e6778-1273">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1273">function</span></span>|<span data-ttu-id="e6778-1274">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1275">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e6778-1276">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e6778-1277">エラー</span><span class="sxs-lookup"><span data-stu-id="e6778-1277">Errors</span></span>

|<span data-ttu-id="e6778-1278">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e6778-1278">Error code</span></span>|<span data-ttu-id="e6778-1279">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="e6778-1280">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="e6778-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1281">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1281">Requirements</span></span>

|<span data-ttu-id="e6778-1282">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1282">Requirement</span></span>|<span data-ttu-id="e6778-1283">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="e6778-1285">1.1</span></span>|
|[<span data-ttu-id="e6778-1286">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-1288">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1289">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-1290">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1290">Example</span></span>

<span data-ttu-id="e6778-1291">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1291">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e6778-1292">removeHandlerAsync (イベントの種類、ハンドラー、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="e6778-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e6778-1293">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="e6778-1294">現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="e6778-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1295">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1295">Parameters:</span></span>

| <span data-ttu-id="e6778-1296">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1296">Name</span></span> | <span data-ttu-id="e6778-1297">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1297">Type</span></span> | <span data-ttu-id="e6778-1298">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1298">Attributes</span></span> | <span data-ttu-id="e6778-1299">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e6778-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e6778-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e6778-1301">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="e6778-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e6778-1302">Function</span><span class="sxs-lookup"><span data-stu-id="e6778-1302">Function</span></span> || <span data-ttu-id="e6778-p179">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="e6778-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e6778-1306">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1306">Object</span></span> | <span data-ttu-id="e6778-1307">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="e6778-1308">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e6778-1309">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1309">Object</span></span> | <span data-ttu-id="e6778-1310">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="e6778-1311">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e6778-1312">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1312">function</span></span>| <span data-ttu-id="e6778-1313">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1314">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1315">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1315">Requirements</span></span>

|<span data-ttu-id="e6778-1316">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1316">Requirement</span></span>| <span data-ttu-id="e6778-1317">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1318">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e6778-1319">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e6778-1319">Preview</span></span> |
|[<span data-ttu-id="e6778-1320">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e6778-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1321">ReadItem</span></span> |
|[<span data-ttu-id="e6778-1322">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e6778-1323">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e6778-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="e6778-1324">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="e6778-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e6778-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="e6778-1326">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="e6778-p180">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1330">アドインを呼び出す場合は、`saveAsync`内のアイテムの作成モードを取得するのには、 `itemId` EWS または REST API を使用するにすると、Outlook キャッシュ モードでは、かかる場合がある項目が実際には、サーバーと同期をとる前にいくつかの時間に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e6778-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="e6778-1331">使用して、項目が同期されるまで、`itemId`エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e6778-p182">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e6778-1335">次のクライアントのさまざまな問題のある`saveAsync`の予定の作成モード。</span><span class="sxs-lookup"><span data-stu-id="e6778-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e6778-1336">Mac の Outlook をサポートしていない`saveAsync`での会議では、作成モードです。</span><span class="sxs-lookup"><span data-stu-id="e6778-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="e6778-1337">呼び出す`saveAsync`Mac の Outlook で会議のエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="e6778-1338">Web 上で outlook が常に招待状を送信または更新する場合`saveAsync`予定で作成モードです。</span><span class="sxs-lookup"><span data-stu-id="e6778-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1339">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1339">Parameters:</span></span>

|<span data-ttu-id="e6778-1340">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1340">Name</span></span>|<span data-ttu-id="e6778-1341">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1341">Type</span></span>|<span data-ttu-id="e6778-1342">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1342">Attributes</span></span>|<span data-ttu-id="e6778-1343">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e6778-1344">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e6778-1344">Object</span></span>|<span data-ttu-id="e6778-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1346">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-1347">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1347">Object</span></span>|<span data-ttu-id="e6778-1348">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1349">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e6778-1350">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1350">function</span></span>||<span data-ttu-id="e6778-1351">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e6778-1352">成功した場合、項目の識別子が提供されている、`asyncResult.value`プロパティ。</span><span class="sxs-lookup"><span data-stu-id="e6778-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1353">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1353">Requirements</span></span>

|<span data-ttu-id="e6778-1354">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1354">Requirement</span></span>|<span data-ttu-id="e6778-1355">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="e6778-1357">1.3</span></span>|
|[<span data-ttu-id="e6778-1358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-1360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1361">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e6778-1362">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="e6778-p184">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e6778-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e6778-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e6778-1366">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="e6778-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e6778-p185">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e6778-1370">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e6778-1370">Parameters:</span></span>

|<span data-ttu-id="e6778-1371">名前</span><span class="sxs-lookup"><span data-stu-id="e6778-1371">Name</span></span>|<span data-ttu-id="e6778-1372">型</span><span class="sxs-lookup"><span data-stu-id="e6778-1372">Type</span></span>|<span data-ttu-id="e6778-1373">属性</span><span class="sxs-lookup"><span data-stu-id="e6778-1373">Attributes</span></span>|<span data-ttu-id="e6778-1374">説明</span><span class="sxs-lookup"><span data-stu-id="e6778-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="e6778-1375">String</span><span class="sxs-lookup"><span data-stu-id="e6778-1375">String</span></span>||<span data-ttu-id="e6778-p186">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="e6778-1379">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1379">Object</span></span>|<span data-ttu-id="e6778-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1381">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e6778-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e6778-1382">Object</span><span class="sxs-lookup"><span data-stu-id="e6778-1382">Object</span></span>|<span data-ttu-id="e6778-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-1384">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e6778-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e6778-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e6778-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e6778-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="e6778-p187">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e6778-p188">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e6778-1391">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="e6778-1392">function</span><span class="sxs-lookup"><span data-stu-id="e6778-1392">function</span></span>||<span data-ttu-id="e6778-1393">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e6778-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e6778-1394">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1394">Requirements</span></span>

|<span data-ttu-id="e6778-1395">要件</span><span class="sxs-lookup"><span data-stu-id="e6778-1395">Requirement</span></span>|<span data-ttu-id="e6778-1396">値</span><span class="sxs-lookup"><span data-stu-id="e6778-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="e6778-1397">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e6778-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e6778-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="e6778-1398">1.2</span></span>|
|[<span data-ttu-id="e6778-1399">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e6778-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e6778-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e6778-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="e6778-1401">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e6778-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e6778-1402">作成</span><span class="sxs-lookup"><span data-stu-id="e6778-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e6778-1403">例</span><span class="sxs-lookup"><span data-stu-id="e6778-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```