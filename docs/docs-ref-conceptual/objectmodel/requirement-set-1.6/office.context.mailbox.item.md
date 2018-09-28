
# <a name="item"></a><span data-ttu-id="b101f-101">item</span><span class="sxs-lookup"><span data-stu-id="b101f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b101f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b101f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b101f-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-105">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-105">Requirements</span></span>

|<span data-ttu-id="b101f-106">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-106">Requirement</span></span>| <span data-ttu-id="b101f-107">値</span><span class="sxs-lookup"><span data-stu-id="b101f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-109">1.0</span></span>|
|[<span data-ttu-id="b101f-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="b101f-111">Restricted</span></span>|
|[<span data-ttu-id="b101f-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b101f-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-114">Members and methods</span></span>

| <span data-ttu-id="b101f-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-115">Member</span></span> | <span data-ttu-id="b101f-116">種類</span><span class="sxs-lookup"><span data-stu-id="b101f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b101f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="b101f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="b101f-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-118">Member</span></span> |
| [<span data-ttu-id="b101f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="b101f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b101f-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-120">Member</span></span> |
| [<span data-ttu-id="b101f-121">body</span><span class="sxs-lookup"><span data-stu-id="b101f-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="b101f-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-122">Member</span></span> |
| [<span data-ttu-id="b101f-123">cc</span><span class="sxs-lookup"><span data-stu-id="b101f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b101f-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-124">Member</span></span> |
| [<span data-ttu-id="b101f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="b101f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="b101f-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-126">Member</span></span> |
| [<span data-ttu-id="b101f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="b101f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="b101f-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-128">Member</span></span> |
| [<span data-ttu-id="b101f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b101f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="b101f-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-130">Member</span></span> |
| [<span data-ttu-id="b101f-131">end</span><span class="sxs-lookup"><span data-stu-id="b101f-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="b101f-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-132">Member</span></span> |
| [<span data-ttu-id="b101f-133">from</span><span class="sxs-lookup"><span data-stu-id="b101f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b101f-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-134">Member</span></span> |
| [<span data-ttu-id="b101f-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="b101f-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="b101f-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-136">Member</span></span> |
| [<span data-ttu-id="b101f-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="b101f-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="b101f-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-138">Member</span></span> |
| [<span data-ttu-id="b101f-139">itemId</span><span class="sxs-lookup"><span data-stu-id="b101f-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="b101f-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-140">Member</span></span> |
| [<span data-ttu-id="b101f-141">itemType</span><span class="sxs-lookup"><span data-stu-id="b101f-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="b101f-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-142">Member</span></span> |
| [<span data-ttu-id="b101f-143">location</span><span class="sxs-lookup"><span data-stu-id="b101f-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="b101f-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-144">Member</span></span> |
| [<span data-ttu-id="b101f-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="b101f-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="b101f-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-146">Member</span></span> |
| [<span data-ttu-id="b101f-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="b101f-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="b101f-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-148">Member</span></span> |
| [<span data-ttu-id="b101f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b101f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b101f-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-150">Member</span></span> |
| [<span data-ttu-id="b101f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="b101f-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b101f-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-152">Member</span></span> |
| [<span data-ttu-id="b101f-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b101f-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b101f-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-154">Member</span></span> |
| [<span data-ttu-id="b101f-155">sender</span><span class="sxs-lookup"><span data-stu-id="b101f-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="b101f-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-156">Member</span></span> |
| [<span data-ttu-id="b101f-157">start</span><span class="sxs-lookup"><span data-stu-id="b101f-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="b101f-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-158">Member</span></span> |
| [<span data-ttu-id="b101f-159">subject</span><span class="sxs-lookup"><span data-stu-id="b101f-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="b101f-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-160">Member</span></span> |
| [<span data-ttu-id="b101f-161">to</span><span class="sxs-lookup"><span data-stu-id="b101f-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="b101f-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-162">Member</span></span> |
| [<span data-ttu-id="b101f-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="b101f-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-164">Method</span></span> |
| [<span data-ttu-id="b101f-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="b101f-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-166">Method</span></span> |
| [<span data-ttu-id="b101f-167">close</span><span class="sxs-lookup"><span data-stu-id="b101f-167">close</span></span>](#close) | <span data-ttu-id="b101f-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-168">Method</span></span> |
| [<span data-ttu-id="b101f-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b101f-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="b101f-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-170">Method</span></span> |
| [<span data-ttu-id="b101f-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b101f-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="b101f-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-172">Method</span></span> |
| [<span data-ttu-id="b101f-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="b101f-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="b101f-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-174">Method</span></span> |
| [<span data-ttu-id="b101f-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b101f-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="b101f-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-176">Method</span></span> |
| [<span data-ttu-id="b101f-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b101f-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="b101f-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-178">Method</span></span> |
| [<span data-ttu-id="b101f-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b101f-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="b101f-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-180">Method</span></span> |
| [<span data-ttu-id="b101f-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b101f-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="b101f-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-182">Method</span></span> |
| [<span data-ttu-id="b101f-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="b101f-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-184">Method</span></span> |
| [<span data-ttu-id="b101f-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="b101f-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="b101f-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-186">Method</span></span> |
| [<span data-ttu-id="b101f-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b101f-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="b101f-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-188">Method</span></span> |
| [<span data-ttu-id="b101f-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="b101f-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-190">Method</span></span> |
| [<span data-ttu-id="b101f-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="b101f-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-192">Method</span></span> |
| [<span data-ttu-id="b101f-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="b101f-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-194">Method</span></span> |
| [<span data-ttu-id="b101f-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b101f-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="b101f-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="b101f-197">例</span><span class="sxs-lookup"><span data-stu-id="b101f-197">Example</span></span>

<span data-ttu-id="b101f-198">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b101f-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b101f-199">メンバー</span><span class="sxs-lookup"><span data-stu-id="b101f-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="b101f-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b101f-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="b101f-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-203">ファイルの特定の種類は、潜在的なセキュリティの問題により、Outlook によってブロックされは返されません。</span><span class="sxs-lookup"><span data-stu-id="b101f-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b101f-204">詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b101f-204">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-205">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-205">Type:</span></span>

*   <span data-ttu-id="b101f-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b101f-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-207">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-207">Requirements</span></span>

|<span data-ttu-id="b101f-208">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-208">Requirement</span></span>| <span data-ttu-id="b101f-209">値</span><span class="sxs-lookup"><span data-stu-id="b101f-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-211">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-211">1.0</span></span>|
|[<span data-ttu-id="b101f-212">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-213">ReadItem</span></span>|
|[<span data-ttu-id="b101f-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-216">例</span><span class="sxs-lookup"><span data-stu-id="b101f-216">Example</span></span>

<span data-ttu-id="b101f-217">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="b101f-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b101f-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b101f-219">取得またはメッセージの bcc (ブラインド カーボン コピー) 受信者を更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b101f-220">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-221">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-221">Type:</span></span>

*   [<span data-ttu-id="b101f-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="b101f-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b101f-223">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-223">Requirements</span></span>

|<span data-ttu-id="b101f-224">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-224">Requirement</span></span>| <span data-ttu-id="b101f-225">値</span><span class="sxs-lookup"><span data-stu-id="b101f-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-227">1.1</span><span class="sxs-lookup"><span data-stu-id="b101f-227">1.1</span></span>|
|[<span data-ttu-id="b101f-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-229">ReadItem</span></span>|
|[<span data-ttu-id="b101f-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-231">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-232">例</span><span class="sxs-lookup"><span data-stu-id="b101f-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="b101f-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="b101f-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="b101f-234">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-235">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-235">Type:</span></span>

*   [<span data-ttu-id="b101f-236">Body</span><span class="sxs-lookup"><span data-stu-id="b101f-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="b101f-237">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-237">Requirements</span></span>

|<span data-ttu-id="b101f-238">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-238">Requirement</span></span>| <span data-ttu-id="b101f-239">値</span><span class="sxs-lookup"><span data-stu-id="b101f-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-241">1.1</span><span class="sxs-lookup"><span data-stu-id="b101f-241">1.1</span></span>|
|[<span data-ttu-id="b101f-242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-243">ReadItem</span></span>|
|[<span data-ttu-id="b101f-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-245">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b101f-246">[cc]: 配列 <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook_1_6/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="b101f-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b101f-247">メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b101f-248">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b101f-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-249">Read mode</span></span>

<span data-ttu-id="b101f-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-252">Compose mode</span></span>

<span data-ttu-id="b101f-253">`cc`を`Recipients`オブジェクトを取得または、メッセージの**Cc**行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-254">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-254">Type:</span></span>

*   <span data-ttu-id="b101f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-256">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-256">Requirements</span></span>

|<span data-ttu-id="b101f-257">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-257">Requirement</span></span>| <span data-ttu-id="b101f-258">値</span><span class="sxs-lookup"><span data-stu-id="b101f-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-260">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-260">1.0</span></span>|
|[<span data-ttu-id="b101f-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-262">ReadItem</span></span>|
|[<span data-ttu-id="b101f-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-264">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-265">例</span><span class="sxs-lookup"><span data-stu-id="b101f-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b101f-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b101f-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="b101f-267">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b101f-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="b101f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b101f-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-272">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-272">Type:</span></span>

*   <span data-ttu-id="b101f-273">String</span><span class="sxs-lookup"><span data-stu-id="b101f-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-274">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-274">Requirements</span></span>

|<span data-ttu-id="b101f-275">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-275">Requirement</span></span>| <span data-ttu-id="b101f-276">値</span><span class="sxs-lookup"><span data-stu-id="b101f-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-278">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-278">1.0</span></span>|
|[<span data-ttu-id="b101f-279">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-280">ReadItem</span></span>|
|[<span data-ttu-id="b101f-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-282">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b101f-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b101f-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="b101f-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-286">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-286">Type:</span></span>

*   <span data-ttu-id="b101f-287">日付</span><span class="sxs-lookup"><span data-stu-id="b101f-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-288">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-288">Requirements</span></span>

|<span data-ttu-id="b101f-289">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-289">Requirement</span></span>| <span data-ttu-id="b101f-290">値</span><span class="sxs-lookup"><span data-stu-id="b101f-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-292">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-292">1.0</span></span>|
|[<span data-ttu-id="b101f-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-294">ReadItem</span></span>|
|[<span data-ttu-id="b101f-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-297">例</span><span class="sxs-lookup"><span data-stu-id="b101f-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b101f-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b101f-298">dateTimeModified :Date</span></span>

<span data-ttu-id="b101f-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-301">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-302">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-302">Type:</span></span>

*   <span data-ttu-id="b101f-303">日付</span><span class="sxs-lookup"><span data-stu-id="b101f-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-304">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-304">Requirements</span></span>

|<span data-ttu-id="b101f-305">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-305">Requirement</span></span>| <span data-ttu-id="b101f-306">値</span><span class="sxs-lookup"><span data-stu-id="b101f-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-308">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-308">1.0</span></span>|
|[<span data-ttu-id="b101f-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-310">ReadItem</span></span>|
|[<span data-ttu-id="b101f-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-313">例</span><span class="sxs-lookup"><span data-stu-id="b101f-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="b101f-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b101f-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="b101f-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b101f-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-318">Read mode</span></span>

<span data-ttu-id="b101f-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-320">Compose mode</span></span>

<span data-ttu-id="b101f-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b101f-322">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b101f-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-323">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-323">Type:</span></span>

*   <span data-ttu-id="b101f-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b101f-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-325">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-325">Requirements</span></span>

|<span data-ttu-id="b101f-326">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-326">Requirement</span></span>| <span data-ttu-id="b101f-327">値</span><span class="sxs-lookup"><span data-stu-id="b101f-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-329">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-329">1.0</span></span>|
|[<span data-ttu-id="b101f-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-331">ReadItem</span></span>|
|[<span data-ttu-id="b101f-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-333">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-334">例</span><span class="sxs-lookup"><span data-stu-id="b101f-334">Example</span></span>

<span data-ttu-id="b101f-335">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b101f-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b101f-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b101f-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b101f-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-341">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`from`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b101f-341">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-342">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-342">Type:</span></span>

*   [<span data-ttu-id="b101f-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b101f-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b101f-344">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-344">Requirements</span></span>

|<span data-ttu-id="b101f-345">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-345">Requirement</span></span>| <span data-ttu-id="b101f-346">値</span><span class="sxs-lookup"><span data-stu-id="b101f-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-347">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-348">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-348">1.0</span></span>|
|[<span data-ttu-id="b101f-349">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-350">ReadItem</span></span>|
|[<span data-ttu-id="b101f-351">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-352">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b101f-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b101f-353">internetMessageId :String</span></span>

<span data-ttu-id="b101f-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-356">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-356">Type:</span></span>

*   <span data-ttu-id="b101f-357">String</span><span class="sxs-lookup"><span data-stu-id="b101f-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-358">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-358">Requirements</span></span>

|<span data-ttu-id="b101f-359">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-359">Requirement</span></span>| <span data-ttu-id="b101f-360">値</span><span class="sxs-lookup"><span data-stu-id="b101f-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-362">1.0</span></span>|
|[<span data-ttu-id="b101f-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-364">ReadItem</span></span>|
|[<span data-ttu-id="b101f-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-367">例</span><span class="sxs-lookup"><span data-stu-id="b101f-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b101f-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b101f-368">itemClass :String</span></span>

<span data-ttu-id="b101f-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b101f-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b101f-373">種類</span><span class="sxs-lookup"><span data-stu-id="b101f-373">Type</span></span> | <span data-ttu-id="b101f-374">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-374">Description</span></span> | <span data-ttu-id="b101f-375">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="b101f-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b101f-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="b101f-376">Appointment items</span></span> | <span data-ttu-id="b101f-377">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b101f-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="b101f-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="b101f-378">Message items</span></span> | <span data-ttu-id="b101f-379">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b101f-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-381">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-381">Type:</span></span>

*   <span data-ttu-id="b101f-382">String</span><span class="sxs-lookup"><span data-stu-id="b101f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-383">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-383">Requirements</span></span>

|<span data-ttu-id="b101f-384">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-384">Requirement</span></span>| <span data-ttu-id="b101f-385">値</span><span class="sxs-lookup"><span data-stu-id="b101f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-387">1.0</span></span>|
|[<span data-ttu-id="b101f-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-389">ReadItem</span></span>|
|[<span data-ttu-id="b101f-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-392">例</span><span class="sxs-lookup"><span data-stu-id="b101f-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b101f-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b101f-393">(nullable) itemId :String</span></span>

<span data-ttu-id="b101f-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="b101f-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b101f-397">`itemId`プロパティは、Outlook のエントリ ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="b101f-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b101f-398">この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b101f-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b101f-399">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b101f-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b101f-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-402">種類:</span><span class="sxs-lookup"><span data-stu-id="b101f-402">Type:</span></span>

*   <span data-ttu-id="b101f-403">String</span><span class="sxs-lookup"><span data-stu-id="b101f-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-404">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-404">Requirements</span></span>

|<span data-ttu-id="b101f-405">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-405">Requirement</span></span>| <span data-ttu-id="b101f-406">値</span><span class="sxs-lookup"><span data-stu-id="b101f-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-408">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-408">1.0</span></span>|
|[<span data-ttu-id="b101f-409">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-410">ReadItem</span></span>|
|[<span data-ttu-id="b101f-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-413">例</span><span class="sxs-lookup"><span data-stu-id="b101f-413">Example</span></span>

<span data-ttu-id="b101f-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="b101f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b101f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b101f-417">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b101f-418">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="b101f-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-419">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-419">Type:</span></span>

*   [<span data-ttu-id="b101f-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b101f-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b101f-421">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-421">Requirements</span></span>

|<span data-ttu-id="b101f-422">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-422">Requirement</span></span>| <span data-ttu-id="b101f-423">値</span><span class="sxs-lookup"><span data-stu-id="b101f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-425">1.0</span></span>|
|[<span data-ttu-id="b101f-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-427">ReadItem</span></span>|
|[<span data-ttu-id="b101f-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-429">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-430">例</span><span class="sxs-lookup"><span data-stu-id="b101f-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="b101f-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="b101f-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="b101f-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-433">Read mode</span></span>

<span data-ttu-id="b101f-434">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-435">Compose mode</span></span>

<span data-ttu-id="b101f-436">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-437">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-437">Type:</span></span>

*   <span data-ttu-id="b101f-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="b101f-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-439">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-439">Requirements</span></span>

|<span data-ttu-id="b101f-440">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-440">Requirement</span></span>| <span data-ttu-id="b101f-441">値</span><span class="sxs-lookup"><span data-stu-id="b101f-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-443">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-443">1.0</span></span>|
|[<span data-ttu-id="b101f-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-445">ReadItem</span></span>|
|[<span data-ttu-id="b101f-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-447">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-448">例</span><span class="sxs-lookup"><span data-stu-id="b101f-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b101f-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b101f-449">normalizedSubject :String</span></span>

<span data-ttu-id="b101f-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b101f-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-454">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-454">Type:</span></span>

*   <span data-ttu-id="b101f-455">String</span><span class="sxs-lookup"><span data-stu-id="b101f-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-456">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-456">Requirements</span></span>

|<span data-ttu-id="b101f-457">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-457">Requirement</span></span>| <span data-ttu-id="b101f-458">値</span><span class="sxs-lookup"><span data-stu-id="b101f-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-459">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-460">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-460">1.0</span></span>|
|[<span data-ttu-id="b101f-461">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-462">ReadItem</span></span>|
|[<span data-ttu-id="b101f-463">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-464">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-465">例</span><span class="sxs-lookup"><span data-stu-id="b101f-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="b101f-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b101f-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="b101f-467">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-468">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-468">Type:</span></span>

*   [<span data-ttu-id="b101f-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b101f-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b101f-470">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-470">Requirements</span></span>

|<span data-ttu-id="b101f-471">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-471">Requirement</span></span>| <span data-ttu-id="b101f-472">値</span><span class="sxs-lookup"><span data-stu-id="b101f-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-473">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-474">1.3</span><span class="sxs-lookup"><span data-stu-id="b101f-474">1.3</span></span>|
|[<span data-ttu-id="b101f-475">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-476">ReadItem</span></span>|
|[<span data-ttu-id="b101f-477">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-478">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b101f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b101f-480">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b101f-481">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b101f-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-482">Read mode</span></span>

<span data-ttu-id="b101f-483">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-484">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-484">Compose mode</span></span>

<span data-ttu-id="b101f-485">`optionalAttendees`を`Recipients`オブジェクトを取得または省略可能な会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-486">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-486">Type:</span></span>

*   <span data-ttu-id="b101f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-488">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-488">Requirements</span></span>

|<span data-ttu-id="b101f-489">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-489">Requirement</span></span>| <span data-ttu-id="b101f-490">値</span><span class="sxs-lookup"><span data-stu-id="b101f-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-491">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-492">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-492">1.0</span></span>|
|[<span data-ttu-id="b101f-493">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-494">ReadItem</span></span>|
|[<span data-ttu-id="b101f-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-496">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-497">例</span><span class="sxs-lookup"><span data-stu-id="b101f-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b101f-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b101f-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b101f-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-501">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-501">Type:</span></span>

*   [<span data-ttu-id="b101f-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b101f-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b101f-503">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-503">Requirements</span></span>

|<span data-ttu-id="b101f-504">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-504">Requirement</span></span>| <span data-ttu-id="b101f-505">値</span><span class="sxs-lookup"><span data-stu-id="b101f-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-507">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-507">1.0</span></span>|
|[<span data-ttu-id="b101f-508">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-509">ReadItem</span></span>|
|[<span data-ttu-id="b101f-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-511">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-512">例</span><span class="sxs-lookup"><span data-stu-id="b101f-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b101f-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b101f-514">イベントの出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b101f-515">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b101f-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-516">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-516">Read mode</span></span>

<span data-ttu-id="b101f-517">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-518">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-518">Compose mode</span></span>

<span data-ttu-id="b101f-519">`requiredAttendees`を`Recipients`オブジェクトを取得または会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-520">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-520">Type:</span></span>

*   <span data-ttu-id="b101f-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-522">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-522">Requirements</span></span>

|<span data-ttu-id="b101f-523">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-523">Requirement</span></span>| <span data-ttu-id="b101f-524">値</span><span class="sxs-lookup"><span data-stu-id="b101f-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-525">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-526">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-526">1.0</span></span>|
|[<span data-ttu-id="b101f-527">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-528">ReadItem</span></span>|
|[<span data-ttu-id="b101f-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-530">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-531">例</span><span class="sxs-lookup"><span data-stu-id="b101f-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="b101f-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b101f-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="b101f-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="b101f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b101f-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-537">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`sender`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b101f-537">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-538">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-538">Type:</span></span>

*   [<span data-ttu-id="b101f-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b101f-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b101f-540">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-540">Requirements</span></span>

|<span data-ttu-id="b101f-541">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-541">Requirement</span></span>| <span data-ttu-id="b101f-542">値</span><span class="sxs-lookup"><span data-stu-id="b101f-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-543">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-544">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-544">1.0</span></span>|
|[<span data-ttu-id="b101f-545">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-546">ReadItem</span></span>|
|[<span data-ttu-id="b101f-547">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-548">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-549">例</span><span class="sxs-lookup"><span data-stu-id="b101f-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="b101f-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b101f-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="b101f-551">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b101f-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-554">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-554">Read mode</span></span>

<span data-ttu-id="b101f-555">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-556">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-556">Compose mode</span></span>

<span data-ttu-id="b101f-557">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b101f-558">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b101f-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-559">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-559">Type:</span></span>

*   <span data-ttu-id="b101f-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="b101f-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-561">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-561">Requirements</span></span>

|<span data-ttu-id="b101f-562">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-562">Requirement</span></span>| <span data-ttu-id="b101f-563">値</span><span class="sxs-lookup"><span data-stu-id="b101f-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-565">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-565">1.0</span></span>|
|[<span data-ttu-id="b101f-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-567">ReadItem</span></span>|
|[<span data-ttu-id="b101f-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-569">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-570">例</span><span class="sxs-lookup"><span data-stu-id="b101f-570">Example</span></span>

<span data-ttu-id="b101f-571">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="b101f-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b101f-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="b101f-573">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b101f-574">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="b101f-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-575">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-575">Read mode</span></span>

<span data-ttu-id="b101f-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b101f-578">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-578">Compose mode</span></span>

<span data-ttu-id="b101f-579">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b101f-580">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-580">Type:</span></span>

*   <span data-ttu-id="b101f-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b101f-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-582">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-582">Requirements</span></span>

|<span data-ttu-id="b101f-583">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-583">Requirement</span></span>| <span data-ttu-id="b101f-584">値</span><span class="sxs-lookup"><span data-stu-id="b101f-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-585">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-586">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-586">1.0</span></span>|
|[<span data-ttu-id="b101f-587">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-588">ReadItem</span></span>|
|[<span data-ttu-id="b101f-589">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-590">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="b101f-591">: 配列 <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook_1_6/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="b101f-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="b101f-592">[メッセージの [**宛先**] 行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b101f-593">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="b101f-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b101f-594">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="b101f-594">Read mode</span></span>

<span data-ttu-id="b101f-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b101f-597">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="b101f-597">Compose mode</span></span>

<span data-ttu-id="b101f-598">`to`を`Recipients`オブジェクトを取得または、メッセージの [**宛先**] 行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b101f-598">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b101f-599">型:</span><span class="sxs-lookup"><span data-stu-id="b101f-599">Type:</span></span>

*   <span data-ttu-id="b101f-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b101f-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-601">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-601">Requirements</span></span>

|<span data-ttu-id="b101f-602">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-602">Requirement</span></span>| <span data-ttu-id="b101f-603">値</span><span class="sxs-lookup"><span data-stu-id="b101f-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-604">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-605">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-605">1.0</span></span>|
|[<span data-ttu-id="b101f-606">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-607">ReadItem</span></span>|
|[<span data-ttu-id="b101f-608">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-609">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-610">例</span><span class="sxs-lookup"><span data-stu-id="b101f-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b101f-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="b101f-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b101f-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b101f-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b101f-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="b101f-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b101f-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="b101f-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b101f-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-616">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-616">Parameters:</span></span>

|<span data-ttu-id="b101f-617">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-617">Name</span></span>| <span data-ttu-id="b101f-618">型</span><span class="sxs-lookup"><span data-stu-id="b101f-618">Type</span></span>| <span data-ttu-id="b101f-619">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-619">Attributes</span></span>| <span data-ttu-id="b101f-620">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b101f-621">String</span><span class="sxs-lookup"><span data-stu-id="b101f-621">String</span></span>||<span data-ttu-id="b101f-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b101f-624">String</span><span class="sxs-lookup"><span data-stu-id="b101f-624">String</span></span>||<span data-ttu-id="b101f-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b101f-627">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-627">Object</span></span>| <span data-ttu-id="b101f-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-628">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="b101f-630">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-630">Object</span></span> | <span data-ttu-id="b101f-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-631">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="b101f-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="b101f-633">Boolean</span></span> | <span data-ttu-id="b101f-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-634">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="b101f-636">function</span><span class="sxs-lookup"><span data-stu-id="b101f-636">function</span></span>| <span data-ttu-id="b101f-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-637">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b101f-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b101f-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b101f-641">エラー</span><span class="sxs-lookup"><span data-stu-id="b101f-641">Errors</span></span>

| <span data-ttu-id="b101f-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b101f-642">Error code</span></span> | <span data-ttu-id="b101f-643">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b101f-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="b101f-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b101f-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="b101f-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b101f-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b101f-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-647">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-647">Requirements</span></span>

|<span data-ttu-id="b101f-648">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-648">Requirement</span></span>| <span data-ttu-id="b101f-649">値</span><span class="sxs-lookup"><span data-stu-id="b101f-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-651">1.1</span><span class="sxs-lookup"><span data-stu-id="b101f-651">1.1</span></span>|
|[<span data-ttu-id="b101f-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-655">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b101f-656">例</span><span class="sxs-lookup"><span data-stu-id="b101f-656">Examples</span></span>

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

<span data-ttu-id="b101f-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="b101f-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b101f-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b101f-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b101f-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="b101f-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b101f-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b101f-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b101f-664">Office アドインは、Outlook Web App で実行されている場合、`addItemAttachmentAsync`メソッドが項目を編集しているアイテム以外のアイテムに関連付けることができますただし、これはサポートされていません、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="b101f-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-665">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-665">Parameters:</span></span>

|<span data-ttu-id="b101f-666">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-666">Name</span></span>| <span data-ttu-id="b101f-667">型</span><span class="sxs-lookup"><span data-stu-id="b101f-667">Type</span></span>| <span data-ttu-id="b101f-668">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-668">Attributes</span></span>| <span data-ttu-id="b101f-669">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b101f-670">String</span><span class="sxs-lookup"><span data-stu-id="b101f-670">String</span></span>||<span data-ttu-id="b101f-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b101f-673">String</span><span class="sxs-lookup"><span data-stu-id="b101f-673">String</span></span>||<span data-ttu-id="b101f-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b101f-676">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-676">Object</span></span>| <span data-ttu-id="b101f-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-677">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b101f-679">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-679">Object</span></span>| <span data-ttu-id="b101f-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-680">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b101f-682">function</span><span class="sxs-lookup"><span data-stu-id="b101f-682">function</span></span>| <span data-ttu-id="b101f-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-683">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b101f-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b101f-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b101f-687">エラー</span><span class="sxs-lookup"><span data-stu-id="b101f-687">Errors</span></span>

| <span data-ttu-id="b101f-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b101f-688">Error code</span></span> | <span data-ttu-id="b101f-689">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b101f-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b101f-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-691">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-691">Requirements</span></span>

|<span data-ttu-id="b101f-692">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-692">Requirement</span></span>| <span data-ttu-id="b101f-693">値</span><span class="sxs-lookup"><span data-stu-id="b101f-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-695">1.1</span><span class="sxs-lookup"><span data-stu-id="b101f-695">1.1</span></span>|
|[<span data-ttu-id="b101f-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-699">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-700">例</span><span class="sxs-lookup"><span data-stu-id="b101f-700">Example</span></span>

<span data-ttu-id="b101f-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="b101f-702">close()</span><span class="sxs-lookup"><span data-stu-id="b101f-702">close()</span></span>

<span data-ttu-id="b101f-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="b101f-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b101f-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-706">アイテム予定は、以前保存されたを使用する場合は、web 上の Outlook で`saveAsync`を求めるメッセージを保存、破棄、または、キャンセル場合でも、変更が発生していないから、項目を保存します。</span><span class="sxs-lookup"><span data-stu-id="b101f-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b101f-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="b101f-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-708">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-708">Requirements</span></span>

|<span data-ttu-id="b101f-709">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-709">Requirement</span></span>| <span data-ttu-id="b101f-710">値</span><span class="sxs-lookup"><span data-stu-id="b101f-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-712">1.3</span><span class="sxs-lookup"><span data-stu-id="b101f-712">1.3</span></span>|
|[<span data-ttu-id="b101f-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="b101f-714">Restricted</span></span>|
|[<span data-ttu-id="b101f-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-716">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b101f-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b101f-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b101f-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-719">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b101f-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="b101f-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b101f-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="b101f-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-725">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-725">Parameters:</span></span>

| <span data-ttu-id="b101f-726">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-726">Name</span></span> | <span data-ttu-id="b101f-727">型</span><span class="sxs-lookup"><span data-stu-id="b101f-727">Type</span></span> | <span data-ttu-id="b101f-728">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-728">Attributes</span></span> | <span data-ttu-id="b101f-729">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b101f-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b101f-730">String &#124; Object</span></span>| |<span data-ttu-id="b101f-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b101f-733">**または**</span><span class="sxs-lookup"><span data-stu-id="b101f-733">**OR**</span></span><br/><span data-ttu-id="b101f-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b101f-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b101f-736">String</span><span class="sxs-lookup"><span data-stu-id="b101f-736">String</span></span> | <span data-ttu-id="b101f-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-737">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b101f-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b101f-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-741">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="b101f-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b101f-743">String</span><span class="sxs-lookup"><span data-stu-id="b101f-743">String</span></span> | | <span data-ttu-id="b101f-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b101f-746">String</span><span class="sxs-lookup"><span data-stu-id="b101f-746">String</span></span> | | <span data-ttu-id="b101f-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b101f-748">String</span><span class="sxs-lookup"><span data-stu-id="b101f-748">String</span></span> | | <span data-ttu-id="b101f-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="b101f-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b101f-751">Boolean</span><span class="sxs-lookup"><span data-stu-id="b101f-751">Boolean</span></span> | | <span data-ttu-id="b101f-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b101f-754">String</span><span class="sxs-lookup"><span data-stu-id="b101f-754">String</span></span> | | <span data-ttu-id="b101f-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b101f-758">function</span><span class="sxs-lookup"><span data-stu-id="b101f-758">function</span></span> | <span data-ttu-id="b101f-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-759">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-761">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-761">Requirements</span></span>

|<span data-ttu-id="b101f-762">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-762">Requirement</span></span>| <span data-ttu-id="b101f-763">値</span><span class="sxs-lookup"><span data-stu-id="b101f-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-765">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-765">1.0</span></span>|
|[<span data-ttu-id="b101f-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-767">ReadItem</span></span>|
|[<span data-ttu-id="b101f-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b101f-770">例</span><span class="sxs-lookup"><span data-stu-id="b101f-770">Examples</span></span>

<span data-ttu-id="b101f-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="b101f-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b101f-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b101f-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b101f-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b101f-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b101f-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b101f-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b101f-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="b101f-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-779">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b101f-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="b101f-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b101f-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="b101f-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-785">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-785">Parameters:</span></span>

| <span data-ttu-id="b101f-786">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-786">Name</span></span> | <span data-ttu-id="b101f-787">型</span><span class="sxs-lookup"><span data-stu-id="b101f-787">Type</span></span> | <span data-ttu-id="b101f-788">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-788">Attributes</span></span> | <span data-ttu-id="b101f-789">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="b101f-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b101f-790">String &#124; Object</span></span>| | <span data-ttu-id="b101f-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b101f-793">**または**</span><span class="sxs-lookup"><span data-stu-id="b101f-793">**OR**</span></span><br/><span data-ttu-id="b101f-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b101f-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b101f-796">String</span><span class="sxs-lookup"><span data-stu-id="b101f-796">String</span></span> | <span data-ttu-id="b101f-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-797">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="b101f-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b101f-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b101f-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-801">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="b101f-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b101f-803">String</span><span class="sxs-lookup"><span data-stu-id="b101f-803">String</span></span> | | <span data-ttu-id="b101f-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b101f-806">String</span><span class="sxs-lookup"><span data-stu-id="b101f-806">String</span></span> | | <span data-ttu-id="b101f-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b101f-808">String</span><span class="sxs-lookup"><span data-stu-id="b101f-808">String</span></span> | | <span data-ttu-id="b101f-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="b101f-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="b101f-811">Boolean</span><span class="sxs-lookup"><span data-stu-id="b101f-811">Boolean</span></span> | | <span data-ttu-id="b101f-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b101f-814">String</span><span class="sxs-lookup"><span data-stu-id="b101f-814">String</span></span> | | <span data-ttu-id="b101f-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b101f-818">function</span><span class="sxs-lookup"><span data-stu-id="b101f-818">function</span></span> | <span data-ttu-id="b101f-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-819">&lt;optional&gt;</span></span> | <span data-ttu-id="b101f-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-821">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-821">Requirements</span></span>

|<span data-ttu-id="b101f-822">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-822">Requirement</span></span>| <span data-ttu-id="b101f-823">値</span><span class="sxs-lookup"><span data-stu-id="b101f-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-825">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-825">1.0</span></span>|
|[<span data-ttu-id="b101f-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-827">ReadItem</span></span>|
|[<span data-ttu-id="b101f-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b101f-830">例</span><span class="sxs-lookup"><span data-stu-id="b101f-830">Examples</span></span>

<span data-ttu-id="b101f-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="b101f-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b101f-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b101f-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b101f-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b101f-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b101f-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="b101f-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="b101f-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b101f-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="b101f-838">選択したアイテムの本文内のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-839">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-840">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-840">Requirements</span></span>

|<span data-ttu-id="b101f-841">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-841">Requirement</span></span>| <span data-ttu-id="b101f-842">値</span><span class="sxs-lookup"><span data-stu-id="b101f-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-844">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-844">1.0</span></span>|
|[<span data-ttu-id="b101f-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-846">ReadItem</span></span>|
|[<span data-ttu-id="b101f-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-849">Returns:</span></span>

<span data-ttu-id="b101f-850">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b101f-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b101f-851">例</span><span class="sxs-lookup"><span data-stu-id="b101f-851">Example</span></span>

<span data-ttu-id="b101f-852">次の使用例は、現在の項目の本文に連絡先のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-852">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="b101f-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b101f-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b101f-854">選択したアイテムの本文に指定されたエンティティ型のすべてのエンティティの配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-855">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-856">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-856">Parameters:</span></span>

|<span data-ttu-id="b101f-857">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-857">Name</span></span>| <span data-ttu-id="b101f-858">種類</span><span class="sxs-lookup"><span data-stu-id="b101f-858">Type</span></span>| <span data-ttu-id="b101f-859">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b101f-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b101f-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="b101f-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="b101f-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="b101f-862">Requirements</span></span>

|<span data-ttu-id="b101f-863">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-863">Requirement</span></span>| <span data-ttu-id="b101f-864">値</span><span class="sxs-lookup"><span data-stu-id="b101f-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-866">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-866">1.0</span></span>|
|[<span data-ttu-id="b101f-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="b101f-868">Restricted</span></span>|
|[<span data-ttu-id="b101f-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-871">Returns:</span></span>

<span data-ttu-id="b101f-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b101f-873">アイテムの本文に指定した型のエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b101f-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="b101f-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b101f-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="b101f-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b101f-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="b101f-876">Value of `entityType`</span></span> | <span data-ttu-id="b101f-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="b101f-877">Type of objects in returned array</span></span> | <span data-ttu-id="b101f-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="b101f-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b101f-879">文字列</span><span class="sxs-lookup"><span data-stu-id="b101f-879">String</span></span> | <span data-ttu-id="b101f-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b101f-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b101f-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="b101f-881">Contact</span></span> | <span data-ttu-id="b101f-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b101f-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b101f-883">文字列</span><span class="sxs-lookup"><span data-stu-id="b101f-883">String</span></span> | <span data-ttu-id="b101f-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b101f-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b101f-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b101f-885">MeetingSuggestion</span></span> | <span data-ttu-id="b101f-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b101f-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b101f-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b101f-887">PhoneNumber</span></span> | <span data-ttu-id="b101f-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b101f-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b101f-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b101f-889">TaskSuggestion</span></span> | <span data-ttu-id="b101f-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b101f-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b101f-891">文字列</span><span class="sxs-lookup"><span data-stu-id="b101f-891">String</span></span> | <span data-ttu-id="b101f-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="b101f-892">**Restricted**</span></span> |

<span data-ttu-id="b101f-893">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b101f-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b101f-894">例</span><span class="sxs-lookup"><span data-stu-id="b101f-894">Example</span></span>

<span data-ttu-id="b101f-895">次の例では、現在の項目の本文に郵便番号のアドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b101f-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="b101f-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b101f-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b101f-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-898">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-900">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-900">Parameters:</span></span>

|<span data-ttu-id="b101f-901">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-901">Name</span></span>| <span data-ttu-id="b101f-902">種類</span><span class="sxs-lookup"><span data-stu-id="b101f-902">Type</span></span>| <span data-ttu-id="b101f-903">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b101f-904">String</span><span class="sxs-lookup"><span data-stu-id="b101f-904">String</span></span>|<span data-ttu-id="b101f-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="b101f-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-906">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-906">Requirements</span></span>

|<span data-ttu-id="b101f-907">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-907">Requirement</span></span>| <span data-ttu-id="b101f-908">値</span><span class="sxs-lookup"><span data-stu-id="b101f-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-910">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-910">1.0</span></span>|
|[<span data-ttu-id="b101f-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-912">ReadItem</span></span>|
|[<span data-ttu-id="b101f-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-915">Returns:</span></span>

<span data-ttu-id="b101f-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b101f-918">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b101f-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b101f-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b101f-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b101f-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-921">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b101f-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="b101f-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b101f-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b101f-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b101f-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-930">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-930">Requirements</span></span>

|<span data-ttu-id="b101f-931">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-931">Requirement</span></span>| <span data-ttu-id="b101f-932">値</span><span class="sxs-lookup"><span data-stu-id="b101f-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-934">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-934">1.0</span></span>|
|[<span data-ttu-id="b101f-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-936">ReadItem</span></span>|
|[<span data-ttu-id="b101f-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-939">Returns:</span></span>

<span data-ttu-id="b101f-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="b101f-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b101f-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b101f-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b101f-943">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b101f-944">例</span><span class="sxs-lookup"><span data-stu-id="b101f-944">Example</span></span>

<span data-ttu-id="b101f-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b101f-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b101f-946">getRegExMatchesByName(name)] → [(許容) {配列。 < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="b101f-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b101f-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-948">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="b101f-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b101f-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="b101f-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-952">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-952">Parameters:</span></span>

|<span data-ttu-id="b101f-953">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-953">Name</span></span>| <span data-ttu-id="b101f-954">種類</span><span class="sxs-lookup"><span data-stu-id="b101f-954">Type</span></span>| <span data-ttu-id="b101f-955">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b101f-956">String</span><span class="sxs-lookup"><span data-stu-id="b101f-956">String</span></span>|<span data-ttu-id="b101f-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="b101f-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-958">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-958">Requirements</span></span>

|<span data-ttu-id="b101f-959">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-959">Requirement</span></span>| <span data-ttu-id="b101f-960">値</span><span class="sxs-lookup"><span data-stu-id="b101f-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-962">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-962">1.0</span></span>|
|[<span data-ttu-id="b101f-963">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-964">ReadItem</span></span>|
|[<span data-ttu-id="b101f-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-967">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-967">Returns:</span></span>

<span data-ttu-id="b101f-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="b101f-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b101f-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b101f-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b101f-970">配列。 < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="b101f-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b101f-971">例</span><span class="sxs-lookup"><span data-stu-id="b101f-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b101f-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b101f-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b101f-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b101f-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-976">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-976">Parameters:</span></span>

|<span data-ttu-id="b101f-977">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-977">Name</span></span>| <span data-ttu-id="b101f-978">型</span><span class="sxs-lookup"><span data-stu-id="b101f-978">Type</span></span>| <span data-ttu-id="b101f-979">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-979">Attributes</span></span>| <span data-ttu-id="b101f-980">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b101f-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b101f-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b101f-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b101f-985">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-985">Object</span></span>| <span data-ttu-id="b101f-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-986">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b101f-988">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-988">Object</span></span>| <span data-ttu-id="b101f-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-989">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b101f-991">function</span><span class="sxs-lookup"><span data-stu-id="b101f-991">function</span></span>||<span data-ttu-id="b101f-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b101f-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b101f-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b101f-994">選択範囲は、source プロパティにアクセスするには、呼び出す`asyncResult.value.sourceProperty`、いずれかの方法となる`body`または`subject`。</span><span class="sxs-lookup"><span data-stu-id="b101f-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-995">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-995">Requirements</span></span>

|<span data-ttu-id="b101f-996">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-996">Requirement</span></span>| <span data-ttu-id="b101f-997">値</span><span class="sxs-lookup"><span data-stu-id="b101f-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-999">1.2</span><span class="sxs-lookup"><span data-stu-id="b101f-999">1.2</span></span>|
|[<span data-ttu-id="b101f-1000">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1003">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-1004">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-1004">Returns:</span></span>

<span data-ttu-id="b101f-1005">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="b101f-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b101f-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b101f-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b101f-1007">String</span><span class="sxs-lookup"><span data-stu-id="b101f-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b101f-1008">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="b101f-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b101f-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="b101f-p163">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-1012">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-1013">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1013">Requirements</span></span>

|<span data-ttu-id="b101f-1014">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1014">Requirement</span></span>| <span data-ttu-id="b101f-1015">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="b101f-1017">1.6</span></span> |
|[<span data-ttu-id="b101f-1018">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1019">ReadItem</span></span>|
|[<span data-ttu-id="b101f-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-1022">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-1022">Returns:</span></span>

<span data-ttu-id="b101f-1023">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b101f-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b101f-1024">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1024">Example</span></span>

<span data-ttu-id="b101f-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="b101f-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="b101f-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b101f-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="b101f-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-1029">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b101f-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b101f-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b101f-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="b101f-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b101f-1034">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="b101f-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b101f-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b101f-1038">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1038">Requirements</span></span>

|<span data-ttu-id="b101f-1039">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1039">Requirement</span></span>| <span data-ttu-id="b101f-1040">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="b101f-1042">1.6</span></span> |
|[<span data-ttu-id="b101f-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1044">ReadItem</span></span>|
|[<span data-ttu-id="b101f-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b101f-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b101f-1047">Returns:</span></span>

<span data-ttu-id="b101f-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="b101f-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="b101f-1050">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1050">Example</span></span>

<span data-ttu-id="b101f-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b101f-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b101f-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b101f-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b101f-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b101f-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="b101f-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-1057">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-1057">Parameters:</span></span>

|<span data-ttu-id="b101f-1058">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-1058">Name</span></span>| <span data-ttu-id="b101f-1059">型</span><span class="sxs-lookup"><span data-stu-id="b101f-1059">Type</span></span>| <span data-ttu-id="b101f-1060">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-1060">Attributes</span></span>| <span data-ttu-id="b101f-1061">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b101f-1062">function</span><span class="sxs-lookup"><span data-stu-id="b101f-1062">function</span></span>||<span data-ttu-id="b101f-1063">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b101f-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b101f-1065">取得し、アイテムのカスタム プロパティを削除してサーバーにバックアップを設定するカスタム プロパティに対する変更を保存するのには、このオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b101f-1066">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b101f-1066">Object</span></span>| <span data-ttu-id="b101f-1067">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1068">開発者は、コールバック関数にアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b101f-1069">によってこのオブジェクトにアクセスできる、`asyncResult.asyncContext`コールバック関数のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="b101f-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-1070">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1070">Requirements</span></span>

|<span data-ttu-id="b101f-1071">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1071">Requirement</span></span>| <span data-ttu-id="b101f-1072">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="b101f-1074">1.0</span></span>|
|[<span data-ttu-id="b101f-1075">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1076">ReadItem</span></span>|
|[<span data-ttu-id="b101f-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1078">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b101f-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-1079">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1079">Example</span></span>

<span data-ttu-id="b101f-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b101f-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b101f-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b101f-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="b101f-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b101f-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="b101f-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-1089">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-1089">Parameters:</span></span>

|<span data-ttu-id="b101f-1090">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-1090">Name</span></span>| <span data-ttu-id="b101f-1091">型</span><span class="sxs-lookup"><span data-stu-id="b101f-1091">Type</span></span>| <span data-ttu-id="b101f-1092">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-1092">Attributes</span></span>| <span data-ttu-id="b101f-1093">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b101f-1094">String</span><span class="sxs-lookup"><span data-stu-id="b101f-1094">String</span></span>||<span data-ttu-id="b101f-p173">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="b101f-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="b101f-1097">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-1097">Object</span></span>| <span data-ttu-id="b101f-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1099">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b101f-1100">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-1100">Object</span></span>| <span data-ttu-id="b101f-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1102">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b101f-1103">function</span><span class="sxs-lookup"><span data-stu-id="b101f-1103">function</span></span>| <span data-ttu-id="b101f-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1105">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b101f-1106">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b101f-1107">エラー</span><span class="sxs-lookup"><span data-stu-id="b101f-1107">Errors</span></span>

| <span data-ttu-id="b101f-1108">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b101f-1108">Error code</span></span> | <span data-ttu-id="b101f-1109">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b101f-1110">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="b101f-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-1111">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1111">Requirements</span></span>

|<span data-ttu-id="b101f-1112">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1112">Requirement</span></span>| <span data-ttu-id="b101f-1113">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1114">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="b101f-1115">1.1</span></span>|
|[<span data-ttu-id="b101f-1116">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-1118">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1119">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-1120">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1120">Example</span></span>

<span data-ttu-id="b101f-1121">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="b101f-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b101f-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b101f-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="b101f-1123">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="b101f-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="b101f-p174">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-1127">アドインを呼び出す場合は、`saveAsync`内のアイテムの作成モードを取得するのには、 `itemId` EWS または REST API を使用するにすると、Outlook キャッシュ モードでは、かかる場合がある項目が実際には、サーバーと同期をとる前にいくつかの時間に注意してください。</span><span class="sxs-lookup"><span data-stu-id="b101f-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b101f-1128">使用して、項目が同期されるまで、`itemId`エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b101f-p176">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b101f-1132">次のクライアントのさまざまな問題のある`saveAsync`の予定の作成モード。</span><span class="sxs-lookup"><span data-stu-id="b101f-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b101f-1133">Mac の Outlook をサポートしていない`saveAsync`での会議では、作成モードです。</span><span class="sxs-lookup"><span data-stu-id="b101f-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b101f-1134">呼び出す`saveAsync`Mac の Outlook で会議のエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b101f-1135">Web 上で outlook が常に招待状を送信または更新する場合`saveAsync`予定で作成モードです。</span><span class="sxs-lookup"><span data-stu-id="b101f-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-1136">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-1136">Parameters:</span></span>

|<span data-ttu-id="b101f-1137">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-1137">Name</span></span>| <span data-ttu-id="b101f-1138">型</span><span class="sxs-lookup"><span data-stu-id="b101f-1138">Type</span></span>| <span data-ttu-id="b101f-1139">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-1139">Attributes</span></span>| <span data-ttu-id="b101f-1140">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="b101f-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b101f-1141">Object</span></span>| <span data-ttu-id="b101f-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1143">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b101f-1144">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-1144">Object</span></span>| <span data-ttu-id="b101f-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1146">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b101f-1147">function</span><span class="sxs-lookup"><span data-stu-id="b101f-1147">function</span></span>||<span data-ttu-id="b101f-1148">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b101f-1149">成功した場合、項目の識別子が提供されている、`asyncResult.value`プロパティ。</span><span class="sxs-lookup"><span data-stu-id="b101f-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b101f-1150">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1150">Requirements</span></span>

|<span data-ttu-id="b101f-1151">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1151">Requirement</span></span>| <span data-ttu-id="b101f-1152">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="b101f-1154">1.3</span></span>|
|[<span data-ttu-id="b101f-1155">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-1157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1158">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b101f-1159">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b101f-p178">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b101f-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b101f-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b101f-1163">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="b101f-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b101f-p179">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b101f-1167">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="b101f-1167">Parameters:</span></span>

|<span data-ttu-id="b101f-1168">名前</span><span class="sxs-lookup"><span data-stu-id="b101f-1168">Name</span></span>| <span data-ttu-id="b101f-1169">型</span><span class="sxs-lookup"><span data-stu-id="b101f-1169">Type</span></span>| <span data-ttu-id="b101f-1170">属性</span><span class="sxs-lookup"><span data-stu-id="b101f-1170">Attributes</span></span>| <span data-ttu-id="b101f-1171">説明</span><span class="sxs-lookup"><span data-stu-id="b101f-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b101f-1172">String</span><span class="sxs-lookup"><span data-stu-id="b101f-1172">String</span></span>||<span data-ttu-id="b101f-p180">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b101f-1176">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-1176">Object</span></span>| <span data-ttu-id="b101f-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1178">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b101f-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b101f-1179">Object</span><span class="sxs-lookup"><span data-stu-id="b101f-1179">Object</span></span>| <span data-ttu-id="b101f-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-1181">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b101f-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b101f-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b101f-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b101f-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="b101f-p181">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b101f-p182">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b101f-1188">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b101f-1189">function</span><span class="sxs-lookup"><span data-stu-id="b101f-1189">function</span></span>||<span data-ttu-id="b101f-1190">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b101f-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b101f-1191">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1191">Requirements</span></span>

|<span data-ttu-id="b101f-1192">要件</span><span class="sxs-lookup"><span data-stu-id="b101f-1192">Requirement</span></span>| <span data-ttu-id="b101f-1193">値</span><span class="sxs-lookup"><span data-stu-id="b101f-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="b101f-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b101f-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b101f-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="b101f-1195">1.2</span></span>|
|[<span data-ttu-id="b101f-1196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b101f-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b101f-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b101f-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="b101f-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b101f-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b101f-1199">作成</span><span class="sxs-lookup"><span data-stu-id="b101f-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b101f-1200">例</span><span class="sxs-lookup"><span data-stu-id="b101f-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```