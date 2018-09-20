
# <a name="item"></a><span data-ttu-id="6fc71-101">item</span><span class="sxs-lookup"><span data-stu-id="6fc71-101">item</span></span>

### <span data-ttu-id="6fc71-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="6fc71-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="6fc71-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-106">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-106">Requirements</span></span>

|<span data-ttu-id="6fc71-107">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-107">Requirement</span></span>| <span data-ttu-id="6fc71-108">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-110">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-110">1.0</span></span>|
|[<span data-ttu-id="6fc71-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="6fc71-112">Restricted</span></span>|
|[<span data-ttu-id="6fc71-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="6fc71-115">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-115">Example</span></span>

<span data-ttu-id="6fc71-116">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="6fc71-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="6fc71-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="6fc71-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6fc71-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="6fc71-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-121">ファイルの特定の種類は、潜在的なセキュリティの問題により、Outlook によってブロックされは返されません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6fc71-122">詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6fc71-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-123">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-123">Type:</span></span>

*   <span data-ttu-id="6fc71-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6fc71-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-125">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-125">Requirements</span></span>

|<span data-ttu-id="6fc71-126">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-126">Requirement</span></span>| <span data-ttu-id="6fc71-127">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-128">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-129">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-129">1.0</span></span>|
|[<span data-ttu-id="6fc71-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-131">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-134">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-134">Example</span></span>

<span data-ttu-id="6fc71-135">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6fc71-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6fc71-137">取得またはメッセージの bcc (ブラインド カーボン コピー) 受信者を更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6fc71-138">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-139">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-139">Type:</span></span>

*   [<span data-ttu-id="6fc71-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="6fc71-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="6fc71-141">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-141">Requirements</span></span>

|<span data-ttu-id="6fc71-142">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-142">Requirement</span></span>| <span data-ttu-id="6fc71-143">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-145">1.1</span><span class="sxs-lookup"><span data-stu-id="6fc71-145">1.1</span></span>|
|[<span data-ttu-id="6fc71-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-147">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-149">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-150">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="6fc71-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="6fc71-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="6fc71-152">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-153">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-153">Type:</span></span>

*   [<span data-ttu-id="6fc71-154">Body</span><span class="sxs-lookup"><span data-stu-id="6fc71-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="6fc71-155">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-155">Requirements</span></span>

|<span data-ttu-id="6fc71-156">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-156">Requirement</span></span>| <span data-ttu-id="6fc71-157">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-159">1.1</span><span class="sxs-lookup"><span data-stu-id="6fc71-159">1.1</span></span>|
|[<span data-ttu-id="6fc71-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-161">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6fc71-164">[cc]: 配列 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook_1_2/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="6fc71-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6fc71-165">メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6fc71-166">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-167">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-167">Read mode</span></span>

<span data-ttu-id="6fc71-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-170">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-170">Compose mode</span></span>

<span data-ttu-id="6fc71-171">`cc`を`Recipients`オブジェクトを取得または、メッセージの**Cc**行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-172">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-172">Type:</span></span>

*   <span data-ttu-id="6fc71-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-174">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-174">Requirements</span></span>

|<span data-ttu-id="6fc71-175">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-175">Requirement</span></span>| <span data-ttu-id="6fc71-176">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-177">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-178">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-178">1.0</span></span>|
|[<span data-ttu-id="6fc71-179">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-180">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-181">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-182">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-183">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="6fc71-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="6fc71-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="6fc71-185">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6fc71-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6fc71-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-190">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-190">Type:</span></span>

*   <span data-ttu-id="6fc71-191">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-192">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-192">Requirements</span></span>

|<span data-ttu-id="6fc71-193">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-193">Requirement</span></span>| <span data-ttu-id="6fc71-194">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-195">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-196">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-196">1.0</span></span>|
|[<span data-ttu-id="6fc71-197">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-198">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-200">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="6fc71-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="6fc71-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="6fc71-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-204">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-204">Type:</span></span>

*   <span data-ttu-id="6fc71-205">日付</span><span class="sxs-lookup"><span data-stu-id="6fc71-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-206">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-206">Requirements</span></span>

|<span data-ttu-id="6fc71-207">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-207">Requirement</span></span>| <span data-ttu-id="6fc71-208">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-210">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-210">1.0</span></span>|
|[<span data-ttu-id="6fc71-211">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-212">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-214">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-215">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="6fc71-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="6fc71-216">dateTimeModified :Date</span></span>

<span data-ttu-id="6fc71-p111">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-219">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-220">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-220">Type:</span></span>

*   <span data-ttu-id="6fc71-221">日付</span><span class="sxs-lookup"><span data-stu-id="6fc71-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-222">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-222">Requirements</span></span>

|<span data-ttu-id="6fc71-223">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-223">Requirement</span></span>| <span data-ttu-id="6fc71-224">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-226">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-226">1.0</span></span>|
|[<span data-ttu-id="6fc71-227">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-228">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-230">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-231">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="6fc71-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6fc71-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="6fc71-233">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6fc71-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-236">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-236">Read mode</span></span>

<span data-ttu-id="6fc71-237">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-238">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-238">Compose mode</span></span>

<span data-ttu-id="6fc71-239">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6fc71-240">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-241">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-241">Type:</span></span>

*   <span data-ttu-id="6fc71-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6fc71-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-243">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-243">Requirements</span></span>

|<span data-ttu-id="6fc71-244">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-244">Requirement</span></span>| <span data-ttu-id="6fc71-245">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-246">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-247">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-247">1.0</span></span>|
|[<span data-ttu-id="6fc71-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-249">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-251">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-252">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-252">Example</span></span>

<span data-ttu-id="6fc71-253">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6fc71-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6fc71-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6fc71-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="6fc71-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-259">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`from`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="6fc71-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-260">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-260">Type:</span></span>

*   [<span data-ttu-id="6fc71-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6fc71-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6fc71-262">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-262">Requirements</span></span>

|<span data-ttu-id="6fc71-263">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-263">Requirement</span></span>| <span data-ttu-id="6fc71-264">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-266">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-266">1.0</span></span>|
|[<span data-ttu-id="6fc71-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-268">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-270">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="6fc71-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="6fc71-271">internetMessageId :String</span></span>

<span data-ttu-id="6fc71-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-274">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-274">Type:</span></span>

*   <span data-ttu-id="6fc71-275">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-276">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-276">Requirements</span></span>

|<span data-ttu-id="6fc71-277">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-277">Requirement</span></span>| <span data-ttu-id="6fc71-278">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-280">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-280">1.0</span></span>|
|[<span data-ttu-id="6fc71-281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-282">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-284">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-285">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="6fc71-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="6fc71-286">itemClass :String</span></span>

<span data-ttu-id="6fc71-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6fc71-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="6fc71-291">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-291">Type</span></span> | <span data-ttu-id="6fc71-292">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-292">Description</span></span> | <span data-ttu-id="6fc71-293">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="6fc71-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="6fc71-294">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="6fc71-294">Appointment items</span></span> | <span data-ttu-id="6fc71-295">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="6fc71-296">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="6fc71-296">Message items</span></span> | <span data-ttu-id="6fc71-297">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="6fc71-298">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-299">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-299">Type:</span></span>

*   <span data-ttu-id="6fc71-300">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-301">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-301">Requirements</span></span>

|<span data-ttu-id="6fc71-302">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-302">Requirement</span></span>| <span data-ttu-id="6fc71-303">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-305">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-305">1.0</span></span>|
|[<span data-ttu-id="6fc71-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-307">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-310">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6fc71-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="6fc71-311">(nullable) itemId :String</span></span>

<span data-ttu-id="6fc71-p118">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-314">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6fc71-315">`itemId`プロパティは、Outlook のエントリ ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6fc71-316">前にこの値を使用して REST API の呼び出しを行う、それを変換する`Office.context.mailbox.convertToRestId`、1.3 を設定する要件から利用できるようであります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="6fc71-317">詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6fc71-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-318">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-318">Type:</span></span>

*   <span data-ttu-id="6fc71-319">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-320">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-320">Requirements</span></span>

|<span data-ttu-id="6fc71-321">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-321">Requirement</span></span>| <span data-ttu-id="6fc71-322">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-323">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-324">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-324">1.0</span></span>|
|[<span data-ttu-id="6fc71-325">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-326">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-327">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-328">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-329">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-329">Example</span></span>

<span data-ttu-id="6fc71-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="6fc71-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="6fc71-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="6fc71-333">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6fc71-334">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-335">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-335">Type:</span></span>

*   [<span data-ttu-id="6fc71-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6fc71-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="6fc71-337">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-337">Requirements</span></span>

|<span data-ttu-id="6fc71-338">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-338">Requirement</span></span>| <span data-ttu-id="6fc71-339">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-340">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-341">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-341">1.0</span></span>|
|[<span data-ttu-id="6fc71-342">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-343">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-344">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-345">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-346">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="6fc71-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="6fc71-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="6fc71-348">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-349">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-349">Read mode</span></span>

<span data-ttu-id="6fc71-350">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-351">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-351">Compose mode</span></span>

<span data-ttu-id="6fc71-352">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-353">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-353">Type:</span></span>

*   <span data-ttu-id="6fc71-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="6fc71-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-355">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-355">Requirements</span></span>

|<span data-ttu-id="6fc71-356">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-356">Requirement</span></span>| <span data-ttu-id="6fc71-357">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-358">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-359">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-359">1.0</span></span>|
|[<span data-ttu-id="6fc71-360">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-361">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-362">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-363">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-364">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6fc71-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="6fc71-365">normalizedSubject :String</span></span>

<span data-ttu-id="6fc71-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6fc71-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-370">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-370">Type:</span></span>

*   <span data-ttu-id="6fc71-371">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-372">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-372">Requirements</span></span>

|<span data-ttu-id="6fc71-373">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-373">Requirement</span></span>| <span data-ttu-id="6fc71-374">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-375">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-376">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-376">1.0</span></span>|
|[<span data-ttu-id="6fc71-377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-378">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-380">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-381">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6fc71-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6fc71-383">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6fc71-384">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-385">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-385">Read mode</span></span>

<span data-ttu-id="6fc71-386">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-387">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-387">Compose mode</span></span>

<span data-ttu-id="6fc71-388">`optionalAttendees`を`Recipients`オブジェクトを取得または省略可能な会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-389">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-389">Type:</span></span>

*   <span data-ttu-id="6fc71-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-391">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-391">Requirements</span></span>

|<span data-ttu-id="6fc71-392">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-392">Requirement</span></span>| <span data-ttu-id="6fc71-393">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-394">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-395">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-395">1.0</span></span>|
|[<span data-ttu-id="6fc71-396">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-397">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-398">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-399">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-400">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6fc71-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6fc71-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6fc71-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-404">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-404">Type:</span></span>

*   [<span data-ttu-id="6fc71-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6fc71-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6fc71-406">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-406">Requirements</span></span>

|<span data-ttu-id="6fc71-407">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-407">Requirement</span></span>| <span data-ttu-id="6fc71-408">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-410">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-410">1.0</span></span>|
|[<span data-ttu-id="6fc71-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-412">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-415">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6fc71-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6fc71-417">イベントの出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6fc71-418">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-419">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-419">Read mode</span></span>

<span data-ttu-id="6fc71-420">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-421">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-421">Compose mode</span></span>

<span data-ttu-id="6fc71-422">`requiredAttendees`を`Recipients`オブジェクトを取得または会議の出席者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-423">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-423">Type:</span></span>

*   <span data-ttu-id="6fc71-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-425">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-425">Requirements</span></span>

|<span data-ttu-id="6fc71-426">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-426">Requirement</span></span>| <span data-ttu-id="6fc71-427">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-429">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-429">1.0</span></span>|
|[<span data-ttu-id="6fc71-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-431">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-433">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-434">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6fc71-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6fc71-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6fc71-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6fc71-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-440">`recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`sender`プロパティは、 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="6fc71-440">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-441">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-441">Type:</span></span>

*   [<span data-ttu-id="6fc71-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6fc71-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6fc71-443">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-443">Requirements</span></span>

|<span data-ttu-id="6fc71-444">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-444">Requirement</span></span>| <span data-ttu-id="6fc71-445">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-447">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-447">1.0</span></span>|
|[<span data-ttu-id="6fc71-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-449">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-451">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-452">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="6fc71-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6fc71-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="6fc71-454">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6fc71-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-457">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-457">Read mode</span></span>

<span data-ttu-id="6fc71-458">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-459">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-459">Compose mode</span></span>

<span data-ttu-id="6fc71-460">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6fc71-461">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-462">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-462">Type:</span></span>

*   <span data-ttu-id="6fc71-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6fc71-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-464">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-464">Requirements</span></span>

|<span data-ttu-id="6fc71-465">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-465">Requirement</span></span>| <span data-ttu-id="6fc71-466">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-468">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-468">1.0</span></span>|
|[<span data-ttu-id="6fc71-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-470">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-472">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-473">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-473">Example</span></span>

<span data-ttu-id="6fc71-474">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="6fc71-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6fc71-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="6fc71-476">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6fc71-477">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-478">Read mode</span></span>

<span data-ttu-id="6fc71-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-481">Compose mode</span></span>

<span data-ttu-id="6fc71-482">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6fc71-483">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-483">Type:</span></span>

*   <span data-ttu-id="6fc71-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6fc71-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-485">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-485">Requirements</span></span>

|<span data-ttu-id="6fc71-486">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-486">Requirement</span></span>| <span data-ttu-id="6fc71-487">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-489">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-489">1.0</span></span>|
|[<span data-ttu-id="6fc71-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-491">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-493">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6fc71-494">: 配列 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook_1_2/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="6fc71-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6fc71-495">[メッセージの [**宛先**] 行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6fc71-496">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6fc71-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-497">Read mode</span></span>

<span data-ttu-id="6fc71-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6fc71-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="6fc71-500">Compose mode</span></span>

<span data-ttu-id="6fc71-501">`to`を`Recipients`オブジェクトを取得または、メッセージの [**宛先**] 行の受信者を更新するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6fc71-502">型:</span><span class="sxs-lookup"><span data-stu-id="6fc71-502">Type:</span></span>

*   <span data-ttu-id="6fc71-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6fc71-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-504">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-504">Requirements</span></span>

|<span data-ttu-id="6fc71-505">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-505">Requirement</span></span>| <span data-ttu-id="6fc71-506">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-508">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-508">1.0</span></span>|
|[<span data-ttu-id="6fc71-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-510">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-512">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-513">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="6fc71-514">メソッド</span><span class="sxs-lookup"><span data-stu-id="6fc71-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6fc71-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6fc71-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6fc71-516">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6fc71-517">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6fc71-518">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-519">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-519">Parameters:</span></span>

|<span data-ttu-id="6fc71-520">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-520">Name</span></span>| <span data-ttu-id="6fc71-521">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-521">Type</span></span>| <span data-ttu-id="6fc71-522">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-522">Attributes</span></span>| <span data-ttu-id="6fc71-523">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="6fc71-524">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-524">String</span></span>||<span data-ttu-id="6fc71-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6fc71-527">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-527">String</span></span>||<span data-ttu-id="6fc71-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6fc71-530">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-530">Object</span></span>| <span data-ttu-id="6fc71-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-531">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-532">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="6fc71-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6fc71-533">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-533">Object</span></span>| <span data-ttu-id="6fc71-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-534">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-535">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6fc71-536">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-536">function</span></span>| <span data-ttu-id="6fc71-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-537">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-538">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6fc71-539">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6fc71-540">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6fc71-541">エラー</span><span class="sxs-lookup"><span data-stu-id="6fc71-541">Errors</span></span>

| <span data-ttu-id="6fc71-542">エラー コード</span><span class="sxs-lookup"><span data-stu-id="6fc71-542">Error code</span></span> | <span data-ttu-id="6fc71-543">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="6fc71-544">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="6fc71-545">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6fc71-546">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-547">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-547">Requirements</span></span>

|<span data-ttu-id="6fc71-548">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-548">Requirement</span></span>| <span data-ttu-id="6fc71-549">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-550">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-551">1.1</span><span class="sxs-lookup"><span data-stu-id="6fc71-551">1.1</span></span>|
|[<span data-ttu-id="6fc71-552">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="6fc71-554">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-555">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-556">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6fc71-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6fc71-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6fc71-558">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6fc71-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6fc71-562">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6fc71-563">Office アドインは、Outlook Web App で実行されている場合、`addItemAttachmentAsync`メソッドが項目を編集しているアイテム以外のアイテムに関連付けることができますただし、これはサポートされていません、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-564">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-564">Parameters:</span></span>

|<span data-ttu-id="6fc71-565">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-565">Name</span></span>| <span data-ttu-id="6fc71-566">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-566">Type</span></span>| <span data-ttu-id="6fc71-567">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-567">Attributes</span></span>| <span data-ttu-id="6fc71-568">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="6fc71-569">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-569">String</span></span>||<span data-ttu-id="6fc71-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6fc71-572">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-572">String</span></span>||<span data-ttu-id="6fc71-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6fc71-575">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-575">Object</span></span>| <span data-ttu-id="6fc71-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-576">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-577">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="6fc71-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6fc71-578">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-578">Object</span></span>| <span data-ttu-id="6fc71-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-579">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-580">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6fc71-581">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-581">function</span></span>| <span data-ttu-id="6fc71-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-582">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-583">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6fc71-584">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6fc71-585">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6fc71-586">エラー</span><span class="sxs-lookup"><span data-stu-id="6fc71-586">Errors</span></span>

| <span data-ttu-id="6fc71-587">エラー コード</span><span class="sxs-lookup"><span data-stu-id="6fc71-587">Error code</span></span> | <span data-ttu-id="6fc71-588">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6fc71-589">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-590">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-590">Requirements</span></span>

|<span data-ttu-id="6fc71-591">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-591">Requirement</span></span>| <span data-ttu-id="6fc71-592">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-593">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-594">1.1</span><span class="sxs-lookup"><span data-stu-id="6fc71-594">1.1</span></span>|
|[<span data-ttu-id="6fc71-595">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="6fc71-597">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-598">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-599">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-599">Example</span></span>

<span data-ttu-id="6fc71-600">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="6fc71-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6fc71-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="6fc71-602">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-603">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6fc71-604">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6fc71-605">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="6fc71-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6fc71-p137">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-609">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-609">Parameters:</span></span>

|<span data-ttu-id="6fc71-610">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-610">Name</span></span>| <span data-ttu-id="6fc71-611">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-611">Type</span></span>| <span data-ttu-id="6fc71-612">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6fc71-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-613">String &#124; Object</span></span>| |<span data-ttu-id="6fc71-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6fc71-616">**または**</span><span class="sxs-lookup"><span data-stu-id="6fc71-616">**OR**</span></span><br/><span data-ttu-id="6fc71-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6fc71-619">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-619">String</span></span> | <span data-ttu-id="6fc71-620">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-620">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6fc71-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6fc71-624">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-624">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-625">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="6fc71-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6fc71-626">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-626">String</span></span> | | <span data-ttu-id="6fc71-p141">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6fc71-629">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-629">String</span></span> | | <span data-ttu-id="6fc71-630">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6fc71-631">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-631">String</span></span> | | <span data-ttu-id="6fc71-p142">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6fc71-634">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-634">String</span></span> | | <span data-ttu-id="6fc71-p143">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6fc71-638">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-638">function</span></span> | <span data-ttu-id="6fc71-639">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-639">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-640">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-641">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-641">Requirements</span></span>

|<span data-ttu-id="6fc71-642">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-642">Requirement</span></span>| <span data-ttu-id="6fc71-643">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-644">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-645">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-645">1.0</span></span>|
|[<span data-ttu-id="6fc71-646">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-647">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-648">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-649">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6fc71-650">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-650">Examples</span></span>

<span data-ttu-id="6fc71-651">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6fc71-652">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6fc71-653">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6fc71-654">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-654">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6fc71-655">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-655">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6fc71-656">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="6fc71-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6fc71-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="6fc71-658">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-659">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-659">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6fc71-660">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6fc71-661">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="6fc71-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6fc71-p144">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-665">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-665">Parameters:</span></span>

|<span data-ttu-id="6fc71-666">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-666">Name</span></span>| <span data-ttu-id="6fc71-667">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-667">Type</span></span>| <span data-ttu-id="6fc71-668">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6fc71-669">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-669">String &#124; Object</span></span>| | <span data-ttu-id="6fc71-p145">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6fc71-672">**または**</span><span class="sxs-lookup"><span data-stu-id="6fc71-672">**OR**</span></span><br/><span data-ttu-id="6fc71-p146">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6fc71-675">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-675">String</span></span> | <span data-ttu-id="6fc71-676">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-676">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6fc71-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6fc71-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-680">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-681">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="6fc71-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6fc71-682">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-682">String</span></span> | | <span data-ttu-id="6fc71-p148">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6fc71-685">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-685">String</span></span> | | <span data-ttu-id="6fc71-686">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6fc71-687">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-687">String</span></span> | | <span data-ttu-id="6fc71-p149">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6fc71-690">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-690">String</span></span> | | <span data-ttu-id="6fc71-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6fc71-694">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-694">function</span></span> | <span data-ttu-id="6fc71-695">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-695">&lt;optional&gt;</span></span> | <span data-ttu-id="6fc71-696">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-697">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-697">Requirements</span></span>

|<span data-ttu-id="6fc71-698">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-698">Requirement</span></span>| <span data-ttu-id="6fc71-699">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-700">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-701">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-701">1.0</span></span>|
|[<span data-ttu-id="6fc71-702">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-703">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-704">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-705">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6fc71-706">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-706">Examples</span></span>

<span data-ttu-id="6fc71-707">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6fc71-708">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6fc71-709">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6fc71-710">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-710">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6fc71-711">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-711">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6fc71-712">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="6fc71-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="6fc71-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="6fc71-714">選択したアイテムの本文内のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-714">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-715">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-716">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-716">Requirements</span></span>

|<span data-ttu-id="6fc71-717">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-717">Requirement</span></span>| <span data-ttu-id="6fc71-718">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-719">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-720">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-720">1.0</span></span>|
|[<span data-ttu-id="6fc71-721">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-722">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-723">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-724">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-725">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-725">Returns:</span></span>

<span data-ttu-id="6fc71-726">型:[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="6fc71-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="6fc71-727">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-727">Example</span></span>

<span data-ttu-id="6fc71-728">次の使用例は、現在の項目の本文に連絡先のエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-728">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="6fc71-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6fc71-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6fc71-730">選択したアイテムの本文に指定されたエンティティ型のすべてのエンティティの配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-730">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-731">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-731">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-732">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-732">Parameters:</span></span>

|<span data-ttu-id="6fc71-733">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-733">Name</span></span>| <span data-ttu-id="6fc71-734">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-734">Type</span></span>| <span data-ttu-id="6fc71-735">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="6fc71-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6fc71-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="6fc71-737">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="6fc71-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fc71-738">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fc71-738">Requirements</span></span>

|<span data-ttu-id="6fc71-739">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-739">Requirement</span></span>| <span data-ttu-id="6fc71-740">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-741">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-742">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-742">1.0</span></span>|
|[<span data-ttu-id="6fc71-743">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-744">制限あり</span><span class="sxs-lookup"><span data-stu-id="6fc71-744">Restricted</span></span>|
|[<span data-ttu-id="6fc71-745">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-746">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-747">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-747">Returns:</span></span>

<span data-ttu-id="6fc71-748">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6fc71-749">アイテムの本文に指定した型のエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-749">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6fc71-750">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6fc71-751">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="6fc71-752">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="6fc71-752">Value of `entityType`</span></span> | <span data-ttu-id="6fc71-753">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="6fc71-753">Type of objects in returned array</span></span> | <span data-ttu-id="6fc71-754">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="6fc71-755">文字列</span><span class="sxs-lookup"><span data-stu-id="6fc71-755">String</span></span> | <span data-ttu-id="6fc71-756">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="6fc71-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="6fc71-757">連絡先</span><span class="sxs-lookup"><span data-stu-id="6fc71-757">Contact</span></span> | <span data-ttu-id="6fc71-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6fc71-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="6fc71-759">文字列</span><span class="sxs-lookup"><span data-stu-id="6fc71-759">String</span></span> | <span data-ttu-id="6fc71-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6fc71-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="6fc71-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6fc71-761">MeetingSuggestion</span></span> | <span data-ttu-id="6fc71-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6fc71-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="6fc71-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6fc71-763">PhoneNumber</span></span> | <span data-ttu-id="6fc71-764">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="6fc71-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="6fc71-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6fc71-765">TaskSuggestion</span></span> | <span data-ttu-id="6fc71-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6fc71-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="6fc71-767">文字列</span><span class="sxs-lookup"><span data-stu-id="6fc71-767">String</span></span> | <span data-ttu-id="6fc71-768">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="6fc71-768">**Restricted**</span></span> |

<span data-ttu-id="6fc71-769">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6fc71-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="6fc71-770">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-770">Example</span></span>

<span data-ttu-id="6fc71-771">次の例では、現在の項目の本文に郵便番号のアドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-771">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="6fc71-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6fc71-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6fc71-773">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-774">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-774">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6fc71-775">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-776">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-776">Parameters:</span></span>

|<span data-ttu-id="6fc71-777">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-777">Name</span></span>| <span data-ttu-id="6fc71-778">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-778">Type</span></span>| <span data-ttu-id="6fc71-779">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6fc71-780">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-780">String</span></span>|<span data-ttu-id="6fc71-781">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="6fc71-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fc71-782">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-782">Requirements</span></span>

|<span data-ttu-id="6fc71-783">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-783">Requirement</span></span>| <span data-ttu-id="6fc71-784">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-785">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-786">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-786">1.0</span></span>|
|[<span data-ttu-id="6fc71-787">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-788">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-789">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-790">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-791">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-791">Returns:</span></span>

<span data-ttu-id="6fc71-p152">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6fc71-794">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6fc71-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="6fc71-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6fc71-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6fc71-796">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-797">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-797">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6fc71-p153">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6fc71-801">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="6fc71-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6fc71-802">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="6fc71-p154">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fc71-805">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-805">Requirements</span></span>

|<span data-ttu-id="6fc71-806">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-806">Requirement</span></span>| <span data-ttu-id="6fc71-807">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-808">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-809">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-809">1.0</span></span>|
|[<span data-ttu-id="6fc71-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-811">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-813">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-814">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-814">Returns:</span></span>

<span data-ttu-id="6fc71-p155">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="6fc71-817">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="6fc71-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6fc71-818">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6fc71-819">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-819">Example</span></span>

<span data-ttu-id="6fc71-820">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="6fc71-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6fc71-821">getRegExMatchesByName(name)] → [(許容) {配列。 < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="6fc71-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6fc71-822">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6fc71-823">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-823">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6fc71-824">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6fc71-p156">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-827">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-827">Parameters:</span></span>

|<span data-ttu-id="6fc71-828">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-828">Name</span></span>| <span data-ttu-id="6fc71-829">種類</span><span class="sxs-lookup"><span data-stu-id="6fc71-829">Type</span></span>| <span data-ttu-id="6fc71-830">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6fc71-831">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-831">String</span></span>|<span data-ttu-id="6fc71-832">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="6fc71-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fc71-833">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-833">Requirements</span></span>

|<span data-ttu-id="6fc71-834">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-834">Requirement</span></span>| <span data-ttu-id="6fc71-835">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-836">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-837">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-837">1.0</span></span>|
|[<span data-ttu-id="6fc71-838">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-839">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-840">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-841">読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-842">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-842">Returns:</span></span>

<span data-ttu-id="6fc71-843">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="6fc71-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="6fc71-844">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="6fc71-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6fc71-845">配列。 < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="6fc71-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6fc71-846">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6fc71-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6fc71-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6fc71-848">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6fc71-p157">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-851">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-851">Parameters:</span></span>

|<span data-ttu-id="6fc71-852">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-852">Name</span></span>| <span data-ttu-id="6fc71-853">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-853">Type</span></span>| <span data-ttu-id="6fc71-854">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-854">Attributes</span></span>| <span data-ttu-id="6fc71-855">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="6fc71-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6fc71-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6fc71-p158">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="6fc71-860">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-860">Object</span></span>| <span data-ttu-id="6fc71-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-861">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-862">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="6fc71-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6fc71-863">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-863">Object</span></span>| <span data-ttu-id="6fc71-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-864">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-865">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6fc71-866">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-866">function</span></span>||<span data-ttu-id="6fc71-867">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6fc71-868">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6fc71-869">選択範囲は、source プロパティにアクセスするには、呼び出す`asyncResult.value.sourceProperty`、いずれかの方法となる`body`または`subject`。</span><span class="sxs-lookup"><span data-stu-id="6fc71-869">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fc71-870">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-870">Requirements</span></span>

|<span data-ttu-id="6fc71-871">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-871">Requirement</span></span>| <span data-ttu-id="6fc71-872">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-873">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-874">1.2</span><span class="sxs-lookup"><span data-stu-id="6fc71-874">1.2</span></span>|
|[<span data-ttu-id="6fc71-875">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="6fc71-877">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-878">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6fc71-879">戻り値:</span><span class="sxs-lookup"><span data-stu-id="6fc71-879">Returns:</span></span>

<span data-ttu-id="6fc71-880">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="6fc71-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="6fc71-881">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="6fc71-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6fc71-882">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6fc71-883">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-883">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6fc71-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6fc71-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6fc71-885">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6fc71-p160">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-889">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-889">Parameters:</span></span>

|<span data-ttu-id="6fc71-890">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-890">Name</span></span>| <span data-ttu-id="6fc71-891">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-891">Type</span></span>| <span data-ttu-id="6fc71-892">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-892">Attributes</span></span>| <span data-ttu-id="6fc71-893">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6fc71-894">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-894">function</span></span>||<span data-ttu-id="6fc71-895">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6fc71-896">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6fc71-897">取得し、アイテムのカスタム プロパティを削除してサーバーにバックアップを設定するカスタム プロパティに対する変更を保存するのには、このオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-897">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="6fc71-898">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="6fc71-898">Object</span></span>| <span data-ttu-id="6fc71-899">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-899">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-900">開発者は、コールバック関数にアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-900">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6fc71-901">によってこのオブジェクトにアクセスできる、`asyncResult.asyncContext`コールバック関数のプロパティです。</span><span class="sxs-lookup"><span data-stu-id="6fc71-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fc71-902">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-902">Requirements</span></span>

|<span data-ttu-id="6fc71-903">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-903">Requirement</span></span>| <span data-ttu-id="6fc71-904">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-906">1.0</span><span class="sxs-lookup"><span data-stu-id="6fc71-906">1.0</span></span>|
|[<span data-ttu-id="6fc71-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-908">ReadItem</span></span>|
|[<span data-ttu-id="6fc71-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-910">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="6fc71-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-911">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-911">Example</span></span>

<span data-ttu-id="6fc71-p163">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6fc71-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6fc71-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6fc71-916">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6fc71-p164">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-921">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-921">Parameters:</span></span>

|<span data-ttu-id="6fc71-922">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-922">Name</span></span>| <span data-ttu-id="6fc71-923">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-923">Type</span></span>| <span data-ttu-id="6fc71-924">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-924">Attributes</span></span>| <span data-ttu-id="6fc71-925">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="6fc71-926">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-926">String</span></span>||<span data-ttu-id="6fc71-p165">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="6fc71-929">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-929">Object</span></span>| <span data-ttu-id="6fc71-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-930">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-931">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="6fc71-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6fc71-932">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-932">Object</span></span>| <span data-ttu-id="6fc71-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-933">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-934">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6fc71-935">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-935">function</span></span>| <span data-ttu-id="6fc71-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-936">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-937">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6fc71-938">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6fc71-939">エラー</span><span class="sxs-lookup"><span data-stu-id="6fc71-939">Errors</span></span>

| <span data-ttu-id="6fc71-940">エラー コード</span><span class="sxs-lookup"><span data-stu-id="6fc71-940">Error code</span></span> | <span data-ttu-id="6fc71-941">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="6fc71-942">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="6fc71-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-943">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-943">Requirements</span></span>

|<span data-ttu-id="6fc71-944">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-944">Requirement</span></span>| <span data-ttu-id="6fc71-945">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-946">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-947">1.1</span><span class="sxs-lookup"><span data-stu-id="6fc71-947">1.1</span></span>|
|[<span data-ttu-id="6fc71-948">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="6fc71-950">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-951">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-952">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-952">Example</span></span>

<span data-ttu-id="6fc71-953">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6fc71-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6fc71-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6fc71-955">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="6fc71-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6fc71-p166">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6fc71-959">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="6fc71-959">Parameters:</span></span>

|<span data-ttu-id="6fc71-960">名前</span><span class="sxs-lookup"><span data-stu-id="6fc71-960">Name</span></span>| <span data-ttu-id="6fc71-961">型</span><span class="sxs-lookup"><span data-stu-id="6fc71-961">Type</span></span>| <span data-ttu-id="6fc71-962">属性</span><span class="sxs-lookup"><span data-stu-id="6fc71-962">Attributes</span></span>| <span data-ttu-id="6fc71-963">説明</span><span class="sxs-lookup"><span data-stu-id="6fc71-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6fc71-964">String</span><span class="sxs-lookup"><span data-stu-id="6fc71-964">String</span></span>||<span data-ttu-id="6fc71-p167">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="6fc71-968">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-968">Object</span></span>| <span data-ttu-id="6fc71-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-969">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-970">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="6fc71-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6fc71-971">Object</span><span class="sxs-lookup"><span data-stu-id="6fc71-971">Object</span></span>| <span data-ttu-id="6fc71-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-972">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-973">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="6fc71-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6fc71-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="6fc71-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6fc71-975">&lt;optional&gt;</span></span>|<span data-ttu-id="6fc71-p168">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6fc71-p169">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6fc71-980">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="6fc71-981">function</span><span class="sxs-lookup"><span data-stu-id="6fc71-981">function</span></span>||<span data-ttu-id="6fc71-982">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="6fc71-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6fc71-983">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-983">Requirements</span></span>

|<span data-ttu-id="6fc71-984">要件</span><span class="sxs-lookup"><span data-stu-id="6fc71-984">Requirement</span></span>| <span data-ttu-id="6fc71-985">値</span><span class="sxs-lookup"><span data-stu-id="6fc71-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fc71-986">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6fc71-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6fc71-987">1.2</span><span class="sxs-lookup"><span data-stu-id="6fc71-987">1.2</span></span>|
|[<span data-ttu-id="6fc71-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6fc71-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6fc71-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6fc71-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="6fc71-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6fc71-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6fc71-991">作成</span><span class="sxs-lookup"><span data-stu-id="6fc71-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6fc71-992">例</span><span class="sxs-lookup"><span data-stu-id="6fc71-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```