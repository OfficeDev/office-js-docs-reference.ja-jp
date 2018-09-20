
# <a name="mailbox"></a><span data-ttu-id="2e34b-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="2e34b-101">mailbox</span></span>

### <span data-ttu-id="2e34b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="2e34b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="2e34b-104">Microsoft Outlook と web 上の Microsoft Outlook には、Outlook アドインのオブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e34b-105">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-105">Requirements</span></span>

|<span data-ttu-id="2e34b-106">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-106">Requirement</span></span>| <span data-ttu-id="2e34b-107">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-109">1.0</span></span>|
|[<span data-ttu-id="2e34b-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="2e34b-111">Restricted</span></span>|
|[<span data-ttu-id="2e34b-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2e34b-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-114">Members and methods</span></span>

| <span data-ttu-id="2e34b-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="2e34b-115">Member</span></span> | <span data-ttu-id="2e34b-116">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2e34b-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="2e34b-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="2e34b-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="2e34b-118">Member</span></span> |
| [<span data-ttu-id="2e34b-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="2e34b-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="2e34b-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="2e34b-120">Member</span></span> |
| [<span data-ttu-id="2e34b-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2e34b-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2e34b-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-122">Method</span></span> |
| [<span data-ttu-id="2e34b-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="2e34b-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="2e34b-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-124">Method</span></span> |
| [<span data-ttu-id="2e34b-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2e34b-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | <span data-ttu-id="2e34b-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-126">Method</span></span> |
| [<span data-ttu-id="2e34b-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="2e34b-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="2e34b-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-128">Method</span></span> |
| [<span data-ttu-id="2e34b-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="2e34b-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="2e34b-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-130">Method</span></span> |
| [<span data-ttu-id="2e34b-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2e34b-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="2e34b-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-132">Method</span></span> |
| [<span data-ttu-id="2e34b-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="2e34b-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="2e34b-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-134">Method</span></span> |
| [<span data-ttu-id="2e34b-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2e34b-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="2e34b-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-136">Method</span></span> |
| [<span data-ttu-id="2e34b-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="2e34b-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="2e34b-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-138">Method</span></span> |
| [<span data-ttu-id="2e34b-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2e34b-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="2e34b-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-140">Method</span></span> |
| [<span data-ttu-id="2e34b-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2e34b-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="2e34b-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-142">Method</span></span> |
| [<span data-ttu-id="2e34b-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2e34b-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="2e34b-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-144">Method</span></span> |
| [<span data-ttu-id="2e34b-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="2e34b-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="2e34b-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2e34b-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="2e34b-147">Namespaces</span></span>

<span data-ttu-id="2e34b-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="2e34b-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="2e34b-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="2e34b-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="2e34b-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="2e34b-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="2e34b-152">ewsUrl :String</span></span>

<span data-ttu-id="2e34b-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-155">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-p103">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2e34b-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="2e34b-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="2e34b-161">型:</span><span class="sxs-lookup"><span data-stu-id="2e34b-161">Type:</span></span>

*   <span data-ttu-id="2e34b-162">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e34b-163">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-163">Requirements</span></span>

|<span data-ttu-id="2e34b-164">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-164">Requirement</span></span>| <span data-ttu-id="2e34b-165">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-167">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-167">1.0</span></span>|
|[<span data-ttu-id="2e34b-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-169">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="2e34b-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="2e34b-172">restUrl :String</span></span>

<span data-ttu-id="2e34b-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="2e34b-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="2e34b-175">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="2e34b-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="2e34b-178">型:</span><span class="sxs-lookup"><span data-stu-id="2e34b-178">Type:</span></span>

*   <span data-ttu-id="2e34b-179">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2e34b-180">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-180">Requirements</span></span>

|<span data-ttu-id="2e34b-181">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-181">Requirement</span></span>| <span data-ttu-id="2e34b-182">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-183">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-184">1.5</span><span class="sxs-lookup"><span data-stu-id="2e34b-184">1.5</span></span> |
|[<span data-ttu-id="2e34b-185">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-186">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-188">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="2e34b-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="2e34b-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2e34b-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2e34b-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2e34b-191">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2e34b-192">現在、サポートされているイベントの種類は、`Office.EventType.ItemChanged`と`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="2e34b-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-193">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-193">Parameters:</span></span>

| <span data-ttu-id="2e34b-194">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-194">Name</span></span> | <span data-ttu-id="2e34b-195">型</span><span class="sxs-lookup"><span data-stu-id="2e34b-195">Type</span></span> | <span data-ttu-id="2e34b-196">属性</span><span class="sxs-lookup"><span data-stu-id="2e34b-196">Attributes</span></span> | <span data-ttu-id="2e34b-197">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2e34b-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2e34b-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2e34b-199">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="2e34b-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2e34b-200">Function</span><span class="sxs-lookup"><span data-stu-id="2e34b-200">Function</span></span> || <span data-ttu-id="2e34b-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2e34b-204">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-204">Object</span></span> | <span data-ttu-id="2e34b-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-205">&lt;optional&gt;</span></span> | <span data-ttu-id="2e34b-206">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2e34b-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2e34b-207">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-207">Object</span></span> | <span data-ttu-id="2e34b-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-208">&lt;optional&gt;</span></span> | <span data-ttu-id="2e34b-209">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2e34b-210">function</span><span class="sxs-lookup"><span data-stu-id="2e34b-210">function</span></span>| <span data-ttu-id="2e34b-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-211">&lt;optional&gt;</span></span>|<span data-ttu-id="2e34b-212">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-213">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-213">Requirements</span></span>

|<span data-ttu-id="2e34b-214">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-214">Requirement</span></span>| <span data-ttu-id="2e34b-215">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-216">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-217">1.5</span><span class="sxs-lookup"><span data-stu-id="2e34b-217">1.5</span></span> |
|[<span data-ttu-id="2e34b-218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-219">ReadItem</span></span> |
|[<span data-ttu-id="2e34b-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-221">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-222">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="2e34b-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2e34b-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2e34b-224">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-225">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-p107">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-228">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-228">Parameters:</span></span>

|<span data-ttu-id="2e34b-229">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-229">Name</span></span>| <span data-ttu-id="2e34b-230">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-230">Type</span></span>| <span data-ttu-id="2e34b-231">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2e34b-232">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-232">String</span></span>|<span data-ttu-id="2e34b-233">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="2e34b-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="2e34b-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2e34b-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="2e34b-235">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="2e34b-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-236">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-236">Requirements</span></span>

|<span data-ttu-id="2e34b-237">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-237">Requirement</span></span>| <span data-ttu-id="2e34b-238">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-240">1.3</span><span class="sxs-lookup"><span data-stu-id="2e34b-240">1.3</span></span>|
|[<span data-ttu-id="2e34b-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-242">制限あり</span><span class="sxs-lookup"><span data-stu-id="2e34b-242">Restricted</span></span>|
|[<span data-ttu-id="2e34b-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-244">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e34b-245">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2e34b-245">Returns:</span></span>

<span data-ttu-id="2e34b-246">型:String</span><span class="sxs-lookup"><span data-stu-id="2e34b-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2e34b-247">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="2e34b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="2e34b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="2e34b-249">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="2e34b-p108">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="2e34b-p109">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-255">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-255">Parameters:</span></span>

|<span data-ttu-id="2e34b-256">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-256">Name</span></span>| <span data-ttu-id="2e34b-257">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-257">Type</span></span>| <span data-ttu-id="2e34b-258">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="2e34b-259">日付</span><span class="sxs-lookup"><span data-stu-id="2e34b-259">Date</span></span>|<span data-ttu-id="2e34b-260">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2e34b-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-261">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-261">Requirements</span></span>

|<span data-ttu-id="2e34b-262">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-262">Requirement</span></span>| <span data-ttu-id="2e34b-263">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-265">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-265">1.0</span></span>|
|[<span data-ttu-id="2e34b-266">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-267">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-269">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e34b-270">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2e34b-270">Returns:</span></span>

<span data-ttu-id="2e34b-271">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="2e34b-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="2e34b-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2e34b-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2e34b-273">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-274">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-277">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-277">Parameters:</span></span>

|<span data-ttu-id="2e34b-278">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-278">Name</span></span>| <span data-ttu-id="2e34b-279">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-279">Type</span></span>| <span data-ttu-id="2e34b-280">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2e34b-281">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-281">String</span></span>|<span data-ttu-id="2e34b-282">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="2e34b-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="2e34b-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2e34b-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="2e34b-284">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="2e34b-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-285">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-285">Requirements</span></span>

|<span data-ttu-id="2e34b-286">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-286">Requirement</span></span>| <span data-ttu-id="2e34b-287">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-289">1.3</span><span class="sxs-lookup"><span data-stu-id="2e34b-289">1.3</span></span>|
|[<span data-ttu-id="2e34b-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-291">制限あり</span><span class="sxs-lookup"><span data-stu-id="2e34b-291">Restricted</span></span>|
|[<span data-ttu-id="2e34b-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-293">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e34b-294">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2e34b-294">Returns:</span></span>

<span data-ttu-id="2e34b-295">型:String</span><span class="sxs-lookup"><span data-stu-id="2e34b-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2e34b-296">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="2e34b-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="2e34b-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="2e34b-298">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="2e34b-299">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-300">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-300">Parameters:</span></span>

|<span data-ttu-id="2e34b-301">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-301">Name</span></span>| <span data-ttu-id="2e34b-302">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-302">Type</span></span>| <span data-ttu-id="2e34b-303">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="2e34b-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2e34b-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="2e34b-305">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="2e34b-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-306">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-306">Requirements</span></span>

|<span data-ttu-id="2e34b-307">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-307">Requirement</span></span>| <span data-ttu-id="2e34b-308">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-309">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-310">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-310">1.0</span></span>|
|[<span data-ttu-id="2e34b-311">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-312">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-314">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2e34b-315">戻り値:</span><span class="sxs-lookup"><span data-stu-id="2e34b-315">Returns:</span></span>

<span data-ttu-id="2e34b-316">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="2e34b-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="2e34b-317">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="2e34b-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2e34b-318">Date</span><span class="sxs-lookup"><span data-stu-id="2e34b-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="2e34b-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2e34b-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="2e34b-320">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-321">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-322">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2e34b-p111">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="2e34b-325">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="2e34b-326">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-327">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-327">Parameters:</span></span>

|<span data-ttu-id="2e34b-328">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-328">Name</span></span>| <span data-ttu-id="2e34b-329">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-329">Type</span></span>| <span data-ttu-id="2e34b-330">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2e34b-331">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-331">String</span></span>|<span data-ttu-id="2e34b-332">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="2e34b-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-333">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-333">Requirements</span></span>

|<span data-ttu-id="2e34b-334">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-334">Requirement</span></span>| <span data-ttu-id="2e34b-335">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-336">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-337">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-337">1.0</span></span>|
|[<span data-ttu-id="2e34b-338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-339">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-341">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-342">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="2e34b-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2e34b-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="2e34b-344">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-345">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-346">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2e34b-347">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="2e34b-348">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="2e34b-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-351">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-351">Parameters:</span></span>

|<span data-ttu-id="2e34b-352">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-352">Name</span></span>| <span data-ttu-id="2e34b-353">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-353">Type</span></span>| <span data-ttu-id="2e34b-354">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2e34b-355">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-355">String</span></span>|<span data-ttu-id="2e34b-356">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="2e34b-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-357">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-357">Requirements</span></span>

|<span data-ttu-id="2e34b-358">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-358">Requirement</span></span>| <span data-ttu-id="2e34b-359">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-361">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-361">1.0</span></span>|
|[<span data-ttu-id="2e34b-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-363">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-365">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-366">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="2e34b-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2e34b-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="2e34b-368">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-369">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2e34b-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2e34b-p114">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="2e34b-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="2e34b-377">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-378">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-379">すべてのパラメーターはオプションです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-379">All parameters are optional.</span></span>

|<span data-ttu-id="2e34b-380">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-380">Name</span></span>| <span data-ttu-id="2e34b-381">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-381">Type</span></span>| <span data-ttu-id="2e34b-382">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2e34b-383">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-383">Object</span></span> | <span data-ttu-id="2e34b-384">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="2e34b-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="2e34b-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2e34b-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="2e34b-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2e34b-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="2e34b-391">日付</span><span class="sxs-lookup"><span data-stu-id="2e34b-391">Date</span></span> | <span data-ttu-id="2e34b-392">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="2e34b-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="2e34b-393">Date</span><span class="sxs-lookup"><span data-stu-id="2e34b-393">Date</span></span> | <span data-ttu-id="2e34b-394">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="2e34b-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="2e34b-395">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-395">String</span></span> | <span data-ttu-id="2e34b-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="2e34b-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="2e34b-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2e34b-401">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-401">String</span></span> | <span data-ttu-id="2e34b-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="2e34b-404">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-404">String</span></span> | <span data-ttu-id="2e34b-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2e34b-407">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-407">Requirements</span></span>

|<span data-ttu-id="2e34b-408">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-408">Requirement</span></span>| <span data-ttu-id="2e34b-409">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-410">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-411">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-411">1.0</span></span>|
|[<span data-ttu-id="2e34b-412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-413">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-415">読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-416">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-416">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="2e34b-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2e34b-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="2e34b-418">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="2e34b-419">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるようにするフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="2e34b-420">パラメーターを指定すると、メッセージ フォーム フィールドにはパラメーターのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2e34b-421">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-422">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-423">すべてのパラメーターはオプションです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-423">All parameters are optional.</span></span>

|<span data-ttu-id="2e34b-424">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-424">Name</span></span>| <span data-ttu-id="2e34b-425">種類</span><span class="sxs-lookup"><span data-stu-id="2e34b-425">Type</span></span>| <span data-ttu-id="2e34b-426">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2e34b-427">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-427">Object</span></span> | <span data-ttu-id="2e34b-428">新しいメッセージを記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="2e34b-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="2e34b-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2e34b-430">メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="2e34b-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="2e34b-431">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="2e34b-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2e34b-433">メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="2e34b-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="2e34b-434">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="2e34b-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2e34b-436">メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="2e34b-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="2e34b-437">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2e34b-438">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-438">String</span></span> | <span data-ttu-id="2e34b-439">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="2e34b-439">A string containing the subject of the message.</span></span> <span data-ttu-id="2e34b-440">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="2e34b-441">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-441">String</span></span> | <span data-ttu-id="2e34b-442">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="2e34b-442">The HTML body of the message.</span></span> <span data-ttu-id="2e34b-443">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="2e34b-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2e34b-445">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="2e34b-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="2e34b-446">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-446">String</span></span> | <span data-ttu-id="2e34b-p128">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="2e34b-449">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-449">String</span></span> | <span data-ttu-id="2e34b-450">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="2e34b-451">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-451">String</span></span> | <span data-ttu-id="2e34b-p129">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="2e34b-454">Boolean</span><span class="sxs-lookup"><span data-stu-id="2e34b-454">Boolean</span></span> | <span data-ttu-id="2e34b-p130">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="2e34b-457">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-457">String</span></span> | <span data-ttu-id="2e34b-458">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="2e34b-459">新しいメッセージに添付する、既存の電子メールの EWS の項目の id です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="2e34b-460">最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="2e34b-461">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-461">Requirements</span></span>

|<span data-ttu-id="2e34b-462">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-462">Requirement</span></span>| <span data-ttu-id="2e34b-463">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-464">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-464">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-465">1.6</span><span class="sxs-lookup"><span data-stu-id="2e34b-465">1.6</span></span> |
|[<span data-ttu-id="2e34b-466">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-467">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-468">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-469">読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-470">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-470">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="2e34b-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2e34b-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="2e34b-472">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="2e34b-p132">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-475">アドインが可能な場合に、Exchange Web サービスではなく REST Api を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="2e34b-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="2e34b-476">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="2e34b-476">**REST Tokens**</span></span>

<span data-ttu-id="2e34b-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="2e34b-480">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="2e34b-481">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="2e34b-481">**EWS Tokens**</span></span>

<span data-ttu-id="2e34b-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="2e34b-484">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-485">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-485">Parameters:</span></span>

|<span data-ttu-id="2e34b-486">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-486">Name</span></span>| <span data-ttu-id="2e34b-487">型</span><span class="sxs-lookup"><span data-stu-id="2e34b-487">Type</span></span>| <span data-ttu-id="2e34b-488">属性</span><span class="sxs-lookup"><span data-stu-id="2e34b-488">Attributes</span></span>| <span data-ttu-id="2e34b-489">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="2e34b-490">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2e34b-490">Object</span></span> | <span data-ttu-id="2e34b-491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-491">&lt;optional&gt;</span></span> | <span data-ttu-id="2e34b-492">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="2e34b-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="2e34b-493">Boolean</span><span class="sxs-lookup"><span data-stu-id="2e34b-493">Boolean</span></span> |  <span data-ttu-id="2e34b-494">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-494">&lt;optional&gt;</span></span> | <span data-ttu-id="2e34b-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2e34b-497">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-497">Object</span></span> |  <span data-ttu-id="2e34b-498">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-498">&lt;optional&gt;</span></span> | <span data-ttu-id="2e34b-499">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="2e34b-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="2e34b-500">function</span><span class="sxs-lookup"><span data-stu-id="2e34b-500">function</span></span>||<span data-ttu-id="2e34b-p136">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-503">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-503">Requirements</span></span>

|<span data-ttu-id="2e34b-504">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-504">Requirement</span></span>| <span data-ttu-id="2e34b-505">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-507">1.5</span><span class="sxs-lookup"><span data-stu-id="2e34b-507">1.5</span></span> |
|[<span data-ttu-id="2e34b-508">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-509">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-511">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="2e34b-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-512">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-512">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="2e34b-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2e34b-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2e34b-514">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="2e34b-p137">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="2e34b-p138">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2e34b-520">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="2e34b-p139">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-523">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-523">Parameters:</span></span>

|<span data-ttu-id="2e34b-524">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-524">Name</span></span>| <span data-ttu-id="2e34b-525">型</span><span class="sxs-lookup"><span data-stu-id="2e34b-525">Type</span></span>| <span data-ttu-id="2e34b-526">属性</span><span class="sxs-lookup"><span data-stu-id="2e34b-526">Attributes</span></span>| <span data-ttu-id="2e34b-527">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2e34b-528">function</span><span class="sxs-lookup"><span data-stu-id="2e34b-528">function</span></span>||<span data-ttu-id="2e34b-p140">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2e34b-531">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2e34b-531">Object</span></span>| <span data-ttu-id="2e34b-532">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-532">&lt;optional&gt;</span></span>|<span data-ttu-id="2e34b-533">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-534">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-534">Requirements</span></span>

|<span data-ttu-id="2e34b-535">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-535">Requirement</span></span>| <span data-ttu-id="2e34b-536">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-537">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-537">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-538">1.3</span><span class="sxs-lookup"><span data-stu-id="2e34b-538">1.3</span></span>|
|[<span data-ttu-id="2e34b-539">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-540">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-541">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-542">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="2e34b-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-543">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="2e34b-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2e34b-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2e34b-545">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="2e34b-546">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-547">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-547">Parameters:</span></span>

|<span data-ttu-id="2e34b-548">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-548">Name</span></span>| <span data-ttu-id="2e34b-549">型</span><span class="sxs-lookup"><span data-stu-id="2e34b-549">Type</span></span>| <span data-ttu-id="2e34b-550">属性</span><span class="sxs-lookup"><span data-stu-id="2e34b-550">Attributes</span></span>| <span data-ttu-id="2e34b-551">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2e34b-552">function</span><span class="sxs-lookup"><span data-stu-id="2e34b-552">function</span></span>||<span data-ttu-id="2e34b-553">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2e34b-554">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2e34b-555">Object</span><span class="sxs-lookup"><span data-stu-id="2e34b-555">Object</span></span>| <span data-ttu-id="2e34b-556">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-556">&lt;optional&gt;</span></span>|<span data-ttu-id="2e34b-557">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-558">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-558">Requirements</span></span>

|<span data-ttu-id="2e34b-559">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-559">Requirement</span></span>| <span data-ttu-id="2e34b-560">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-561">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-562">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-562">1.0</span></span>|
|[<span data-ttu-id="2e34b-563">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2e34b-564">ReadItem</span></span>|
|[<span data-ttu-id="2e34b-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-566">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-567">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="2e34b-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2e34b-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="2e34b-569">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="2e34b-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-570">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="2e34b-571">IOS は、Outlook またはアプリは、Outlook で</span><span class="sxs-lookup"><span data-stu-id="2e34b-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="2e34b-572">アドインの読み込み時 Gmail のメールボックスに</span><span class="sxs-lookup"><span data-stu-id="2e34b-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="2e34b-573">アドインでは、これらの場合では、 [REST Api を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="2e34b-574">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="2e34b-575">サポートされている EWS 操作の一覧については、 [Outlook のアドインからの web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e34b-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="2e34b-576">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="2e34b-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="2e34b-577">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="2e34b-p142">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="2e34b-580">サーバーの管理者を設定する必要があります`OAuthAuthentication`場合は true を有効にするクライアント アクセス サーバーの EWS のディレクトリに、 `makeEwsRequestAsync` EWS を使用する方法を要求します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-580">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="2e34b-581">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="2e34b-581">Version differences</span></span>

<span data-ttu-id="2e34b-582">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2e34b-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="2e34b-p143">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2e34b-586">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="2e34b-586">Parameters:</span></span>

|<span data-ttu-id="2e34b-587">名前</span><span class="sxs-lookup"><span data-stu-id="2e34b-587">Name</span></span>| <span data-ttu-id="2e34b-588">型</span><span class="sxs-lookup"><span data-stu-id="2e34b-588">Type</span></span>| <span data-ttu-id="2e34b-589">属性</span><span class="sxs-lookup"><span data-stu-id="2e34b-589">Attributes</span></span>| <span data-ttu-id="2e34b-590">説明</span><span class="sxs-lookup"><span data-stu-id="2e34b-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2e34b-591">String</span><span class="sxs-lookup"><span data-stu-id="2e34b-591">String</span></span>||<span data-ttu-id="2e34b-592">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="2e34b-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="2e34b-593">function</span><span class="sxs-lookup"><span data-stu-id="2e34b-593">function</span></span>||<span data-ttu-id="2e34b-594">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2e34b-595">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="2e34b-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="2e34b-596">結果は、サイズの 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="2e34b-596">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="2e34b-597">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="2e34b-597">Object</span></span>| <span data-ttu-id="2e34b-598">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="2e34b-598">&lt;optional&gt;</span></span>|<span data-ttu-id="2e34b-599">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="2e34b-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2e34b-600">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-600">Requirements</span></span>

|<span data-ttu-id="2e34b-601">要件</span><span class="sxs-lookup"><span data-stu-id="2e34b-601">Requirement</span></span>| <span data-ttu-id="2e34b-602">値</span><span class="sxs-lookup"><span data-stu-id="2e34b-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="2e34b-603">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2e34b-603">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2e34b-604">1.0</span><span class="sxs-lookup"><span data-stu-id="2e34b-604">1.0</span></span>|
|[<span data-ttu-id="2e34b-605">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2e34b-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2e34b-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2e34b-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="2e34b-607">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2e34b-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2e34b-608">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="2e34b-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2e34b-609">例</span><span class="sxs-lookup"><span data-stu-id="2e34b-609">Example</span></span>

<span data-ttu-id="2e34b-610">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="2e34b-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```