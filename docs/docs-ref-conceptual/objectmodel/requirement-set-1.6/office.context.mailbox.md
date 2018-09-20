
# <a name="mailbox"></a><span data-ttu-id="728c3-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="728c3-101">mailbox</span></span>

### <span data-ttu-id="728c3-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="728c3-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="728c3-104">Microsoft Outlook と web 上の Microsoft Outlook には、Outlook アドインのオブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="728c3-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="728c3-105">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-105">Requirements</span></span>

|<span data-ttu-id="728c3-106">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-106">Requirement</span></span>| <span data-ttu-id="728c3-107">値</span><span class="sxs-lookup"><span data-stu-id="728c3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-109">1.0</span></span>|
|[<span data-ttu-id="728c3-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="728c3-111">Restricted</span></span>|
|[<span data-ttu-id="728c3-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="728c3-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-114">Members and methods</span></span>

| <span data-ttu-id="728c3-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="728c3-115">Member</span></span> | <span data-ttu-id="728c3-116">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="728c3-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="728c3-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="728c3-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="728c3-118">Member</span></span> |
| [<span data-ttu-id="728c3-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="728c3-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="728c3-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="728c3-120">Member</span></span> |
| [<span data-ttu-id="728c3-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="728c3-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="728c3-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-122">Method</span></span> |
| [<span data-ttu-id="728c3-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="728c3-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="728c3-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-124">Method</span></span> |
| [<span data-ttu-id="728c3-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="728c3-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="728c3-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-126">Method</span></span> |
| [<span data-ttu-id="728c3-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="728c3-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="728c3-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-128">Method</span></span> |
| [<span data-ttu-id="728c3-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="728c3-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="728c3-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-130">Method</span></span> |
| [<span data-ttu-id="728c3-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="728c3-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="728c3-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-132">Method</span></span> |
| [<span data-ttu-id="728c3-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="728c3-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="728c3-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-134">Method</span></span> |
| [<span data-ttu-id="728c3-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="728c3-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="728c3-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-136">Method</span></span> |
| [<span data-ttu-id="728c3-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="728c3-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="728c3-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-138">Method</span></span> |
| [<span data-ttu-id="728c3-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="728c3-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="728c3-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-140">Method</span></span> |
| [<span data-ttu-id="728c3-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="728c3-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="728c3-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-142">Method</span></span> |
| [<span data-ttu-id="728c3-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="728c3-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="728c3-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-144">Method</span></span> |
| [<span data-ttu-id="728c3-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="728c3-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="728c3-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="728c3-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="728c3-147">Namespaces</span></span>

<span data-ttu-id="728c3-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="728c3-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="728c3-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="728c3-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="728c3-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="728c3-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="728c3-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="728c3-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="728c3-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="728c3-152">ewsUrl :String</span></span>

<span data-ttu-id="728c3-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="728c3-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-155">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-p103">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="728c3-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="728c3-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="728c3-161">型:</span><span class="sxs-lookup"><span data-stu-id="728c3-161">Type:</span></span>

*   <span data-ttu-id="728c3-162">String</span><span class="sxs-lookup"><span data-stu-id="728c3-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="728c3-163">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-163">Requirements</span></span>

|<span data-ttu-id="728c3-164">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-164">Requirement</span></span>| <span data-ttu-id="728c3-165">値</span><span class="sxs-lookup"><span data-stu-id="728c3-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-167">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-167">1.0</span></span>|
|[<span data-ttu-id="728c3-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-169">ReadItem</span></span>|
|[<span data-ttu-id="728c3-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="728c3-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="728c3-172">restUrl :String</span></span>

<span data-ttu-id="728c3-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="728c3-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="728c3-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="728c3-175">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="728c3-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="728c3-178">型:</span><span class="sxs-lookup"><span data-stu-id="728c3-178">Type:</span></span>

*   <span data-ttu-id="728c3-179">String</span><span class="sxs-lookup"><span data-stu-id="728c3-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="728c3-180">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-180">Requirements</span></span>

|<span data-ttu-id="728c3-181">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-181">Requirement</span></span>| <span data-ttu-id="728c3-182">値</span><span class="sxs-lookup"><span data-stu-id="728c3-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-183">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-184">1.5</span><span class="sxs-lookup"><span data-stu-id="728c3-184">1.5</span></span> |
|[<span data-ttu-id="728c3-185">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-186">ReadItem</span></span>|
|[<span data-ttu-id="728c3-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-188">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="728c3-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="728c3-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="728c3-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="728c3-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="728c3-191">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="728c3-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="728c3-p106">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="728c3-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-194">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-194">Parameters:</span></span>

| <span data-ttu-id="728c3-195">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-195">Name</span></span> | <span data-ttu-id="728c3-196">型</span><span class="sxs-lookup"><span data-stu-id="728c3-196">Type</span></span> | <span data-ttu-id="728c3-197">属性</span><span class="sxs-lookup"><span data-stu-id="728c3-197">Attributes</span></span> | <span data-ttu-id="728c3-198">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="728c3-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="728c3-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="728c3-200">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="728c3-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="728c3-201">Function</span><span class="sxs-lookup"><span data-stu-id="728c3-201">Function</span></span> || <span data-ttu-id="728c3-p107">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="728c3-205">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-205">Object</span></span> | <span data-ttu-id="728c3-206">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-206">&lt;optional&gt;</span></span> | <span data-ttu-id="728c3-207">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="728c3-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="728c3-208">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-208">Object</span></span> | <span data-ttu-id="728c3-209">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-209">&lt;optional&gt;</span></span> | <span data-ttu-id="728c3-210">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="728c3-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="728c3-211">function</span><span class="sxs-lookup"><span data-stu-id="728c3-211">function</span></span>| <span data-ttu-id="728c3-212">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-212">&lt;optional&gt;</span></span>|<span data-ttu-id="728c3-213">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-214">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-214">Requirements</span></span>

|<span data-ttu-id="728c3-215">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-215">Requirement</span></span>| <span data-ttu-id="728c3-216">値</span><span class="sxs-lookup"><span data-stu-id="728c3-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-218">1.5</span><span class="sxs-lookup"><span data-stu-id="728c3-218">1.5</span></span> |
|[<span data-ttu-id="728c3-219">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-220">ReadItem</span></span> |
|[<span data-ttu-id="728c3-221">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-222">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-223">例</span><span class="sxs-lookup"><span data-stu-id="728c3-223">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="728c3-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="728c3-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="728c3-225">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="728c3-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-226">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-226">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-229">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-229">Parameters:</span></span>

|<span data-ttu-id="728c3-230">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-230">Name</span></span>| <span data-ttu-id="728c3-231">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-231">Type</span></span>| <span data-ttu-id="728c3-232">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="728c3-233">String</span><span class="sxs-lookup"><span data-stu-id="728c3-233">String</span></span>|<span data-ttu-id="728c3-234">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="728c3-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="728c3-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="728c3-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="728c3-236">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="728c3-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-237">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-237">Requirements</span></span>

|<span data-ttu-id="728c3-238">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-238">Requirement</span></span>| <span data-ttu-id="728c3-239">値</span><span class="sxs-lookup"><span data-stu-id="728c3-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-241">1.3</span><span class="sxs-lookup"><span data-stu-id="728c3-241">1.3</span></span>|
|[<span data-ttu-id="728c3-242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-243">制限あり</span><span class="sxs-lookup"><span data-stu-id="728c3-243">Restricted</span></span>|
|[<span data-ttu-id="728c3-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-245">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="728c3-246">戻り値:</span><span class="sxs-lookup"><span data-stu-id="728c3-246">Returns:</span></span>

<span data-ttu-id="728c3-247">型:String</span><span class="sxs-lookup"><span data-stu-id="728c3-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="728c3-248">例</span><span class="sxs-lookup"><span data-stu-id="728c3-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="728c3-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="728c3-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="728c3-250">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="728c3-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="728c3-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-256">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-256">Parameters:</span></span>

|<span data-ttu-id="728c3-257">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-257">Name</span></span>| <span data-ttu-id="728c3-258">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-258">Type</span></span>| <span data-ttu-id="728c3-259">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="728c3-260">日付</span><span class="sxs-lookup"><span data-stu-id="728c3-260">Date</span></span>|<span data-ttu-id="728c3-261">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="728c3-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-262">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-262">Requirements</span></span>

|<span data-ttu-id="728c3-263">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-263">Requirement</span></span>| <span data-ttu-id="728c3-264">値</span><span class="sxs-lookup"><span data-stu-id="728c3-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-266">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-266">1.0</span></span>|
|[<span data-ttu-id="728c3-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-268">ReadItem</span></span>|
|[<span data-ttu-id="728c3-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-270">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="728c3-271">戻り値:</span><span class="sxs-lookup"><span data-stu-id="728c3-271">Returns:</span></span>

<span data-ttu-id="728c3-272">型:[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="728c3-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="728c3-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="728c3-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="728c3-274">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="728c3-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-275">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-275">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-p111">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-278">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-278">Parameters:</span></span>

|<span data-ttu-id="728c3-279">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-279">Name</span></span>| <span data-ttu-id="728c3-280">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-280">Type</span></span>| <span data-ttu-id="728c3-281">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="728c3-282">String</span><span class="sxs-lookup"><span data-stu-id="728c3-282">String</span></span>|<span data-ttu-id="728c3-283">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="728c3-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="728c3-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="728c3-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="728c3-285">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="728c3-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-286">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-286">Requirements</span></span>

|<span data-ttu-id="728c3-287">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-287">Requirement</span></span>| <span data-ttu-id="728c3-288">値</span><span class="sxs-lookup"><span data-stu-id="728c3-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-289">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-290">1.3</span><span class="sxs-lookup"><span data-stu-id="728c3-290">1.3</span></span>|
|[<span data-ttu-id="728c3-291">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-292">制限あり</span><span class="sxs-lookup"><span data-stu-id="728c3-292">Restricted</span></span>|
|[<span data-ttu-id="728c3-293">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-294">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="728c3-295">戻り値:</span><span class="sxs-lookup"><span data-stu-id="728c3-295">Returns:</span></span>

<span data-ttu-id="728c3-296">型:String</span><span class="sxs-lookup"><span data-stu-id="728c3-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="728c3-297">例</span><span class="sxs-lookup"><span data-stu-id="728c3-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="728c3-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="728c3-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="728c3-299">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="728c3-300">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="728c3-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-301">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-301">Parameters:</span></span>

|<span data-ttu-id="728c3-302">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-302">Name</span></span>| <span data-ttu-id="728c3-303">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-303">Type</span></span>| <span data-ttu-id="728c3-304">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="728c3-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="728c3-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="728c3-306">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="728c3-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-307">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-307">Requirements</span></span>

|<span data-ttu-id="728c3-308">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-308">Requirement</span></span>| <span data-ttu-id="728c3-309">値</span><span class="sxs-lookup"><span data-stu-id="728c3-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-310">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-311">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-311">1.0</span></span>|
|[<span data-ttu-id="728c3-312">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-313">ReadItem</span></span>|
|[<span data-ttu-id="728c3-314">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-315">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="728c3-316">戻り値:</span><span class="sxs-lookup"><span data-stu-id="728c3-316">Returns:</span></span>

<span data-ttu-id="728c3-317">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="728c3-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="728c3-318">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="728c3-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="728c3-319">Date</span><span class="sxs-lookup"><span data-stu-id="728c3-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="728c3-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="728c3-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="728c3-321">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="728c3-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-322">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-322">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-323">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="728c3-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="728c3-p112">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="728c3-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="728c3-326">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="728c3-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="728c3-327">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="728c3-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-328">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-328">Parameters:</span></span>

|<span data-ttu-id="728c3-329">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-329">Name</span></span>| <span data-ttu-id="728c3-330">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-330">Type</span></span>| <span data-ttu-id="728c3-331">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="728c3-332">String</span><span class="sxs-lookup"><span data-stu-id="728c3-332">String</span></span>|<span data-ttu-id="728c3-333">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="728c3-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-334">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-334">Requirements</span></span>

|<span data-ttu-id="728c3-335">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-335">Requirement</span></span>| <span data-ttu-id="728c3-336">値</span><span class="sxs-lookup"><span data-stu-id="728c3-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-337">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-338">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-338">1.0</span></span>|
|[<span data-ttu-id="728c3-339">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-340">ReadItem</span></span>|
|[<span data-ttu-id="728c3-341">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-342">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-343">例</span><span class="sxs-lookup"><span data-stu-id="728c3-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="728c3-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="728c3-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="728c3-345">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="728c3-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-346">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-346">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-347">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="728c3-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="728c3-348">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="728c3-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="728c3-349">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="728c3-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="728c3-p113">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-352">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-352">Parameters:</span></span>

|<span data-ttu-id="728c3-353">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-353">Name</span></span>| <span data-ttu-id="728c3-354">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-354">Type</span></span>| <span data-ttu-id="728c3-355">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="728c3-356">String</span><span class="sxs-lookup"><span data-stu-id="728c3-356">String</span></span>|<span data-ttu-id="728c3-357">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="728c3-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-358">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-358">Requirements</span></span>

|<span data-ttu-id="728c3-359">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-359">Requirement</span></span>| <span data-ttu-id="728c3-360">値</span><span class="sxs-lookup"><span data-stu-id="728c3-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-362">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-362">1.0</span></span>|
|[<span data-ttu-id="728c3-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-364">ReadItem</span></span>|
|[<span data-ttu-id="728c3-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-366">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-367">例</span><span class="sxs-lookup"><span data-stu-id="728c3-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="728c3-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="728c3-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="728c3-369">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="728c3-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-370">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-370">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="728c3-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="728c3-p115">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="728c3-p116">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="728c3-378">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="728c3-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-379">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-380">すべてのパラメーターはオプションです。</span><span class="sxs-lookup"><span data-stu-id="728c3-380">All parameters are optional.</span></span>

|<span data-ttu-id="728c3-381">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-381">Name</span></span>| <span data-ttu-id="728c3-382">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-382">Type</span></span>| <span data-ttu-id="728c3-383">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="728c3-384">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-384">Object</span></span> | <span data-ttu-id="728c3-385">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="728c3-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="728c3-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="728c3-p117">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="728c3-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="728c3-p118">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="728c3-392">日付</span><span class="sxs-lookup"><span data-stu-id="728c3-392">Date</span></span> | <span data-ttu-id="728c3-393">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="728c3-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="728c3-394">Date</span><span class="sxs-lookup"><span data-stu-id="728c3-394">Date</span></span> | <span data-ttu-id="728c3-395">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="728c3-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="728c3-396">String</span><span class="sxs-lookup"><span data-stu-id="728c3-396">String</span></span> | <span data-ttu-id="728c3-p119">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="728c3-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="728c3-p120">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="728c3-402">String</span><span class="sxs-lookup"><span data-stu-id="728c3-402">String</span></span> | <span data-ttu-id="728c3-p121">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="728c3-405">String</span><span class="sxs-lookup"><span data-stu-id="728c3-405">String</span></span> | <span data-ttu-id="728c3-p122">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="728c3-408">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-408">Requirements</span></span>

|<span data-ttu-id="728c3-409">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-409">Requirement</span></span>| <span data-ttu-id="728c3-410">値</span><span class="sxs-lookup"><span data-stu-id="728c3-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-412">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-412">1.0</span></span>|
|[<span data-ttu-id="728c3-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-414">ReadItem</span></span>|
|[<span data-ttu-id="728c3-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-417">例</span><span class="sxs-lookup"><span data-stu-id="728c3-417">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="728c3-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="728c3-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="728c3-419">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="728c3-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="728c3-420">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるようにするフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="728c3-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="728c3-421">パラメーターを指定すると、メッセージ フォーム フィールドにはパラメーターのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="728c3-422">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="728c3-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-423">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-424">すべてのパラメーターはオプションです。</span><span class="sxs-lookup"><span data-stu-id="728c3-424">All parameters are optional.</span></span>

|<span data-ttu-id="728c3-425">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-425">Name</span></span>| <span data-ttu-id="728c3-426">種類</span><span class="sxs-lookup"><span data-stu-id="728c3-426">Type</span></span>| <span data-ttu-id="728c3-427">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="728c3-428">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-428">Object</span></span> | <span data-ttu-id="728c3-429">新しいメッセージを記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="728c3-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="728c3-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="728c3-431">メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="728c3-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="728c3-432">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="728c3-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="728c3-434">メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="728c3-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="728c3-435">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="728c3-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="728c3-437">メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="728c3-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="728c3-438">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="728c3-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="728c3-439">String</span><span class="sxs-lookup"><span data-stu-id="728c3-439">String</span></span> | <span data-ttu-id="728c3-440">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="728c3-440">A string containing the subject of the message.</span></span> <span data-ttu-id="728c3-441">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="728c3-442">String</span><span class="sxs-lookup"><span data-stu-id="728c3-442">String</span></span> | <span data-ttu-id="728c3-443">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="728c3-443">The HTML body of the message.</span></span> <span data-ttu-id="728c3-444">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="728c3-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="728c3-446">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="728c3-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="728c3-447">String</span><span class="sxs-lookup"><span data-stu-id="728c3-447">String</span></span> | <span data-ttu-id="728c3-p129">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="728c3-450">String</span><span class="sxs-lookup"><span data-stu-id="728c3-450">String</span></span> | <span data-ttu-id="728c3-451">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="728c3-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="728c3-452">String</span><span class="sxs-lookup"><span data-stu-id="728c3-452">String</span></span> | <span data-ttu-id="728c3-p130">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="728c3-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="728c3-455">Boolean</span><span class="sxs-lookup"><span data-stu-id="728c3-455">Boolean</span></span> | <span data-ttu-id="728c3-p131">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="728c3-458">String</span><span class="sxs-lookup"><span data-stu-id="728c3-458">String</span></span> | <span data-ttu-id="728c3-459">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="728c3-460">新しいメッセージに添付する、既存の電子メールの EWS の項目の id です。</span><span class="sxs-lookup"><span data-stu-id="728c3-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="728c3-461">最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="728c3-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="728c3-462">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-462">Requirements</span></span>

|<span data-ttu-id="728c3-463">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-463">Requirement</span></span>| <span data-ttu-id="728c3-464">値</span><span class="sxs-lookup"><span data-stu-id="728c3-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-465">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-466">1.6</span><span class="sxs-lookup"><span data-stu-id="728c3-466">1.6</span></span> |
|[<span data-ttu-id="728c3-467">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-468">ReadItem</span></span>|
|[<span data-ttu-id="728c3-469">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-470">読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-471">例</span><span class="sxs-lookup"><span data-stu-id="728c3-471">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="728c3-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="728c3-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="728c3-473">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="728c3-p133">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-476">アドインが可能な場合に、Exchange Web サービスではなく REST Api を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="728c3-476">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="728c3-477">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="728c3-477">**REST Tokens**</span></span>

<span data-ttu-id="728c3-p134">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="728c3-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="728c3-481">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="728c3-482">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="728c3-482">**EWS Tokens**</span></span>

<span data-ttu-id="728c3-p135">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="728c3-485">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-486">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-486">Parameters:</span></span>

|<span data-ttu-id="728c3-487">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-487">Name</span></span>| <span data-ttu-id="728c3-488">型</span><span class="sxs-lookup"><span data-stu-id="728c3-488">Type</span></span>| <span data-ttu-id="728c3-489">属性</span><span class="sxs-lookup"><span data-stu-id="728c3-489">Attributes</span></span>| <span data-ttu-id="728c3-490">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="728c3-491">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="728c3-491">Object</span></span> | <span data-ttu-id="728c3-492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-492">&lt;optional&gt;</span></span> | <span data-ttu-id="728c3-493">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="728c3-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="728c3-494">Boolean</span><span class="sxs-lookup"><span data-stu-id="728c3-494">Boolean</span></span> |  <span data-ttu-id="728c3-495">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-495">&lt;optional&gt;</span></span> | <span data-ttu-id="728c3-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="728c3-498">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-498">Object</span></span> |  <span data-ttu-id="728c3-499">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-499">&lt;optional&gt;</span></span> | <span data-ttu-id="728c3-500">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="728c3-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="728c3-501">function</span><span class="sxs-lookup"><span data-stu-id="728c3-501">function</span></span>||<span data-ttu-id="728c3-p137">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-504">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-504">Requirements</span></span>

|<span data-ttu-id="728c3-505">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-505">Requirement</span></span>| <span data-ttu-id="728c3-506">値</span><span class="sxs-lookup"><span data-stu-id="728c3-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-508">1.5</span><span class="sxs-lookup"><span data-stu-id="728c3-508">1.5</span></span> |
|[<span data-ttu-id="728c3-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-510">ReadItem</span></span>|
|[<span data-ttu-id="728c3-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-512">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="728c3-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-513">例</span><span class="sxs-lookup"><span data-stu-id="728c3-513">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="728c3-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="728c3-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="728c3-515">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="728c3-p138">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="728c3-p139">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="728c3-521">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="728c3-p140">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="728c3-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-524">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-524">Parameters:</span></span>

|<span data-ttu-id="728c3-525">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-525">Name</span></span>| <span data-ttu-id="728c3-526">型</span><span class="sxs-lookup"><span data-stu-id="728c3-526">Type</span></span>| <span data-ttu-id="728c3-527">属性</span><span class="sxs-lookup"><span data-stu-id="728c3-527">Attributes</span></span>| <span data-ttu-id="728c3-528">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="728c3-529">function</span><span class="sxs-lookup"><span data-stu-id="728c3-529">function</span></span>||<span data-ttu-id="728c3-p141">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="728c3-532">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="728c3-532">Object</span></span>| <span data-ttu-id="728c3-533">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-533">&lt;optional&gt;</span></span>|<span data-ttu-id="728c3-534">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="728c3-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-535">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-535">Requirements</span></span>

|<span data-ttu-id="728c3-536">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-536">Requirement</span></span>| <span data-ttu-id="728c3-537">値</span><span class="sxs-lookup"><span data-stu-id="728c3-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-539">1.3</span><span class="sxs-lookup"><span data-stu-id="728c3-539">1.3</span></span>|
|[<span data-ttu-id="728c3-540">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-541">ReadItem</span></span>|
|[<span data-ttu-id="728c3-542">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-543">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="728c3-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-544">例</span><span class="sxs-lookup"><span data-stu-id="728c3-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="728c3-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="728c3-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="728c3-546">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="728c3-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="728c3-547">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="728c3-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-548">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-548">Parameters:</span></span>

|<span data-ttu-id="728c3-549">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-549">Name</span></span>| <span data-ttu-id="728c3-550">型</span><span class="sxs-lookup"><span data-stu-id="728c3-550">Type</span></span>| <span data-ttu-id="728c3-551">属性</span><span class="sxs-lookup"><span data-stu-id="728c3-551">Attributes</span></span>| <span data-ttu-id="728c3-552">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="728c3-553">function</span><span class="sxs-lookup"><span data-stu-id="728c3-553">function</span></span>||<span data-ttu-id="728c3-554">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="728c3-555">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="728c3-556">Object</span><span class="sxs-lookup"><span data-stu-id="728c3-556">Object</span></span>| <span data-ttu-id="728c3-557">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-557">&lt;optional&gt;</span></span>|<span data-ttu-id="728c3-558">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="728c3-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-559">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-559">Requirements</span></span>

|<span data-ttu-id="728c3-560">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-560">Requirement</span></span>| <span data-ttu-id="728c3-561">値</span><span class="sxs-lookup"><span data-stu-id="728c3-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-562">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-563">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-563">1.0</span></span>|
|[<span data-ttu-id="728c3-564">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="728c3-565">ReadItem</span></span>|
|[<span data-ttu-id="728c3-566">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-567">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-568">例</span><span class="sxs-lookup"><span data-stu-id="728c3-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="728c3-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="728c3-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="728c3-570">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="728c3-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-571">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="728c3-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="728c3-572">IOS は、Outlook またはアプリは、Outlook で</span><span class="sxs-lookup"><span data-stu-id="728c3-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="728c3-573">アドインの読み込み時 Gmail のメールボックスに</span><span class="sxs-lookup"><span data-stu-id="728c3-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="728c3-574">アドインでは、これらの場合では、 [REST Api を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="728c3-575">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="728c3-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="728c3-576">サポートされている EWS 操作の一覧については、 [Outlook のアドインからの web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="728c3-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="728c3-577">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="728c3-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="728c3-578">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="728c3-p143">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="728c3-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="728c3-581">サーバーの管理者を設定する必要があります`OAuthAuthentication`場合は true を有効にするクライアント アクセス サーバーの EWS のディレクトリに、 `makeEwsRequestAsync` EWS を使用する方法を要求します。</span><span class="sxs-lookup"><span data-stu-id="728c3-581">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="728c3-582">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="728c3-582">Version differences</span></span>

<span data-ttu-id="728c3-583">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="728c3-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="728c3-p144">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="728c3-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="728c3-587">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="728c3-587">Parameters:</span></span>

|<span data-ttu-id="728c3-588">名前</span><span class="sxs-lookup"><span data-stu-id="728c3-588">Name</span></span>| <span data-ttu-id="728c3-589">型</span><span class="sxs-lookup"><span data-stu-id="728c3-589">Type</span></span>| <span data-ttu-id="728c3-590">属性</span><span class="sxs-lookup"><span data-stu-id="728c3-590">Attributes</span></span>| <span data-ttu-id="728c3-591">説明</span><span class="sxs-lookup"><span data-stu-id="728c3-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="728c3-592">String</span><span class="sxs-lookup"><span data-stu-id="728c3-592">String</span></span>||<span data-ttu-id="728c3-593">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="728c3-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="728c3-594">function</span><span class="sxs-lookup"><span data-stu-id="728c3-594">function</span></span>||<span data-ttu-id="728c3-595">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="728c3-596">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="728c3-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="728c3-597">結果は、サイズの 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="728c3-597">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="728c3-598">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="728c3-598">Object</span></span>| <span data-ttu-id="728c3-599">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="728c3-599">&lt;optional&gt;</span></span>|<span data-ttu-id="728c3-600">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="728c3-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="728c3-601">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-601">Requirements</span></span>

|<span data-ttu-id="728c3-602">要件</span><span class="sxs-lookup"><span data-stu-id="728c3-602">Requirement</span></span>| <span data-ttu-id="728c3-603">値</span><span class="sxs-lookup"><span data-stu-id="728c3-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="728c3-604">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="728c3-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="728c3-605">1.0</span><span class="sxs-lookup"><span data-stu-id="728c3-605">1.0</span></span>|
|[<span data-ttu-id="728c3-606">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="728c3-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="728c3-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="728c3-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="728c3-608">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="728c3-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="728c3-609">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="728c3-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="728c3-610">例</span><span class="sxs-lookup"><span data-stu-id="728c3-610">Example</span></span>

<span data-ttu-id="728c3-611">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="728c3-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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