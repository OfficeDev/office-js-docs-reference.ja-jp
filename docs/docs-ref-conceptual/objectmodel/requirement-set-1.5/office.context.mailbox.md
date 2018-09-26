
# <a name="mailbox"></a><span data-ttu-id="fd303-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="fd303-101">mailbox</span></span>

### <span data-ttu-id="fd303-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="fd303-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="fd303-104">Microsoft Outlook と web 上の Microsoft Outlook には、Outlook アドインのオブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fd303-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd303-105">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-105">Requirements</span></span>

|<span data-ttu-id="fd303-106">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-106">Requirement</span></span>| <span data-ttu-id="fd303-107">値</span><span class="sxs-lookup"><span data-stu-id="fd303-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-109">1.0</span></span>|
|[<span data-ttu-id="fd303-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="fd303-111">Restricted</span></span>|
|[<span data-ttu-id="fd303-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fd303-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-114">Members and methods</span></span>

| <span data-ttu-id="fd303-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="fd303-115">Member</span></span> | <span data-ttu-id="fd303-116">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fd303-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="fd303-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="fd303-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="fd303-118">Member</span></span> |
| [<span data-ttu-id="fd303-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="fd303-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="fd303-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="fd303-120">Member</span></span> |
| [<span data-ttu-id="fd303-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fd303-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="fd303-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-122">Method</span></span> |
| [<span data-ttu-id="fd303-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="fd303-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="fd303-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-124">Method</span></span> |
| [<span data-ttu-id="fd303-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="fd303-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="fd303-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-126">Method</span></span> |
| [<span data-ttu-id="fd303-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="fd303-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="fd303-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-128">Method</span></span> |
| [<span data-ttu-id="fd303-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="fd303-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="fd303-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-130">Method</span></span> |
| [<span data-ttu-id="fd303-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="fd303-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="fd303-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-132">Method</span></span> |
| [<span data-ttu-id="fd303-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="fd303-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="fd303-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-134">Method</span></span> |
| [<span data-ttu-id="fd303-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="fd303-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="fd303-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-136">Method</span></span> |
| [<span data-ttu-id="fd303-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fd303-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="fd303-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-138">Method</span></span> |
| [<span data-ttu-id="fd303-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fd303-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="fd303-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-140">Method</span></span> |
| [<span data-ttu-id="fd303-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fd303-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="fd303-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-142">Method</span></span> |
| [<span data-ttu-id="fd303-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="fd303-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="fd303-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fd303-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="fd303-145">Namespaces</span></span>

<span data-ttu-id="fd303-146">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="fd303-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="fd303-147">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="fd303-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="fd303-148">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="fd303-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="fd303-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="fd303-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="fd303-150">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="fd303-150">ewsUrl :String</span></span>

<span data-ttu-id="fd303-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="fd303-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-153">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-153">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-p103">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="fd303-156">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="fd303-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="fd303-159">型:</span><span class="sxs-lookup"><span data-stu-id="fd303-159">Type:</span></span>

*   <span data-ttu-id="fd303-160">String</span><span class="sxs-lookup"><span data-stu-id="fd303-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd303-161">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-161">Requirements</span></span>

|<span data-ttu-id="fd303-162">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-162">Requirement</span></span>| <span data-ttu-id="fd303-163">値</span><span class="sxs-lookup"><span data-stu-id="fd303-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-165">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-165">1.0</span></span>|
|[<span data-ttu-id="fd303-166">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-167">ReadItem</span></span>|
|[<span data-ttu-id="fd303-168">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-169">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="fd303-170">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="fd303-170">restUrl :String</span></span>

<span data-ttu-id="fd303-171">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="fd303-172">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="fd303-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="fd303-173">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="fd303-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-176">Outlook クライアントが Exchange の 2016 年の設置型インストールに接続されているか、構成されている他のカスタム URL を後でもうは無効な値を返す`restUrl`。</span><span class="sxs-lookup"><span data-stu-id="fd303-176">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="fd303-177">型:</span><span class="sxs-lookup"><span data-stu-id="fd303-177">Type:</span></span>

*   <span data-ttu-id="fd303-178">String</span><span class="sxs-lookup"><span data-stu-id="fd303-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd303-179">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-179">Requirements</span></span>

|<span data-ttu-id="fd303-180">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-180">Requirement</span></span>| <span data-ttu-id="fd303-181">値</span><span class="sxs-lookup"><span data-stu-id="fd303-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-183">1.5</span><span class="sxs-lookup"><span data-stu-id="fd303-183">1.5</span></span> |
|[<span data-ttu-id="fd303-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-185">ReadItem</span></span>|
|[<span data-ttu-id="fd303-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="fd303-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="fd303-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="fd303-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fd303-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="fd303-190">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="fd303-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="fd303-p106">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="fd303-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-193">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-193">Parameters:</span></span>

| <span data-ttu-id="fd303-194">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-194">Name</span></span> | <span data-ttu-id="fd303-195">型</span><span class="sxs-lookup"><span data-stu-id="fd303-195">Type</span></span> | <span data-ttu-id="fd303-196">属性</span><span class="sxs-lookup"><span data-stu-id="fd303-196">Attributes</span></span> | <span data-ttu-id="fd303-197">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fd303-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fd303-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fd303-199">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="fd303-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="fd303-200">Function</span><span class="sxs-lookup"><span data-stu-id="fd303-200">Function</span></span> || <span data-ttu-id="fd303-p107">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="fd303-204">Object</span><span class="sxs-lookup"><span data-stu-id="fd303-204">Object</span></span> | <span data-ttu-id="fd303-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-205">&lt;optional&gt;</span></span> | <span data-ttu-id="fd303-206">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fd303-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fd303-207">Object</span><span class="sxs-lookup"><span data-stu-id="fd303-207">Object</span></span> | <span data-ttu-id="fd303-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-208">&lt;optional&gt;</span></span> | <span data-ttu-id="fd303-209">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fd303-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fd303-210">function</span><span class="sxs-lookup"><span data-stu-id="fd303-210">function</span></span>| <span data-ttu-id="fd303-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-211">&lt;optional&gt;</span></span>|<span data-ttu-id="fd303-212">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-213">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-213">Requirements</span></span>

|<span data-ttu-id="fd303-214">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-214">Requirement</span></span>| <span data-ttu-id="fd303-215">値</span><span class="sxs-lookup"><span data-stu-id="fd303-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-216">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-217">1.5</span><span class="sxs-lookup"><span data-stu-id="fd303-217">1.5</span></span> |
|[<span data-ttu-id="fd303-218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-219">ReadItem</span></span> |
|[<span data-ttu-id="fd303-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-221">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-222">例</span><span class="sxs-lookup"><span data-stu-id="fd303-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="fd303-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="fd303-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="fd303-224">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="fd303-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-225">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-228">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-228">Parameters:</span></span>

|<span data-ttu-id="fd303-229">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-229">Name</span></span>| <span data-ttu-id="fd303-230">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-230">Type</span></span>| <span data-ttu-id="fd303-231">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fd303-232">String</span><span class="sxs-lookup"><span data-stu-id="fd303-232">String</span></span>|<span data-ttu-id="fd303-233">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="fd303-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="fd303-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="fd303-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="fd303-235">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="fd303-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-236">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-236">Requirements</span></span>

|<span data-ttu-id="fd303-237">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-237">Requirement</span></span>| <span data-ttu-id="fd303-238">値</span><span class="sxs-lookup"><span data-stu-id="fd303-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-240">1.3</span><span class="sxs-lookup"><span data-stu-id="fd303-240">1.3</span></span>|
|[<span data-ttu-id="fd303-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-242">制限あり</span><span class="sxs-lookup"><span data-stu-id="fd303-242">Restricted</span></span>|
|[<span data-ttu-id="fd303-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-244">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd303-245">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fd303-245">Returns:</span></span>

<span data-ttu-id="fd303-246">型:String</span><span class="sxs-lookup"><span data-stu-id="fd303-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="fd303-247">例</span><span class="sxs-lookup"><span data-stu-id="fd303-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="fd303-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="fd303-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="fd303-249">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="fd303-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="fd303-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-255">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-255">Parameters:</span></span>

|<span data-ttu-id="fd303-256">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-256">Name</span></span>| <span data-ttu-id="fd303-257">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-257">Type</span></span>| <span data-ttu-id="fd303-258">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="fd303-259">日付</span><span class="sxs-lookup"><span data-stu-id="fd303-259">Date</span></span>|<span data-ttu-id="fd303-260">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fd303-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-261">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-261">Requirements</span></span>

|<span data-ttu-id="fd303-262">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-262">Requirement</span></span>| <span data-ttu-id="fd303-263">値</span><span class="sxs-lookup"><span data-stu-id="fd303-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-265">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-265">1.0</span></span>|
|[<span data-ttu-id="fd303-266">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-267">ReadItem</span></span>|
|[<span data-ttu-id="fd303-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-269">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd303-270">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fd303-270">Returns:</span></span>

<span data-ttu-id="fd303-271">型:[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="fd303-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="fd303-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="fd303-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="fd303-273">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="fd303-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-274">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-p111">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-277">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-277">Parameters:</span></span>

|<span data-ttu-id="fd303-278">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-278">Name</span></span>| <span data-ttu-id="fd303-279">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-279">Type</span></span>| <span data-ttu-id="fd303-280">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fd303-281">String</span><span class="sxs-lookup"><span data-stu-id="fd303-281">String</span></span>|<span data-ttu-id="fd303-282">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="fd303-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="fd303-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="fd303-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="fd303-284">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="fd303-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-285">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-285">Requirements</span></span>

|<span data-ttu-id="fd303-286">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-286">Requirement</span></span>| <span data-ttu-id="fd303-287">値</span><span class="sxs-lookup"><span data-stu-id="fd303-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-289">1.3</span><span class="sxs-lookup"><span data-stu-id="fd303-289">1.3</span></span>|
|[<span data-ttu-id="fd303-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-291">制限あり</span><span class="sxs-lookup"><span data-stu-id="fd303-291">Restricted</span></span>|
|[<span data-ttu-id="fd303-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-293">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd303-294">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fd303-294">Returns:</span></span>

<span data-ttu-id="fd303-295">型:String</span><span class="sxs-lookup"><span data-stu-id="fd303-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="fd303-296">例</span><span class="sxs-lookup"><span data-stu-id="fd303-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="fd303-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="fd303-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="fd303-298">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="fd303-299">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="fd303-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-300">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-300">Parameters:</span></span>

|<span data-ttu-id="fd303-301">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-301">Name</span></span>| <span data-ttu-id="fd303-302">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-302">Type</span></span>| <span data-ttu-id="fd303-303">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="fd303-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="fd303-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="fd303-305">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="fd303-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-306">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-306">Requirements</span></span>

|<span data-ttu-id="fd303-307">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-307">Requirement</span></span>| <span data-ttu-id="fd303-308">値</span><span class="sxs-lookup"><span data-stu-id="fd303-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-309">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-310">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-310">1.0</span></span>|
|[<span data-ttu-id="fd303-311">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-312">ReadItem</span></span>|
|[<span data-ttu-id="fd303-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-314">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd303-315">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fd303-315">Returns:</span></span>

<span data-ttu-id="fd303-316">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="fd303-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="fd303-317">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="fd303-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fd303-318">Date</span><span class="sxs-lookup"><span data-stu-id="fd303-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="fd303-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="fd303-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="fd303-320">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="fd303-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-321">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-322">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="fd303-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="fd303-p112">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="fd303-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="fd303-325">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="fd303-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="fd303-326">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="fd303-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-327">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-327">Parameters:</span></span>

|<span data-ttu-id="fd303-328">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-328">Name</span></span>| <span data-ttu-id="fd303-329">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-329">Type</span></span>| <span data-ttu-id="fd303-330">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fd303-331">String</span><span class="sxs-lookup"><span data-stu-id="fd303-331">String</span></span>|<span data-ttu-id="fd303-332">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="fd303-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-333">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-333">Requirements</span></span>

|<span data-ttu-id="fd303-334">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-334">Requirement</span></span>| <span data-ttu-id="fd303-335">値</span><span class="sxs-lookup"><span data-stu-id="fd303-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-336">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-337">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-337">1.0</span></span>|
|[<span data-ttu-id="fd303-338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-339">ReadItem</span></span>|
|[<span data-ttu-id="fd303-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-341">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-342">例</span><span class="sxs-lookup"><span data-stu-id="fd303-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="fd303-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="fd303-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="fd303-344">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="fd303-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-345">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-346">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="fd303-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="fd303-347">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="fd303-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="fd303-348">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="fd303-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="fd303-p113">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-351">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-351">Parameters:</span></span>

|<span data-ttu-id="fd303-352">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-352">Name</span></span>| <span data-ttu-id="fd303-353">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-353">Type</span></span>| <span data-ttu-id="fd303-354">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fd303-355">String</span><span class="sxs-lookup"><span data-stu-id="fd303-355">String</span></span>|<span data-ttu-id="fd303-356">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="fd303-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-357">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-357">Requirements</span></span>

|<span data-ttu-id="fd303-358">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-358">Requirement</span></span>| <span data-ttu-id="fd303-359">値</span><span class="sxs-lookup"><span data-stu-id="fd303-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-361">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-361">1.0</span></span>|
|[<span data-ttu-id="fd303-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-363">ReadItem</span></span>|
|[<span data-ttu-id="fd303-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-365">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-366">例</span><span class="sxs-lookup"><span data-stu-id="fd303-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="fd303-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="fd303-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="fd303-368">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="fd303-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-369">IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd303-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="fd303-p115">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="fd303-p116">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="fd303-377">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="fd303-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-378">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-378">Parameters:</span></span>

|<span data-ttu-id="fd303-379">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-379">Name</span></span>| <span data-ttu-id="fd303-380">種類</span><span class="sxs-lookup"><span data-stu-id="fd303-380">Type</span></span>| <span data-ttu-id="fd303-381">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="fd303-382">Object</span><span class="sxs-lookup"><span data-stu-id="fd303-382">Object</span></span> | <span data-ttu-id="fd303-383">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="fd303-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="fd303-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fd303-p117">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="fd303-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="fd303-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fd303-p118">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="fd303-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="fd303-390">日付</span><span class="sxs-lookup"><span data-stu-id="fd303-390">Date</span></span> | <span data-ttu-id="fd303-391">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="fd303-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="fd303-392">Date</span><span class="sxs-lookup"><span data-stu-id="fd303-392">Date</span></span> | <span data-ttu-id="fd303-393">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="fd303-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="fd303-394">String</span><span class="sxs-lookup"><span data-stu-id="fd303-394">String</span></span> | <span data-ttu-id="fd303-p119">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="fd303-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="fd303-p120">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="fd303-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="fd303-400">String</span><span class="sxs-lookup"><span data-stu-id="fd303-400">String</span></span> | <span data-ttu-id="fd303-p121">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="fd303-403">String</span><span class="sxs-lookup"><span data-stu-id="fd303-403">String</span></span> | <span data-ttu-id="fd303-p122">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd303-406">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-406">Requirements</span></span>

|<span data-ttu-id="fd303-407">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-407">Requirement</span></span>| <span data-ttu-id="fd303-408">値</span><span class="sxs-lookup"><span data-stu-id="fd303-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-410">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-410">1.0</span></span>|
|[<span data-ttu-id="fd303-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-412">ReadItem</span></span>|
|[<span data-ttu-id="fd303-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-415">例</span><span class="sxs-lookup"><span data-stu-id="fd303-415">Example</span></span>

```js
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="fd303-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fd303-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="fd303-417">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="fd303-p123">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-420">アドインが可能な場合に、Exchange Web サービスではなく REST Api を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="fd303-420">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="fd303-421">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="fd303-421">**REST Tokens**</span></span>

<span data-ttu-id="fd303-p124">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="fd303-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="fd303-425">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="fd303-426">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="fd303-426">**EWS Tokens**</span></span>

<span data-ttu-id="fd303-p125">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="fd303-429">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-430">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-430">Parameters:</span></span>

|<span data-ttu-id="fd303-431">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-431">Name</span></span>| <span data-ttu-id="fd303-432">型</span><span class="sxs-lookup"><span data-stu-id="fd303-432">Type</span></span>| <span data-ttu-id="fd303-433">属性</span><span class="sxs-lookup"><span data-stu-id="fd303-433">Attributes</span></span>| <span data-ttu-id="fd303-434">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="fd303-435">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fd303-435">Object</span></span> | <span data-ttu-id="fd303-436">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-436">&lt;optional&gt;</span></span> | <span data-ttu-id="fd303-437">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fd303-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="fd303-438">Boolean</span><span class="sxs-lookup"><span data-stu-id="fd303-438">Boolean</span></span> |  <span data-ttu-id="fd303-439">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-439">&lt;optional&gt;</span></span> | <span data-ttu-id="fd303-p126">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fd303-442">Object</span><span class="sxs-lookup"><span data-stu-id="fd303-442">Object</span></span> |  <span data-ttu-id="fd303-443">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-443">&lt;optional&gt;</span></span> | <span data-ttu-id="fd303-444">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="fd303-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="fd303-445">function</span><span class="sxs-lookup"><span data-stu-id="fd303-445">function</span></span>||<span data-ttu-id="fd303-p127">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-448">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-448">Requirements</span></span>

|<span data-ttu-id="fd303-449">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-449">Requirement</span></span>| <span data-ttu-id="fd303-450">値</span><span class="sxs-lookup"><span data-stu-id="fd303-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-451">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-452">1.5</span><span class="sxs-lookup"><span data-stu-id="fd303-452">1.5</span></span> |
|[<span data-ttu-id="fd303-453">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-454">ReadItem</span></span>|
|[<span data-ttu-id="fd303-455">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-456">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="fd303-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-457">例</span><span class="sxs-lookup"><span data-stu-id="fd303-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="fd303-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fd303-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="fd303-459">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="fd303-p128">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="fd303-p129">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="fd303-465">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="fd303-p130">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="fd303-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-468">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-468">Parameters:</span></span>

|<span data-ttu-id="fd303-469">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-469">Name</span></span>| <span data-ttu-id="fd303-470">型</span><span class="sxs-lookup"><span data-stu-id="fd303-470">Type</span></span>| <span data-ttu-id="fd303-471">属性</span><span class="sxs-lookup"><span data-stu-id="fd303-471">Attributes</span></span>| <span data-ttu-id="fd303-472">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fd303-473">function</span><span class="sxs-lookup"><span data-stu-id="fd303-473">function</span></span>||<span data-ttu-id="fd303-p131">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="fd303-476">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fd303-476">Object</span></span>| <span data-ttu-id="fd303-477">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-477">&lt;optional&gt;</span></span>|<span data-ttu-id="fd303-478">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="fd303-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-479">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-479">Requirements</span></span>

|<span data-ttu-id="fd303-480">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-480">Requirement</span></span>| <span data-ttu-id="fd303-481">値</span><span class="sxs-lookup"><span data-stu-id="fd303-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-482">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-483">1.3</span><span class="sxs-lookup"><span data-stu-id="fd303-483">1.3</span></span>|
|[<span data-ttu-id="fd303-484">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-485">ReadItem</span></span>|
|[<span data-ttu-id="fd303-486">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-487">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="fd303-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-488">例</span><span class="sxs-lookup"><span data-stu-id="fd303-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="fd303-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fd303-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="fd303-490">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="fd303-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="fd303-491">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="fd303-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-492">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-492">Parameters:</span></span>

|<span data-ttu-id="fd303-493">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-493">Name</span></span>| <span data-ttu-id="fd303-494">型</span><span class="sxs-lookup"><span data-stu-id="fd303-494">Type</span></span>| <span data-ttu-id="fd303-495">属性</span><span class="sxs-lookup"><span data-stu-id="fd303-495">Attributes</span></span>| <span data-ttu-id="fd303-496">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fd303-497">function</span><span class="sxs-lookup"><span data-stu-id="fd303-497">function</span></span>||<span data-ttu-id="fd303-498">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fd303-499">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="fd303-500">Object</span><span class="sxs-lookup"><span data-stu-id="fd303-500">Object</span></span>| <span data-ttu-id="fd303-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-501">&lt;optional&gt;</span></span>|<span data-ttu-id="fd303-502">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="fd303-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-503">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-503">Requirements</span></span>

|<span data-ttu-id="fd303-504">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-504">Requirement</span></span>| <span data-ttu-id="fd303-505">値</span><span class="sxs-lookup"><span data-stu-id="fd303-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-507">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-507">1.0</span></span>|
|[<span data-ttu-id="fd303-508">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd303-509">ReadItem</span></span>|
|[<span data-ttu-id="fd303-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-511">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-512">例</span><span class="sxs-lookup"><span data-stu-id="fd303-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="fd303-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fd303-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="fd303-514">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="fd303-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-515">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fd303-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="fd303-516">IOS は、Outlook またはアプリは、Outlook で</span><span class="sxs-lookup"><span data-stu-id="fd303-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="fd303-517">アドインの読み込み時 Gmail のメールボックスに</span><span class="sxs-lookup"><span data-stu-id="fd303-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="fd303-518">アドインでは、これらの場合では、 [REST Api を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="fd303-519">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="fd303-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="fd303-520">サポートされている EWS 操作の一覧については、 [Outlook のアドインからの web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd303-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="fd303-521">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="fd303-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="fd303-522">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="fd303-p133">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd303-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="fd303-525">サーバーの管理者を設定する必要があります`OAuthAuthentication`場合は true を有効にするクライアント アクセス サーバーの EWS のディレクトリに、 `makeEwsRequestAsync` EWS を使用する方法を要求します。</span><span class="sxs-lookup"><span data-stu-id="fd303-525">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="fd303-526">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="fd303-526">Version differences</span></span>

<span data-ttu-id="fd303-527">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd303-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="fd303-p134">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="fd303-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd303-531">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="fd303-531">Parameters:</span></span>

|<span data-ttu-id="fd303-532">名前</span><span class="sxs-lookup"><span data-stu-id="fd303-532">Name</span></span>| <span data-ttu-id="fd303-533">型</span><span class="sxs-lookup"><span data-stu-id="fd303-533">Type</span></span>| <span data-ttu-id="fd303-534">属性</span><span class="sxs-lookup"><span data-stu-id="fd303-534">Attributes</span></span>| <span data-ttu-id="fd303-535">説明</span><span class="sxs-lookup"><span data-stu-id="fd303-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fd303-536">String</span><span class="sxs-lookup"><span data-stu-id="fd303-536">String</span></span>||<span data-ttu-id="fd303-537">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="fd303-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="fd303-538">function</span><span class="sxs-lookup"><span data-stu-id="fd303-538">function</span></span>||<span data-ttu-id="fd303-539">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fd303-540">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="fd303-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="fd303-541">結果は、サイズの 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="fd303-541">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="fd303-542">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fd303-542">Object</span></span>| <span data-ttu-id="fd303-543">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="fd303-543">&lt;optional&gt;</span></span>|<span data-ttu-id="fd303-544">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="fd303-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd303-545">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-545">Requirements</span></span>

|<span data-ttu-id="fd303-546">要件</span><span class="sxs-lookup"><span data-stu-id="fd303-546">Requirement</span></span>| <span data-ttu-id="fd303-547">値</span><span class="sxs-lookup"><span data-stu-id="fd303-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd303-548">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fd303-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd303-549">1.0</span><span class="sxs-lookup"><span data-stu-id="fd303-549">1.0</span></span>|
|[<span data-ttu-id="fd303-550">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fd303-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd303-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="fd303-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="fd303-552">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fd303-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fd303-553">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fd303-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd303-554">例</span><span class="sxs-lookup"><span data-stu-id="fd303-554">Example</span></span>

<span data-ttu-id="fd303-555">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="fd303-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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