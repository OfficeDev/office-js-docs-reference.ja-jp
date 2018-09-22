 

# <a name="office"></a><span data-ttu-id="9e306-101">Office</span><span class="sxs-lookup"><span data-stu-id="9e306-101">Office</span></span>

<span data-ttu-id="9e306-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9e306-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e306-104">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-104">Requirements</span></span>

|<span data-ttu-id="9e306-105">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-105">Requirement</span></span>| <span data-ttu-id="9e306-106">値</span><span class="sxs-lookup"><span data-stu-id="9e306-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e306-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e306-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e306-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9e306-108">1.0</span></span>|
|[<span data-ttu-id="9e306-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e306-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e306-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e306-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9e306-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9e306-111">Members and methods</span></span>

| <span data-ttu-id="9e306-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-112">Member</span></span> | <span data-ttu-id="9e306-113">種類</span><span class="sxs-lookup"><span data-stu-id="9e306-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9e306-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9e306-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9e306-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-115">Member</span></span> |
| [<span data-ttu-id="9e306-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9e306-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9e306-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-117">Member</span></span> |
| [<span data-ttu-id="9e306-118">EventType</span><span class="sxs-lookup"><span data-stu-id="9e306-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9e306-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-119">Member</span></span> |
| [<span data-ttu-id="9e306-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9e306-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9e306-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9e306-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="9e306-122">Namespaces</span></span>

<span data-ttu-id="9e306-123">[コンテキスト](office.context.md): Outlook アドイン API で使用するための Office アドイン API のコンテキストの名前空間からの共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="9e306-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9e306-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="9e306-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9e306-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e306-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9e306-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9e306-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="9e306-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="9e306-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9e306-128">型:</span><span class="sxs-lookup"><span data-stu-id="9e306-128">Type:</span></span>

*   <span data-ttu-id="9e306-129">String</span><span class="sxs-lookup"><span data-stu-id="9e306-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e306-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e306-130">Properties:</span></span>

|<span data-ttu-id="9e306-131">名前</span><span class="sxs-lookup"><span data-stu-id="9e306-131">Name</span></span>| <span data-ttu-id="9e306-132">種類</span><span class="sxs-lookup"><span data-stu-id="9e306-132">Type</span></span>| <span data-ttu-id="9e306-133">説明</span><span class="sxs-lookup"><span data-stu-id="9e306-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9e306-134">String</span><span class="sxs-lookup"><span data-stu-id="9e306-134">String</span></span>|<span data-ttu-id="9e306-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="9e306-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9e306-136">String</span><span class="sxs-lookup"><span data-stu-id="9e306-136">String</span></span>|<span data-ttu-id="9e306-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9e306-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e306-138">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-138">Requirements</span></span>

|<span data-ttu-id="9e306-139">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-139">Requirement</span></span>| <span data-ttu-id="9e306-140">値</span><span class="sxs-lookup"><span data-stu-id="9e306-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e306-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e306-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e306-142">1.0</span><span class="sxs-lookup"><span data-stu-id="9e306-142">1.0</span></span>|
|[<span data-ttu-id="9e306-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e306-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e306-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e306-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9e306-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9e306-145">CoercionType :String</span></span>

<span data-ttu-id="9e306-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="9e306-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e306-147">型:</span><span class="sxs-lookup"><span data-stu-id="9e306-147">Type:</span></span>

*   <span data-ttu-id="9e306-148">String</span><span class="sxs-lookup"><span data-stu-id="9e306-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e306-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e306-149">Properties:</span></span>

|<span data-ttu-id="9e306-150">名前</span><span class="sxs-lookup"><span data-stu-id="9e306-150">Name</span></span>| <span data-ttu-id="9e306-151">種類</span><span class="sxs-lookup"><span data-stu-id="9e306-151">Type</span></span>| <span data-ttu-id="9e306-152">説明</span><span class="sxs-lookup"><span data-stu-id="9e306-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9e306-153">String</span><span class="sxs-lookup"><span data-stu-id="9e306-153">String</span></span>|<span data-ttu-id="9e306-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9e306-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9e306-155">String</span><span class="sxs-lookup"><span data-stu-id="9e306-155">String</span></span>|<span data-ttu-id="9e306-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9e306-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e306-157">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-157">Requirements</span></span>

|<span data-ttu-id="9e306-158">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-158">Requirement</span></span>| <span data-ttu-id="9e306-159">値</span><span class="sxs-lookup"><span data-stu-id="9e306-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e306-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e306-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e306-161">1.0</span><span class="sxs-lookup"><span data-stu-id="9e306-161">1.0</span></span>|
|[<span data-ttu-id="9e306-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e306-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e306-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e306-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9e306-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="9e306-164">EventType :String</span></span>

<span data-ttu-id="9e306-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="9e306-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9e306-166">型:</span><span class="sxs-lookup"><span data-stu-id="9e306-166">Type:</span></span>

*   <span data-ttu-id="9e306-167">String</span><span class="sxs-lookup"><span data-stu-id="9e306-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e306-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e306-168">Properties:</span></span>

| <span data-ttu-id="9e306-169">名前</span><span class="sxs-lookup"><span data-stu-id="9e306-169">Name</span></span> | <span data-ttu-id="9e306-170">種類</span><span class="sxs-lookup"><span data-stu-id="9e306-170">Type</span></span> | <span data-ttu-id="9e306-171">説明</span><span class="sxs-lookup"><span data-stu-id="9e306-171">Description</span></span> | <span data-ttu-id="9e306-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="9e306-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="9e306-173">String</span><span class="sxs-lookup"><span data-stu-id="9e306-173">String</span></span> | <span data-ttu-id="9e306-174">日付または時刻を選択した予定の系列が変更されました。</span><span class="sxs-lookup"><span data-stu-id="9e306-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9e306-175">1.7</span><span class="sxs-lookup"><span data-stu-id="9e306-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="9e306-176">String</span><span class="sxs-lookup"><span data-stu-id="9e306-176">String</span></span> | <span data-ttu-id="9e306-177">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="9e306-177">The selected item has changed.</span></span> | <span data-ttu-id="9e306-178">1.5</span><span class="sxs-lookup"><span data-stu-id="9e306-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9e306-179">String</span><span class="sxs-lookup"><span data-stu-id="9e306-179">String</span></span> | <span data-ttu-id="9e306-180">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="9e306-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9e306-181">1.7</span><span class="sxs-lookup"><span data-stu-id="9e306-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9e306-182">String</span><span class="sxs-lookup"><span data-stu-id="9e306-182">String</span></span> | <span data-ttu-id="9e306-183">選択した一連の定期的なパターンが変更されています。</span><span class="sxs-lookup"><span data-stu-id="9e306-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9e306-184">1.7</span><span class="sxs-lookup"><span data-stu-id="9e306-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e306-185">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-185">Requirements</span></span>

|<span data-ttu-id="9e306-186">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-186">Requirement</span></span>| <span data-ttu-id="9e306-187">値</span><span class="sxs-lookup"><span data-stu-id="9e306-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e306-188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e306-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e306-189">1.5</span><span class="sxs-lookup"><span data-stu-id="9e306-189">1.5</span></span> |
|[<span data-ttu-id="9e306-190">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e306-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e306-191">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e306-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9e306-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9e306-192">SourceProperty :String</span></span>

<span data-ttu-id="9e306-193">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="9e306-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e306-194">型:</span><span class="sxs-lookup"><span data-stu-id="9e306-194">Type:</span></span>

*   <span data-ttu-id="9e306-195">String</span><span class="sxs-lookup"><span data-stu-id="9e306-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e306-196">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9e306-196">Properties:</span></span>

|<span data-ttu-id="9e306-197">名前</span><span class="sxs-lookup"><span data-stu-id="9e306-197">Name</span></span>| <span data-ttu-id="9e306-198">種類</span><span class="sxs-lookup"><span data-stu-id="9e306-198">Type</span></span>| <span data-ttu-id="9e306-199">説明</span><span class="sxs-lookup"><span data-stu-id="9e306-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9e306-200">String</span><span class="sxs-lookup"><span data-stu-id="9e306-200">String</span></span>|<span data-ttu-id="9e306-201">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="9e306-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9e306-202">String</span><span class="sxs-lookup"><span data-stu-id="9e306-202">String</span></span>|<span data-ttu-id="9e306-203">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="9e306-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e306-204">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-204">Requirements</span></span>

|<span data-ttu-id="9e306-205">要件</span><span class="sxs-lookup"><span data-stu-id="9e306-205">Requirement</span></span>| <span data-ttu-id="9e306-206">値</span><span class="sxs-lookup"><span data-stu-id="9e306-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e306-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e306-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e306-208">1.0</span><span class="sxs-lookup"><span data-stu-id="9e306-208">1.0</span></span>|
|[<span data-ttu-id="9e306-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e306-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e306-210">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9e306-210">Compose or read</span></span>|