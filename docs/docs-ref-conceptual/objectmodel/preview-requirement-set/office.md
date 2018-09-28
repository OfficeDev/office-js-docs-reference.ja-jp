 

# <a name="office"></a><span data-ttu-id="93b19-101">Office</span><span class="sxs-lookup"><span data-stu-id="93b19-101">Office</span></span>

<span data-ttu-id="93b19-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="93b19-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="93b19-104">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-104">Requirements</span></span>

|<span data-ttu-id="93b19-105">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-105">Requirement</span></span>| <span data-ttu-id="93b19-106">値</span><span class="sxs-lookup"><span data-stu-id="93b19-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="93b19-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="93b19-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93b19-108">1.0</span><span class="sxs-lookup"><span data-stu-id="93b19-108">1.0</span></span>|
|[<span data-ttu-id="93b19-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="93b19-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93b19-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="93b19-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="93b19-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="93b19-111">Members and methods</span></span>

| <span data-ttu-id="93b19-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-112">Member</span></span> | <span data-ttu-id="93b19-113">種類</span><span class="sxs-lookup"><span data-stu-id="93b19-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="93b19-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="93b19-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="93b19-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-115">Member</span></span> |
| [<span data-ttu-id="93b19-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="93b19-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="93b19-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-117">Member</span></span> |
| [<span data-ttu-id="93b19-118">EventType</span><span class="sxs-lookup"><span data-stu-id="93b19-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="93b19-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-119">Member</span></span> |
| [<span data-ttu-id="93b19-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="93b19-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="93b19-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="93b19-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="93b19-122">Namespaces</span></span>

<span data-ttu-id="93b19-123">[コンテキスト](office.context.md): Outlook アドイン API で使用するための Office アドイン API のコンテキストの名前空間からの共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="93b19-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="93b19-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="93b19-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="93b19-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="93b19-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="93b19-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="93b19-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="93b19-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="93b19-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="93b19-128">型:</span><span class="sxs-lookup"><span data-stu-id="93b19-128">Type:</span></span>

*   <span data-ttu-id="93b19-129">String</span><span class="sxs-lookup"><span data-stu-id="93b19-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93b19-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="93b19-130">Properties:</span></span>

|<span data-ttu-id="93b19-131">名前</span><span class="sxs-lookup"><span data-stu-id="93b19-131">Name</span></span>| <span data-ttu-id="93b19-132">種類</span><span class="sxs-lookup"><span data-stu-id="93b19-132">Type</span></span>| <span data-ttu-id="93b19-133">説明</span><span class="sxs-lookup"><span data-stu-id="93b19-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="93b19-134">String</span><span class="sxs-lookup"><span data-stu-id="93b19-134">String</span></span>|<span data-ttu-id="93b19-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="93b19-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="93b19-136">String</span><span class="sxs-lookup"><span data-stu-id="93b19-136">String</span></span>|<span data-ttu-id="93b19-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="93b19-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93b19-138">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-138">Requirements</span></span>

|<span data-ttu-id="93b19-139">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-139">Requirement</span></span>| <span data-ttu-id="93b19-140">値</span><span class="sxs-lookup"><span data-stu-id="93b19-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="93b19-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="93b19-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93b19-142">1.0</span><span class="sxs-lookup"><span data-stu-id="93b19-142">1.0</span></span>|
|[<span data-ttu-id="93b19-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="93b19-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93b19-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="93b19-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="93b19-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="93b19-145">CoercionType :String</span></span>

<span data-ttu-id="93b19-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="93b19-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93b19-147">型:</span><span class="sxs-lookup"><span data-stu-id="93b19-147">Type:</span></span>

*   <span data-ttu-id="93b19-148">String</span><span class="sxs-lookup"><span data-stu-id="93b19-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93b19-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="93b19-149">Properties:</span></span>

|<span data-ttu-id="93b19-150">名前</span><span class="sxs-lookup"><span data-stu-id="93b19-150">Name</span></span>| <span data-ttu-id="93b19-151">種類</span><span class="sxs-lookup"><span data-stu-id="93b19-151">Type</span></span>| <span data-ttu-id="93b19-152">説明</span><span class="sxs-lookup"><span data-stu-id="93b19-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="93b19-153">String</span><span class="sxs-lookup"><span data-stu-id="93b19-153">String</span></span>|<span data-ttu-id="93b19-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="93b19-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="93b19-155">String</span><span class="sxs-lookup"><span data-stu-id="93b19-155">String</span></span>|<span data-ttu-id="93b19-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="93b19-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93b19-157">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-157">Requirements</span></span>

|<span data-ttu-id="93b19-158">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-158">Requirement</span></span>| <span data-ttu-id="93b19-159">値</span><span class="sxs-lookup"><span data-stu-id="93b19-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="93b19-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="93b19-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93b19-161">1.0</span><span class="sxs-lookup"><span data-stu-id="93b19-161">1.0</span></span>|
|[<span data-ttu-id="93b19-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="93b19-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93b19-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="93b19-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="93b19-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="93b19-164">EventType :String</span></span>

<span data-ttu-id="93b19-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="93b19-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="93b19-166">型:</span><span class="sxs-lookup"><span data-stu-id="93b19-166">Type:</span></span>

*   <span data-ttu-id="93b19-167">String</span><span class="sxs-lookup"><span data-stu-id="93b19-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93b19-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="93b19-168">Properties:</span></span>

| <span data-ttu-id="93b19-169">名前</span><span class="sxs-lookup"><span data-stu-id="93b19-169">Name</span></span> | <span data-ttu-id="93b19-170">種類</span><span class="sxs-lookup"><span data-stu-id="93b19-170">Type</span></span> | <span data-ttu-id="93b19-171">説明</span><span class="sxs-lookup"><span data-stu-id="93b19-171">Description</span></span> | <span data-ttu-id="93b19-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="93b19-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="93b19-173">String</span><span class="sxs-lookup"><span data-stu-id="93b19-173">String</span></span> | <span data-ttu-id="93b19-174">日付または時刻を選択した予定の系列が変更されました。</span><span class="sxs-lookup"><span data-stu-id="93b19-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="93b19-175">1.7</span><span class="sxs-lookup"><span data-stu-id="93b19-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="93b19-176">String</span><span class="sxs-lookup"><span data-stu-id="93b19-176">String</span></span> | <span data-ttu-id="93b19-177">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="93b19-177">The selected item has changed.</span></span> | <span data-ttu-id="93b19-178">1.5</span><span class="sxs-lookup"><span data-stu-id="93b19-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="93b19-179">String</span><span class="sxs-lookup"><span data-stu-id="93b19-179">String</span></span> | <span data-ttu-id="93b19-180">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="93b19-180">The selected item has changed.</span></span> | <span data-ttu-id="93b19-181">Preview</span><span class="sxs-lookup"><span data-stu-id="93b19-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="93b19-182">String</span><span class="sxs-lookup"><span data-stu-id="93b19-182">String</span></span> | <span data-ttu-id="93b19-183">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="93b19-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="93b19-184">1.7</span><span class="sxs-lookup"><span data-stu-id="93b19-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="93b19-185">String</span><span class="sxs-lookup"><span data-stu-id="93b19-185">String</span></span> | <span data-ttu-id="93b19-186">選択した一連の定期的なパターンが変更されています。</span><span class="sxs-lookup"><span data-stu-id="93b19-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="93b19-187">1.7</span><span class="sxs-lookup"><span data-stu-id="93b19-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="93b19-188">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-188">Requirements</span></span>

|<span data-ttu-id="93b19-189">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-189">Requirement</span></span>| <span data-ttu-id="93b19-190">値</span><span class="sxs-lookup"><span data-stu-id="93b19-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="93b19-191">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="93b19-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93b19-192">1.5</span><span class="sxs-lookup"><span data-stu-id="93b19-192">1.5</span></span> |
|[<span data-ttu-id="93b19-193">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="93b19-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93b19-194">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="93b19-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="93b19-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="93b19-195">SourceProperty :String</span></span>

<span data-ttu-id="93b19-196">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="93b19-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93b19-197">型:</span><span class="sxs-lookup"><span data-stu-id="93b19-197">Type:</span></span>

*   <span data-ttu-id="93b19-198">String</span><span class="sxs-lookup"><span data-stu-id="93b19-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93b19-199">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="93b19-199">Properties:</span></span>

|<span data-ttu-id="93b19-200">名前</span><span class="sxs-lookup"><span data-stu-id="93b19-200">Name</span></span>| <span data-ttu-id="93b19-201">種類</span><span class="sxs-lookup"><span data-stu-id="93b19-201">Type</span></span>| <span data-ttu-id="93b19-202">説明</span><span class="sxs-lookup"><span data-stu-id="93b19-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="93b19-203">String</span><span class="sxs-lookup"><span data-stu-id="93b19-203">String</span></span>|<span data-ttu-id="93b19-204">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="93b19-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="93b19-205">String</span><span class="sxs-lookup"><span data-stu-id="93b19-205">String</span></span>|<span data-ttu-id="93b19-206">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="93b19-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93b19-207">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-207">Requirements</span></span>

|<span data-ttu-id="93b19-208">要件</span><span class="sxs-lookup"><span data-stu-id="93b19-208">Requirement</span></span>| <span data-ttu-id="93b19-209">値</span><span class="sxs-lookup"><span data-stu-id="93b19-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="93b19-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="93b19-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93b19-211">1.0</span><span class="sxs-lookup"><span data-stu-id="93b19-211">1.0</span></span>|
|[<span data-ttu-id="93b19-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="93b19-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="93b19-213">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="93b19-213">Compose or read</span></span>|