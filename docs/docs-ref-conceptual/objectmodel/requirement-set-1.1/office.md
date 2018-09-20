 

# <a name="office"></a><span data-ttu-id="b2e87-101">Office</span><span class="sxs-lookup"><span data-stu-id="b2e87-101">Office</span></span>

<span data-ttu-id="b2e87-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2e87-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2e87-104">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-104">Requirements</span></span>

|<span data-ttu-id="b2e87-105">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-105">Requirement</span></span>| <span data-ttu-id="b2e87-106">値</span><span class="sxs-lookup"><span data-stu-id="b2e87-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2e87-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b2e87-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2e87-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b2e87-108">1.0</span></span>|
|[<span data-ttu-id="b2e87-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b2e87-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2e87-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b2e87-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="b2e87-111">名前空間</span><span class="sxs-lookup"><span data-stu-id="b2e87-111">Namespaces</span></span>

<span data-ttu-id="b2e87-112">[コンテキスト](office.context.md): Outlook アドイン API で使用するための Office アドイン API のコンテキストの名前空間からの共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b2e87-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b2e87-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b2e87-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="b2e87-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b2e87-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b2e87-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="b2e87-116">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b2e87-117">型:</span><span class="sxs-lookup"><span data-stu-id="b2e87-117">Type:</span></span>

*   <span data-ttu-id="b2e87-118">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b2e87-119">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b2e87-119">Properties:</span></span>

|<span data-ttu-id="b2e87-120">名前</span><span class="sxs-lookup"><span data-stu-id="b2e87-120">Name</span></span>| <span data-ttu-id="b2e87-121">種類</span><span class="sxs-lookup"><span data-stu-id="b2e87-121">Type</span></span>| <span data-ttu-id="b2e87-122">説明</span><span class="sxs-lookup"><span data-stu-id="b2e87-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b2e87-123">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-123">String</span></span>|<span data-ttu-id="b2e87-124">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="b2e87-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b2e87-125">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-125">String</span></span>|<span data-ttu-id="b2e87-126">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="b2e87-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b2e87-127">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-127">Requirements</span></span>

|<span data-ttu-id="b2e87-128">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-128">Requirement</span></span>| <span data-ttu-id="b2e87-129">値</span><span class="sxs-lookup"><span data-stu-id="b2e87-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2e87-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b2e87-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2e87-131">1.0</span><span class="sxs-lookup"><span data-stu-id="b2e87-131">1.0</span></span>|
|[<span data-ttu-id="b2e87-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b2e87-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2e87-133">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b2e87-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="b2e87-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b2e87-134">CoercionType :String</span></span>

<span data-ttu-id="b2e87-135">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b2e87-136">型:</span><span class="sxs-lookup"><span data-stu-id="b2e87-136">Type:</span></span>

*   <span data-ttu-id="b2e87-137">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b2e87-138">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b2e87-138">Properties:</span></span>

|<span data-ttu-id="b2e87-139">名前</span><span class="sxs-lookup"><span data-stu-id="b2e87-139">Name</span></span>| <span data-ttu-id="b2e87-140">種類</span><span class="sxs-lookup"><span data-stu-id="b2e87-140">Type</span></span>| <span data-ttu-id="b2e87-141">説明</span><span class="sxs-lookup"><span data-stu-id="b2e87-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b2e87-142">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-142">String</span></span>|<span data-ttu-id="b2e87-143">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b2e87-144">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-144">String</span></span>|<span data-ttu-id="b2e87-145">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b2e87-146">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-146">Requirements</span></span>

|<span data-ttu-id="b2e87-147">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-147">Requirement</span></span>| <span data-ttu-id="b2e87-148">値</span><span class="sxs-lookup"><span data-stu-id="b2e87-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2e87-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b2e87-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2e87-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b2e87-150">1.0</span></span>|
|[<span data-ttu-id="b2e87-151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b2e87-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2e87-152">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b2e87-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="b2e87-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b2e87-153">SourceProperty :String</span></span>

<span data-ttu-id="b2e87-154">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="b2e87-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b2e87-155">型:</span><span class="sxs-lookup"><span data-stu-id="b2e87-155">Type:</span></span>

*   <span data-ttu-id="b2e87-156">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b2e87-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b2e87-157">Properties:</span></span>

|<span data-ttu-id="b2e87-158">名前</span><span class="sxs-lookup"><span data-stu-id="b2e87-158">Name</span></span>| <span data-ttu-id="b2e87-159">種類</span><span class="sxs-lookup"><span data-stu-id="b2e87-159">Type</span></span>| <span data-ttu-id="b2e87-160">説明</span><span class="sxs-lookup"><span data-stu-id="b2e87-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b2e87-161">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-161">String</span></span>|<span data-ttu-id="b2e87-162">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="b2e87-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b2e87-163">String</span><span class="sxs-lookup"><span data-stu-id="b2e87-163">String</span></span>|<span data-ttu-id="b2e87-164">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="b2e87-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b2e87-165">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-165">Requirements</span></span>

|<span data-ttu-id="b2e87-166">要件</span><span class="sxs-lookup"><span data-stu-id="b2e87-166">Requirement</span></span>| <span data-ttu-id="b2e87-167">値</span><span class="sxs-lookup"><span data-stu-id="b2e87-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2e87-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b2e87-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2e87-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b2e87-169">1.0</span></span>|
|[<span data-ttu-id="b2e87-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b2e87-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2e87-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="b2e87-171">Compose or read</span></span>|