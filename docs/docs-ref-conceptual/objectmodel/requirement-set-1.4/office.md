 

# <a name="office"></a><span data-ttu-id="01db7-101">Office</span><span class="sxs-lookup"><span data-stu-id="01db7-101">Office</span></span>

<span data-ttu-id="01db7-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="01db7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="01db7-104">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-104">Requirements</span></span>

|<span data-ttu-id="01db7-105">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-105">Requirement</span></span>| <span data-ttu-id="01db7-106">値</span><span class="sxs-lookup"><span data-stu-id="01db7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="01db7-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01db7-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01db7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="01db7-108">1.0</span></span>|
|[<span data-ttu-id="01db7-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01db7-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01db7-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="01db7-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="01db7-111">名前空間</span><span class="sxs-lookup"><span data-stu-id="01db7-111">Namespaces</span></span>

<span data-ttu-id="01db7-112">[コンテキスト](Office.context.md): Outlook アドイン API で使用するための Office アドイン API のコンテキストの名前空間からの共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="01db7-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="01db7-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="01db7-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="01db7-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="01db7-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="01db7-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="01db7-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="01db7-116">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="01db7-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="01db7-117">型:</span><span class="sxs-lookup"><span data-stu-id="01db7-117">Type:</span></span>

*   <span data-ttu-id="01db7-118">String</span><span class="sxs-lookup"><span data-stu-id="01db7-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01db7-119">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="01db7-119">Properties:</span></span>

|<span data-ttu-id="01db7-120">名前</span><span class="sxs-lookup"><span data-stu-id="01db7-120">Name</span></span>| <span data-ttu-id="01db7-121">種類</span><span class="sxs-lookup"><span data-stu-id="01db7-121">Type</span></span>| <span data-ttu-id="01db7-122">説明</span><span class="sxs-lookup"><span data-stu-id="01db7-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="01db7-123">String</span><span class="sxs-lookup"><span data-stu-id="01db7-123">String</span></span>|<span data-ttu-id="01db7-124">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="01db7-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="01db7-125">String</span><span class="sxs-lookup"><span data-stu-id="01db7-125">String</span></span>|<span data-ttu-id="01db7-126">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="01db7-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01db7-127">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-127">Requirements</span></span>

|<span data-ttu-id="01db7-128">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-128">Requirement</span></span>| <span data-ttu-id="01db7-129">値</span><span class="sxs-lookup"><span data-stu-id="01db7-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="01db7-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01db7-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01db7-131">1.0</span><span class="sxs-lookup"><span data-stu-id="01db7-131">1.0</span></span>|
|[<span data-ttu-id="01db7-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01db7-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01db7-133">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="01db7-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="01db7-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="01db7-134">CoercionType :String</span></span>

<span data-ttu-id="01db7-135">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="01db7-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="01db7-136">型:</span><span class="sxs-lookup"><span data-stu-id="01db7-136">Type:</span></span>

*   <span data-ttu-id="01db7-137">String</span><span class="sxs-lookup"><span data-stu-id="01db7-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01db7-138">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="01db7-138">Properties:</span></span>

|<span data-ttu-id="01db7-139">名前</span><span class="sxs-lookup"><span data-stu-id="01db7-139">Name</span></span>| <span data-ttu-id="01db7-140">種類</span><span class="sxs-lookup"><span data-stu-id="01db7-140">Type</span></span>| <span data-ttu-id="01db7-141">説明</span><span class="sxs-lookup"><span data-stu-id="01db7-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="01db7-142">String</span><span class="sxs-lookup"><span data-stu-id="01db7-142">String</span></span>|<span data-ttu-id="01db7-143">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="01db7-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="01db7-144">String</span><span class="sxs-lookup"><span data-stu-id="01db7-144">String</span></span>|<span data-ttu-id="01db7-145">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="01db7-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01db7-146">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-146">Requirements</span></span>

|<span data-ttu-id="01db7-147">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-147">Requirement</span></span>| <span data-ttu-id="01db7-148">値</span><span class="sxs-lookup"><span data-stu-id="01db7-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="01db7-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01db7-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01db7-150">1.0</span><span class="sxs-lookup"><span data-stu-id="01db7-150">1.0</span></span>|
|[<span data-ttu-id="01db7-151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01db7-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01db7-152">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="01db7-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="01db7-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="01db7-153">SourceProperty :String</span></span>

<span data-ttu-id="01db7-154">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="01db7-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="01db7-155">型:</span><span class="sxs-lookup"><span data-stu-id="01db7-155">Type:</span></span>

*   <span data-ttu-id="01db7-156">String</span><span class="sxs-lookup"><span data-stu-id="01db7-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01db7-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="01db7-157">Properties:</span></span>

|<span data-ttu-id="01db7-158">名前</span><span class="sxs-lookup"><span data-stu-id="01db7-158">Name</span></span>| <span data-ttu-id="01db7-159">種類</span><span class="sxs-lookup"><span data-stu-id="01db7-159">Type</span></span>| <span data-ttu-id="01db7-160">説明</span><span class="sxs-lookup"><span data-stu-id="01db7-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="01db7-161">String</span><span class="sxs-lookup"><span data-stu-id="01db7-161">String</span></span>|<span data-ttu-id="01db7-162">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="01db7-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="01db7-163">String</span><span class="sxs-lookup"><span data-stu-id="01db7-163">String</span></span>|<span data-ttu-id="01db7-164">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="01db7-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01db7-165">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-165">Requirements</span></span>

|<span data-ttu-id="01db7-166">要件</span><span class="sxs-lookup"><span data-stu-id="01db7-166">Requirement</span></span>| <span data-ttu-id="01db7-167">値</span><span class="sxs-lookup"><span data-stu-id="01db7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="01db7-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01db7-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01db7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="01db7-169">1.0</span></span>|
|[<span data-ttu-id="01db7-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01db7-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01db7-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="01db7-171">Compose or read</span></span>|