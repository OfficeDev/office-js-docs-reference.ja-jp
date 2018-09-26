
# <a name="userprofile"></a><span data-ttu-id="138ef-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="138ef-101">userProfile</span></span>

### <span data-ttu-id="138ef-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="138ef-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="138ef-104">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-104">Requirements</span></span>

|<span data-ttu-id="138ef-105">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-105">Requirement</span></span>| <span data-ttu-id="138ef-106">値</span><span class="sxs-lookup"><span data-stu-id="138ef-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="138ef-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="138ef-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="138ef-108">1.0</span><span class="sxs-lookup"><span data-stu-id="138ef-108">1.0</span></span>|
|[<span data-ttu-id="138ef-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="138ef-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="138ef-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="138ef-110">ReadItem</span></span>|
|[<span data-ttu-id="138ef-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="138ef-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="138ef-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="138ef-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="138ef-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="138ef-113">Members and methods</span></span>

| <span data-ttu-id="138ef-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="138ef-114">Member</span></span> | <span data-ttu-id="138ef-115">種類</span><span class="sxs-lookup"><span data-stu-id="138ef-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="138ef-116">accountType</span><span class="sxs-lookup"><span data-stu-id="138ef-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="138ef-117">Member</span><span class="sxs-lookup"><span data-stu-id="138ef-117">Member</span></span> |
| [<span data-ttu-id="138ef-118">displayName</span><span class="sxs-lookup"><span data-stu-id="138ef-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="138ef-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="138ef-119">Member</span></span> |
| [<span data-ttu-id="138ef-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="138ef-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="138ef-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="138ef-121">Member</span></span> |
| [<span data-ttu-id="138ef-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="138ef-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="138ef-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="138ef-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="138ef-124">Members</span><span class="sxs-lookup"><span data-stu-id="138ef-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="138ef-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="138ef-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="138ef-126">このメンバーは、現在のみ 2016 の Outlook でサポートされている Mac の後で (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="138ef-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="138ef-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="138ef-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="138ef-128">使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="138ef-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="138ef-129">値</span><span class="sxs-lookup"><span data-stu-id="138ef-129">Value</span></span> | <span data-ttu-id="138ef-130">説明</span><span class="sxs-lookup"><span data-stu-id="138ef-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="138ef-131">メールボックスは、オンプレミスの Exchange サーバーには。</span><span class="sxs-lookup"><span data-stu-id="138ef-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="138ef-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="138ef-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="138ef-133">メールボックスが関連付けられている、Office 365 の機能や、学校のアカウントです。</span><span class="sxs-lookup"><span data-stu-id="138ef-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="138ef-134">メールボックスは、個人、Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="138ef-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="138ef-135">型:</span><span class="sxs-lookup"><span data-stu-id="138ef-135">Type:</span></span>

*   <span data-ttu-id="138ef-136">String</span><span class="sxs-lookup"><span data-stu-id="138ef-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="138ef-137">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-137">Requirements</span></span>

|<span data-ttu-id="138ef-138">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-138">Requirement</span></span>| <span data-ttu-id="138ef-139">値</span><span class="sxs-lookup"><span data-stu-id="138ef-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="138ef-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="138ef-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="138ef-141">1.6</span><span class="sxs-lookup"><span data-stu-id="138ef-141">1.6</span></span> |
|[<span data-ttu-id="138ef-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="138ef-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="138ef-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="138ef-143">ReadItem</span></span>|
|[<span data-ttu-id="138ef-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="138ef-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="138ef-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="138ef-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="138ef-146">例</span><span class="sxs-lookup"><span data-stu-id="138ef-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="138ef-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="138ef-147">displayName :String</span></span>

<span data-ttu-id="138ef-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="138ef-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="138ef-149">型:</span><span class="sxs-lookup"><span data-stu-id="138ef-149">Type:</span></span>

*   <span data-ttu-id="138ef-150">String</span><span class="sxs-lookup"><span data-stu-id="138ef-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="138ef-151">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-151">Requirements</span></span>

|<span data-ttu-id="138ef-152">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-152">Requirement</span></span>| <span data-ttu-id="138ef-153">値</span><span class="sxs-lookup"><span data-stu-id="138ef-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="138ef-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="138ef-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="138ef-155">1.0</span><span class="sxs-lookup"><span data-stu-id="138ef-155">1.0</span></span>|
|[<span data-ttu-id="138ef-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="138ef-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="138ef-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="138ef-157">ReadItem</span></span>|
|[<span data-ttu-id="138ef-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="138ef-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="138ef-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="138ef-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="138ef-160">例</span><span class="sxs-lookup"><span data-stu-id="138ef-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="138ef-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="138ef-161">emailAddress :String</span></span>

<span data-ttu-id="138ef-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="138ef-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="138ef-163">型:</span><span class="sxs-lookup"><span data-stu-id="138ef-163">Type:</span></span>

*   <span data-ttu-id="138ef-164">String</span><span class="sxs-lookup"><span data-stu-id="138ef-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="138ef-165">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-165">Requirements</span></span>

|<span data-ttu-id="138ef-166">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-166">Requirement</span></span>| <span data-ttu-id="138ef-167">値</span><span class="sxs-lookup"><span data-stu-id="138ef-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="138ef-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="138ef-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="138ef-169">1.0</span><span class="sxs-lookup"><span data-stu-id="138ef-169">1.0</span></span>|
|[<span data-ttu-id="138ef-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="138ef-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="138ef-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="138ef-171">ReadItem</span></span>|
|[<span data-ttu-id="138ef-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="138ef-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="138ef-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="138ef-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="138ef-174">例</span><span class="sxs-lookup"><span data-stu-id="138ef-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="138ef-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="138ef-175">timeZone :String</span></span>

<span data-ttu-id="138ef-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="138ef-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="138ef-177">型:</span><span class="sxs-lookup"><span data-stu-id="138ef-177">Type:</span></span>

*   <span data-ttu-id="138ef-178">String</span><span class="sxs-lookup"><span data-stu-id="138ef-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="138ef-179">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-179">Requirements</span></span>

|<span data-ttu-id="138ef-180">要件</span><span class="sxs-lookup"><span data-stu-id="138ef-180">Requirement</span></span>| <span data-ttu-id="138ef-181">値</span><span class="sxs-lookup"><span data-stu-id="138ef-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="138ef-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="138ef-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="138ef-183">1.0</span><span class="sxs-lookup"><span data-stu-id="138ef-183">1.0</span></span>|
|[<span data-ttu-id="138ef-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="138ef-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="138ef-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="138ef-185">ReadItem</span></span>|
|[<span data-ttu-id="138ef-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="138ef-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="138ef-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="138ef-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="138ef-188">例</span><span class="sxs-lookup"><span data-stu-id="138ef-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```