
# <a name="userprofile"></a><span data-ttu-id="d23de-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="d23de-101">userProfile</span></span>

### <span data-ttu-id="d23de-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="d23de-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d23de-104">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-104">Requirements</span></span>

|<span data-ttu-id="d23de-105">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-105">Requirement</span></span>| <span data-ttu-id="d23de-106">値</span><span class="sxs-lookup"><span data-stu-id="d23de-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d23de-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d23de-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d23de-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d23de-108">1.0</span></span>|
|[<span data-ttu-id="d23de-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d23de-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d23de-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d23de-110">ReadItem</span></span>|
|[<span data-ttu-id="d23de-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d23de-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d23de-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d23de-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d23de-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d23de-113">Members and methods</span></span>

| <span data-ttu-id="d23de-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="d23de-114">Member</span></span> | <span data-ttu-id="d23de-115">種類</span><span class="sxs-lookup"><span data-stu-id="d23de-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d23de-116">accountType</span><span class="sxs-lookup"><span data-stu-id="d23de-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="d23de-117">Member</span><span class="sxs-lookup"><span data-stu-id="d23de-117">Member</span></span> |
| [<span data-ttu-id="d23de-118">displayName</span><span class="sxs-lookup"><span data-stu-id="d23de-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="d23de-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="d23de-119">Member</span></span> |
| [<span data-ttu-id="d23de-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d23de-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d23de-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="d23de-121">Member</span></span> |
| [<span data-ttu-id="d23de-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="d23de-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d23de-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="d23de-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d23de-124">Members</span><span class="sxs-lookup"><span data-stu-id="d23de-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="d23de-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="d23de-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="d23de-126">このメンバーは、現在のみ 2016 の Outlook でサポートされている Mac の後で (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="d23de-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="d23de-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="d23de-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="d23de-128">使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="d23de-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="d23de-129">値</span><span class="sxs-lookup"><span data-stu-id="d23de-129">Value</span></span> | <span data-ttu-id="d23de-130">説明</span><span class="sxs-lookup"><span data-stu-id="d23de-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="d23de-131">メールボックスは、オンプレミスの Exchange サーバーには。</span><span class="sxs-lookup"><span data-stu-id="d23de-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="d23de-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="d23de-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="d23de-133">メールボックスが関連付けられている、Office 365 の機能や、学校のアカウントです。</span><span class="sxs-lookup"><span data-stu-id="d23de-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="d23de-134">メールボックスは、個人、Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="d23de-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="d23de-135">型:</span><span class="sxs-lookup"><span data-stu-id="d23de-135">Type:</span></span>

*   <span data-ttu-id="d23de-136">String</span><span class="sxs-lookup"><span data-stu-id="d23de-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d23de-137">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-137">Requirements</span></span>

|<span data-ttu-id="d23de-138">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-138">Requirement</span></span>| <span data-ttu-id="d23de-139">値</span><span class="sxs-lookup"><span data-stu-id="d23de-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d23de-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d23de-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d23de-141">1.6</span><span class="sxs-lookup"><span data-stu-id="d23de-141">1.6</span></span> |
|[<span data-ttu-id="d23de-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d23de-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d23de-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d23de-143">ReadItem</span></span>|
|[<span data-ttu-id="d23de-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d23de-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d23de-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d23de-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d23de-146">例</span><span class="sxs-lookup"><span data-stu-id="d23de-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="d23de-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d23de-147">displayName :String</span></span>

<span data-ttu-id="d23de-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d23de-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d23de-149">型:</span><span class="sxs-lookup"><span data-stu-id="d23de-149">Type:</span></span>

*   <span data-ttu-id="d23de-150">String</span><span class="sxs-lookup"><span data-stu-id="d23de-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d23de-151">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-151">Requirements</span></span>

|<span data-ttu-id="d23de-152">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-152">Requirement</span></span>| <span data-ttu-id="d23de-153">値</span><span class="sxs-lookup"><span data-stu-id="d23de-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="d23de-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d23de-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d23de-155">1.0</span><span class="sxs-lookup"><span data-stu-id="d23de-155">1.0</span></span>|
|[<span data-ttu-id="d23de-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d23de-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d23de-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d23de-157">ReadItem</span></span>|
|[<span data-ttu-id="d23de-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d23de-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d23de-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d23de-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d23de-160">例</span><span class="sxs-lookup"><span data-stu-id="d23de-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d23de-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d23de-161">emailAddress :String</span></span>

<span data-ttu-id="d23de-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="d23de-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d23de-163">型:</span><span class="sxs-lookup"><span data-stu-id="d23de-163">Type:</span></span>

*   <span data-ttu-id="d23de-164">String</span><span class="sxs-lookup"><span data-stu-id="d23de-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d23de-165">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-165">Requirements</span></span>

|<span data-ttu-id="d23de-166">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-166">Requirement</span></span>| <span data-ttu-id="d23de-167">値</span><span class="sxs-lookup"><span data-stu-id="d23de-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d23de-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d23de-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d23de-169">1.0</span><span class="sxs-lookup"><span data-stu-id="d23de-169">1.0</span></span>|
|[<span data-ttu-id="d23de-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d23de-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d23de-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d23de-171">ReadItem</span></span>|
|[<span data-ttu-id="d23de-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d23de-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d23de-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d23de-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d23de-174">例</span><span class="sxs-lookup"><span data-stu-id="d23de-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d23de-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d23de-175">timeZone :String</span></span>

<span data-ttu-id="d23de-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d23de-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d23de-177">型:</span><span class="sxs-lookup"><span data-stu-id="d23de-177">Type:</span></span>

*   <span data-ttu-id="d23de-178">String</span><span class="sxs-lookup"><span data-stu-id="d23de-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d23de-179">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-179">Requirements</span></span>

|<span data-ttu-id="d23de-180">要件</span><span class="sxs-lookup"><span data-stu-id="d23de-180">Requirement</span></span>| <span data-ttu-id="d23de-181">値</span><span class="sxs-lookup"><span data-stu-id="d23de-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d23de-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d23de-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d23de-183">1.0</span><span class="sxs-lookup"><span data-stu-id="d23de-183">1.0</span></span>|
|[<span data-ttu-id="d23de-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d23de-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d23de-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d23de-185">ReadItem</span></span>|
|[<span data-ttu-id="d23de-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d23de-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d23de-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d23de-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d23de-188">例</span><span class="sxs-lookup"><span data-stu-id="d23de-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```