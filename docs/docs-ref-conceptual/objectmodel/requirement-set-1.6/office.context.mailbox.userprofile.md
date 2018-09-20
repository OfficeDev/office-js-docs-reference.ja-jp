
# <a name="userprofile"></a><span data-ttu-id="60eb2-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="60eb2-101">userProfile</span></span>

### <span data-ttu-id="60eb2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="60eb2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="60eb2-104">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-104">Requirements</span></span>

|<span data-ttu-id="60eb2-105">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-105">Requirement</span></span>| <span data-ttu-id="60eb2-106">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="60eb2-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60eb2-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60eb2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="60eb2-108">1.0</span></span>|
|[<span data-ttu-id="60eb2-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60eb2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60eb2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60eb2-110">ReadItem</span></span>|
|[<span data-ttu-id="60eb2-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60eb2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="60eb2-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="60eb2-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="60eb2-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="60eb2-113">Members and methods</span></span>

| <span data-ttu-id="60eb2-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="60eb2-114">Member</span></span> | <span data-ttu-id="60eb2-115">種類</span><span class="sxs-lookup"><span data-stu-id="60eb2-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="60eb2-116">accountType</span><span class="sxs-lookup"><span data-stu-id="60eb2-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="60eb2-117">Member</span><span class="sxs-lookup"><span data-stu-id="60eb2-117">Member</span></span> |
| [<span data-ttu-id="60eb2-118">displayName</span><span class="sxs-lookup"><span data-stu-id="60eb2-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="60eb2-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="60eb2-119">Member</span></span> |
| [<span data-ttu-id="60eb2-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="60eb2-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="60eb2-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="60eb2-121">Member</span></span> |
| [<span data-ttu-id="60eb2-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="60eb2-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="60eb2-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="60eb2-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="60eb2-124">Members</span><span class="sxs-lookup"><span data-stu-id="60eb2-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="60eb2-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="60eb2-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="60eb2-126">このメンバーは、現在、Mac 用のみでサポートされている Outlook の 2016年、16.9.1212 を構築し、大きいです。</span><span class="sxs-lookup"><span data-stu-id="60eb2-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="60eb2-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="60eb2-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="60eb2-128">使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="60eb2-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="60eb2-129">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-129">Value</span></span> | <span data-ttu-id="60eb2-130">説明</span><span class="sxs-lookup"><span data-stu-id="60eb2-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="60eb2-131">メールボックスは、オンプレミスの Exchange サーバーには。</span><span class="sxs-lookup"><span data-stu-id="60eb2-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="60eb2-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="60eb2-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="60eb2-133">メールボックスが関連付けられている、Office 365 の機能や、学校のアカウントです。</span><span class="sxs-lookup"><span data-stu-id="60eb2-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="60eb2-134">メールボックスは、個人、Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="60eb2-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="60eb2-135">型:</span><span class="sxs-lookup"><span data-stu-id="60eb2-135">Type:</span></span>

*   <span data-ttu-id="60eb2-136">String</span><span class="sxs-lookup"><span data-stu-id="60eb2-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60eb2-137">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-137">Requirements</span></span>

|<span data-ttu-id="60eb2-138">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-138">Requirement</span></span>| <span data-ttu-id="60eb2-139">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="60eb2-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60eb2-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60eb2-141">1.6</span><span class="sxs-lookup"><span data-stu-id="60eb2-141">1.6</span></span> |
|[<span data-ttu-id="60eb2-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60eb2-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60eb2-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60eb2-143">ReadItem</span></span>|
|[<span data-ttu-id="60eb2-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60eb2-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="60eb2-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="60eb2-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="60eb2-146">例</span><span class="sxs-lookup"><span data-stu-id="60eb2-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="60eb2-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="60eb2-147">displayName :String</span></span>

<span data-ttu-id="60eb2-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="60eb2-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="60eb2-149">型:</span><span class="sxs-lookup"><span data-stu-id="60eb2-149">Type:</span></span>

*   <span data-ttu-id="60eb2-150">String</span><span class="sxs-lookup"><span data-stu-id="60eb2-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60eb2-151">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-151">Requirements</span></span>

|<span data-ttu-id="60eb2-152">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-152">Requirement</span></span>| <span data-ttu-id="60eb2-153">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="60eb2-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60eb2-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60eb2-155">1.0</span><span class="sxs-lookup"><span data-stu-id="60eb2-155">1.0</span></span>|
|[<span data-ttu-id="60eb2-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60eb2-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60eb2-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60eb2-157">ReadItem</span></span>|
|[<span data-ttu-id="60eb2-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60eb2-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="60eb2-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="60eb2-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="60eb2-160">例</span><span class="sxs-lookup"><span data-stu-id="60eb2-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="60eb2-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="60eb2-161">emailAddress :String</span></span>

<span data-ttu-id="60eb2-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="60eb2-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="60eb2-163">型:</span><span class="sxs-lookup"><span data-stu-id="60eb2-163">Type:</span></span>

*   <span data-ttu-id="60eb2-164">String</span><span class="sxs-lookup"><span data-stu-id="60eb2-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60eb2-165">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-165">Requirements</span></span>

|<span data-ttu-id="60eb2-166">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-166">Requirement</span></span>| <span data-ttu-id="60eb2-167">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="60eb2-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60eb2-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60eb2-169">1.0</span><span class="sxs-lookup"><span data-stu-id="60eb2-169">1.0</span></span>|
|[<span data-ttu-id="60eb2-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60eb2-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60eb2-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60eb2-171">ReadItem</span></span>|
|[<span data-ttu-id="60eb2-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60eb2-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="60eb2-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="60eb2-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="60eb2-174">例</span><span class="sxs-lookup"><span data-stu-id="60eb2-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="60eb2-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="60eb2-175">timeZone :String</span></span>

<span data-ttu-id="60eb2-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="60eb2-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="60eb2-177">型:</span><span class="sxs-lookup"><span data-stu-id="60eb2-177">Type:</span></span>

*   <span data-ttu-id="60eb2-178">String</span><span class="sxs-lookup"><span data-stu-id="60eb2-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60eb2-179">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-179">Requirements</span></span>

|<span data-ttu-id="60eb2-180">要件</span><span class="sxs-lookup"><span data-stu-id="60eb2-180">Requirement</span></span>| <span data-ttu-id="60eb2-181">値</span><span class="sxs-lookup"><span data-stu-id="60eb2-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="60eb2-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60eb2-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60eb2-183">1.0</span><span class="sxs-lookup"><span data-stu-id="60eb2-183">1.0</span></span>|
|[<span data-ttu-id="60eb2-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60eb2-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60eb2-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60eb2-185">ReadItem</span></span>|
|[<span data-ttu-id="60eb2-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60eb2-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="60eb2-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="60eb2-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="60eb2-188">例</span><span class="sxs-lookup"><span data-stu-id="60eb2-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```