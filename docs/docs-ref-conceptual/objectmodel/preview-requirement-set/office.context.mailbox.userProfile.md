
# <a name="userprofile"></a><span data-ttu-id="5dba6-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="5dba6-101">userProfile</span></span>

### <span data-ttu-id="5dba6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="5dba6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dba6-104">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-104">Requirements</span></span>

|<span data-ttu-id="5dba6-105">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-105">Requirement</span></span>| <span data-ttu-id="5dba6-106">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dba6-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5dba6-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dba6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5dba6-108">1.0</span></span>|
|[<span data-ttu-id="5dba6-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5dba6-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5dba6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5dba6-110">ReadItem</span></span>|
|[<span data-ttu-id="5dba6-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5dba6-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5dba6-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5dba6-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5dba6-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="5dba6-113">Members and methods</span></span>

| <span data-ttu-id="5dba6-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="5dba6-114">Member</span></span> | <span data-ttu-id="5dba6-115">種類</span><span class="sxs-lookup"><span data-stu-id="5dba6-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5dba6-116">accountType</span><span class="sxs-lookup"><span data-stu-id="5dba6-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="5dba6-117">Member</span><span class="sxs-lookup"><span data-stu-id="5dba6-117">Member</span></span> |
| [<span data-ttu-id="5dba6-118">displayName</span><span class="sxs-lookup"><span data-stu-id="5dba6-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="5dba6-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="5dba6-119">Member</span></span> |
| [<span data-ttu-id="5dba6-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5dba6-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5dba6-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="5dba6-121">Member</span></span> |
| [<span data-ttu-id="5dba6-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="5dba6-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5dba6-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="5dba6-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5dba6-124">Members</span><span class="sxs-lookup"><span data-stu-id="5dba6-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="5dba6-125">accountType: 文字列</span><span class="sxs-lookup"><span data-stu-id="5dba6-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="5dba6-126">このメンバーは、現在のみ 2016 の Outlook でサポートされている Mac の後で (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="5dba6-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="5dba6-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="5dba6-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="5dba6-128">使用可能な値は、次の表に表示されます。</span><span class="sxs-lookup"><span data-stu-id="5dba6-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="5dba6-129">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-129">Value</span></span> | <span data-ttu-id="5dba6-130">説明</span><span class="sxs-lookup"><span data-stu-id="5dba6-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="5dba6-131">メールボックスは、オンプレミスの Exchange サーバーには。</span><span class="sxs-lookup"><span data-stu-id="5dba6-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="5dba6-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="5dba6-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="5dba6-133">メールボックスが関連付けられている、Office 365 の機能や、学校のアカウントです。</span><span class="sxs-lookup"><span data-stu-id="5dba6-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="5dba6-134">メールボックスは、個人、Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="5dba6-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="5dba6-135">型:</span><span class="sxs-lookup"><span data-stu-id="5dba6-135">Type:</span></span>

*   <span data-ttu-id="5dba6-136">String</span><span class="sxs-lookup"><span data-stu-id="5dba6-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dba6-137">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-137">Requirements</span></span>

|<span data-ttu-id="5dba6-138">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-138">Requirement</span></span>| <span data-ttu-id="5dba6-139">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dba6-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5dba6-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dba6-141">1.6</span><span class="sxs-lookup"><span data-stu-id="5dba6-141">1.6</span></span> |
|[<span data-ttu-id="5dba6-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5dba6-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5dba6-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5dba6-143">ReadItem</span></span>|
|[<span data-ttu-id="5dba6-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5dba6-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5dba6-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5dba6-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5dba6-146">例</span><span class="sxs-lookup"><span data-stu-id="5dba6-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="5dba6-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5dba6-147">displayName :String</span></span>

<span data-ttu-id="5dba6-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="5dba6-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5dba6-149">型:</span><span class="sxs-lookup"><span data-stu-id="5dba6-149">Type:</span></span>

*   <span data-ttu-id="5dba6-150">String</span><span class="sxs-lookup"><span data-stu-id="5dba6-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dba6-151">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-151">Requirements</span></span>

|<span data-ttu-id="5dba6-152">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-152">Requirement</span></span>| <span data-ttu-id="5dba6-153">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dba6-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5dba6-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dba6-155">1.0</span><span class="sxs-lookup"><span data-stu-id="5dba6-155">1.0</span></span>|
|[<span data-ttu-id="5dba6-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5dba6-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5dba6-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5dba6-157">ReadItem</span></span>|
|[<span data-ttu-id="5dba6-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5dba6-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5dba6-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5dba6-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5dba6-160">例</span><span class="sxs-lookup"><span data-stu-id="5dba6-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5dba6-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5dba6-161">emailAddress :String</span></span>

<span data-ttu-id="5dba6-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="5dba6-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5dba6-163">型:</span><span class="sxs-lookup"><span data-stu-id="5dba6-163">Type:</span></span>

*   <span data-ttu-id="5dba6-164">String</span><span class="sxs-lookup"><span data-stu-id="5dba6-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dba6-165">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-165">Requirements</span></span>

|<span data-ttu-id="5dba6-166">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-166">Requirement</span></span>| <span data-ttu-id="5dba6-167">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dba6-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5dba6-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dba6-169">1.0</span><span class="sxs-lookup"><span data-stu-id="5dba6-169">1.0</span></span>|
|[<span data-ttu-id="5dba6-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5dba6-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5dba6-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5dba6-171">ReadItem</span></span>|
|[<span data-ttu-id="5dba6-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5dba6-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5dba6-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5dba6-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5dba6-174">例</span><span class="sxs-lookup"><span data-stu-id="5dba6-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5dba6-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5dba6-175">timeZone :String</span></span>

<span data-ttu-id="5dba6-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="5dba6-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5dba6-177">型:</span><span class="sxs-lookup"><span data-stu-id="5dba6-177">Type:</span></span>

*   <span data-ttu-id="5dba6-178">String</span><span class="sxs-lookup"><span data-stu-id="5dba6-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dba6-179">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-179">Requirements</span></span>

|<span data-ttu-id="5dba6-180">要件</span><span class="sxs-lookup"><span data-stu-id="5dba6-180">Requirement</span></span>| <span data-ttu-id="5dba6-181">値</span><span class="sxs-lookup"><span data-stu-id="5dba6-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dba6-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5dba6-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dba6-183">1.0</span><span class="sxs-lookup"><span data-stu-id="5dba6-183">1.0</span></span>|
|[<span data-ttu-id="5dba6-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5dba6-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5dba6-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5dba6-185">ReadItem</span></span>|
|[<span data-ttu-id="5dba6-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5dba6-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5dba6-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5dba6-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5dba6-188">例</span><span class="sxs-lookup"><span data-stu-id="5dba6-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```