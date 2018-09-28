# <a name="userprofile"></a><span data-ttu-id="83427-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="83427-101">userProfile</span></span>

### <span data-ttu-id="83427-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="83427-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="83427-104">要件</span><span class="sxs-lookup"><span data-stu-id="83427-104">Requirements</span></span>

|<span data-ttu-id="83427-105">要件</span><span class="sxs-lookup"><span data-stu-id="83427-105">Requirement</span></span>| <span data-ttu-id="83427-106">値</span><span class="sxs-lookup"><span data-stu-id="83427-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="83427-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83427-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83427-108">1.0</span><span class="sxs-lookup"><span data-stu-id="83427-108">1.0</span></span>|
|[<span data-ttu-id="83427-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="83427-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83427-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83427-110">ReadItem</span></span>|
|[<span data-ttu-id="83427-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83427-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83427-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="83427-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="83427-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="83427-113">Members and methods</span></span>

| <span data-ttu-id="83427-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="83427-114">Member</span></span> | <span data-ttu-id="83427-115">種類</span><span class="sxs-lookup"><span data-stu-id="83427-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="83427-116">displayName</span><span class="sxs-lookup"><span data-stu-id="83427-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="83427-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="83427-117">Member</span></span> |
| [<span data-ttu-id="83427-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="83427-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="83427-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="83427-119">Member</span></span> |
| [<span data-ttu-id="83427-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="83427-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="83427-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="83427-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="83427-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="83427-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="83427-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="83427-123">displayName :String</span></span>

<span data-ttu-id="83427-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="83427-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="83427-125">型:</span><span class="sxs-lookup"><span data-stu-id="83427-125">Type:</span></span>

*   <span data-ttu-id="83427-126">String</span><span class="sxs-lookup"><span data-stu-id="83427-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83427-127">要件</span><span class="sxs-lookup"><span data-stu-id="83427-127">Requirements</span></span>

|<span data-ttu-id="83427-128">要件</span><span class="sxs-lookup"><span data-stu-id="83427-128">Requirement</span></span>| <span data-ttu-id="83427-129">値</span><span class="sxs-lookup"><span data-stu-id="83427-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="83427-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83427-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83427-131">1.0</span><span class="sxs-lookup"><span data-stu-id="83427-131">1.0</span></span>|
|[<span data-ttu-id="83427-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="83427-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83427-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83427-133">ReadItem</span></span>|
|[<span data-ttu-id="83427-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83427-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83427-135">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="83427-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83427-136">例</span><span class="sxs-lookup"><span data-stu-id="83427-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="83427-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="83427-137">emailAddress :String</span></span>

<span data-ttu-id="83427-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="83427-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="83427-139">型:</span><span class="sxs-lookup"><span data-stu-id="83427-139">Type:</span></span>

*   <span data-ttu-id="83427-140">String</span><span class="sxs-lookup"><span data-stu-id="83427-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83427-141">要件</span><span class="sxs-lookup"><span data-stu-id="83427-141">Requirements</span></span>

|<span data-ttu-id="83427-142">要件</span><span class="sxs-lookup"><span data-stu-id="83427-142">Requirement</span></span>| <span data-ttu-id="83427-143">値</span><span class="sxs-lookup"><span data-stu-id="83427-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="83427-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83427-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83427-145">1.0</span><span class="sxs-lookup"><span data-stu-id="83427-145">1.0</span></span>|
|[<span data-ttu-id="83427-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="83427-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83427-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83427-147">ReadItem</span></span>|
|[<span data-ttu-id="83427-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83427-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83427-149">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="83427-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83427-150">例</span><span class="sxs-lookup"><span data-stu-id="83427-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="83427-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="83427-151">timeZone :String</span></span>

<span data-ttu-id="83427-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="83427-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="83427-153">型:</span><span class="sxs-lookup"><span data-stu-id="83427-153">Type:</span></span>

*   <span data-ttu-id="83427-154">String</span><span class="sxs-lookup"><span data-stu-id="83427-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83427-155">要件</span><span class="sxs-lookup"><span data-stu-id="83427-155">Requirements</span></span>

|<span data-ttu-id="83427-156">要件</span><span class="sxs-lookup"><span data-stu-id="83427-156">Requirement</span></span>| <span data-ttu-id="83427-157">値</span><span class="sxs-lookup"><span data-stu-id="83427-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="83427-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83427-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83427-159">1.0</span><span class="sxs-lookup"><span data-stu-id="83427-159">1.0</span></span>|
|[<span data-ttu-id="83427-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="83427-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83427-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83427-161">ReadItem</span></span>|
|[<span data-ttu-id="83427-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83427-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83427-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="83427-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83427-164">例</span><span class="sxs-lookup"><span data-stu-id="83427-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```