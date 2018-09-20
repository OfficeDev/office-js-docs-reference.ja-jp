# <a name="userprofile"></a><span data-ttu-id="04c39-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="04c39-101">userProfile</span></span>

### <span data-ttu-id="04c39-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="04c39-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c39-104">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-104">Requirements</span></span>

|<span data-ttu-id="04c39-105">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-105">Requirement</span></span>| <span data-ttu-id="04c39-106">値</span><span class="sxs-lookup"><span data-stu-id="04c39-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c39-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04c39-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c39-108">1.0</span><span class="sxs-lookup"><span data-stu-id="04c39-108">1.0</span></span>|
|[<span data-ttu-id="04c39-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04c39-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c39-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c39-110">ReadItem</span></span>|
|[<span data-ttu-id="04c39-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04c39-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c39-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="04c39-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="04c39-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="04c39-113">Members and methods</span></span>

| <span data-ttu-id="04c39-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="04c39-114">Member</span></span> | <span data-ttu-id="04c39-115">種類</span><span class="sxs-lookup"><span data-stu-id="04c39-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="04c39-116">displayName</span><span class="sxs-lookup"><span data-stu-id="04c39-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="04c39-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="04c39-117">Member</span></span> |
| [<span data-ttu-id="04c39-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="04c39-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="04c39-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="04c39-119">Member</span></span> |
| [<span data-ttu-id="04c39-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="04c39-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="04c39-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="04c39-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="04c39-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="04c39-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="04c39-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="04c39-123">displayName :String</span></span>

<span data-ttu-id="04c39-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="04c39-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="04c39-125">型:</span><span class="sxs-lookup"><span data-stu-id="04c39-125">Type:</span></span>

*   <span data-ttu-id="04c39-126">String</span><span class="sxs-lookup"><span data-stu-id="04c39-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c39-127">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-127">Requirements</span></span>

|<span data-ttu-id="04c39-128">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-128">Requirement</span></span>| <span data-ttu-id="04c39-129">値</span><span class="sxs-lookup"><span data-stu-id="04c39-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c39-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04c39-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c39-131">1.0</span><span class="sxs-lookup"><span data-stu-id="04c39-131">1.0</span></span>|
|[<span data-ttu-id="04c39-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04c39-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c39-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c39-133">ReadItem</span></span>|
|[<span data-ttu-id="04c39-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04c39-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c39-135">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="04c39-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c39-136">例</span><span class="sxs-lookup"><span data-stu-id="04c39-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="04c39-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="04c39-137">emailAddress :String</span></span>

<span data-ttu-id="04c39-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="04c39-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="04c39-139">型:</span><span class="sxs-lookup"><span data-stu-id="04c39-139">Type:</span></span>

*   <span data-ttu-id="04c39-140">String</span><span class="sxs-lookup"><span data-stu-id="04c39-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c39-141">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-141">Requirements</span></span>

|<span data-ttu-id="04c39-142">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-142">Requirement</span></span>| <span data-ttu-id="04c39-143">値</span><span class="sxs-lookup"><span data-stu-id="04c39-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c39-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04c39-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c39-145">1.0</span><span class="sxs-lookup"><span data-stu-id="04c39-145">1.0</span></span>|
|[<span data-ttu-id="04c39-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04c39-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c39-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c39-147">ReadItem</span></span>|
|[<span data-ttu-id="04c39-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04c39-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c39-149">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="04c39-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c39-150">例</span><span class="sxs-lookup"><span data-stu-id="04c39-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="04c39-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="04c39-151">timeZone :String</span></span>

<span data-ttu-id="04c39-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="04c39-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="04c39-153">型:</span><span class="sxs-lookup"><span data-stu-id="04c39-153">Type:</span></span>

*   <span data-ttu-id="04c39-154">String</span><span class="sxs-lookup"><span data-stu-id="04c39-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04c39-155">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-155">Requirements</span></span>

|<span data-ttu-id="04c39-156">要件</span><span class="sxs-lookup"><span data-stu-id="04c39-156">Requirement</span></span>| <span data-ttu-id="04c39-157">値</span><span class="sxs-lookup"><span data-stu-id="04c39-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="04c39-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04c39-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04c39-159">1.0</span><span class="sxs-lookup"><span data-stu-id="04c39-159">1.0</span></span>|
|[<span data-ttu-id="04c39-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04c39-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04c39-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04c39-161">ReadItem</span></span>|
|[<span data-ttu-id="04c39-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04c39-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04c39-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="04c39-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="04c39-164">例</span><span class="sxs-lookup"><span data-stu-id="04c39-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```