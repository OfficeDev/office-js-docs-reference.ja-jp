
# <a name="diagnostics"></a><span data-ttu-id="1f21e-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="1f21e-101">diagnostics</span></span>

### <span data-ttu-id="1f21e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="1f21e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="1f21e-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f21e-105">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-105">Requirements</span></span>

|<span data-ttu-id="1f21e-106">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-106">Requirement</span></span>| <span data-ttu-id="1f21e-107">値</span><span class="sxs-lookup"><span data-stu-id="1f21e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f21e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f21e-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f21e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1f21e-109">1.0</span></span>|
|[<span data-ttu-id="1f21e-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1f21e-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f21e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f21e-111">ReadItem</span></span>|
|[<span data-ttu-id="1f21e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f21e-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f21e-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1f21e-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="1f21e-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="1f21e-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="1f21e-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="1f21e-115">hostName :String</span></span>

<span data-ttu-id="1f21e-116">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="1f21e-117">次の値のいずれかの文字列: `Outlook`、 `OutlookIOS`、または`OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="1f21e-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="1f21e-118">型:</span><span class="sxs-lookup"><span data-stu-id="1f21e-118">Type:</span></span>

*   <span data-ttu-id="1f21e-119">String</span><span class="sxs-lookup"><span data-stu-id="1f21e-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f21e-120">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-120">Requirements</span></span>

|<span data-ttu-id="1f21e-121">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-121">Requirement</span></span>| <span data-ttu-id="1f21e-122">値</span><span class="sxs-lookup"><span data-stu-id="1f21e-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f21e-123">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f21e-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f21e-124">1.0</span><span class="sxs-lookup"><span data-stu-id="1f21e-124">1.0</span></span>|
|[<span data-ttu-id="1f21e-125">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1f21e-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f21e-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f21e-126">ReadItem</span></span>|
|[<span data-ttu-id="1f21e-127">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f21e-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f21e-128">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1f21e-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="1f21e-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="1f21e-129">hostVersion :String</span></span>

<span data-ttu-id="1f21e-130">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="1f21e-p102">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="1f21e-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="1f21e-134">種類:</span><span class="sxs-lookup"><span data-stu-id="1f21e-134">Type:</span></span>

*   <span data-ttu-id="1f21e-135">String</span><span class="sxs-lookup"><span data-stu-id="1f21e-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f21e-136">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-136">Requirements</span></span>

|<span data-ttu-id="1f21e-137">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-137">Requirement</span></span>| <span data-ttu-id="1f21e-138">値</span><span class="sxs-lookup"><span data-stu-id="1f21e-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f21e-139">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f21e-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f21e-140">1.0</span><span class="sxs-lookup"><span data-stu-id="1f21e-140">1.0</span></span>|
|[<span data-ttu-id="1f21e-141">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1f21e-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f21e-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f21e-142">ReadItem</span></span>|
|[<span data-ttu-id="1f21e-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f21e-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f21e-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1f21e-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="1f21e-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="1f21e-145">OWAView :String</span></span>

<span data-ttu-id="1f21e-146">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="1f21e-147">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="1f21e-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="1f21e-148">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="1f21e-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="1f21e-149">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="1f21e-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="1f21e-p103">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="1f21e-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="1f21e-p104">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="1f21e-p105">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="1f21e-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="1f21e-156">型:</span><span class="sxs-lookup"><span data-stu-id="1f21e-156">Type:</span></span>

*   <span data-ttu-id="1f21e-157">String</span><span class="sxs-lookup"><span data-stu-id="1f21e-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f21e-158">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-158">Requirements</span></span>

|<span data-ttu-id="1f21e-159">要件</span><span class="sxs-lookup"><span data-stu-id="1f21e-159">Requirement</span></span>| <span data-ttu-id="1f21e-160">値</span><span class="sxs-lookup"><span data-stu-id="1f21e-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f21e-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f21e-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f21e-162">1.0</span><span class="sxs-lookup"><span data-stu-id="1f21e-162">1.0</span></span>|
|[<span data-ttu-id="1f21e-163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1f21e-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f21e-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f21e-164">ReadItem</span></span>|
|[<span data-ttu-id="1f21e-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f21e-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f21e-166">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="1f21e-166">Compose or read</span></span>|