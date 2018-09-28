
# <a name="context"></a><span data-ttu-id="fb39f-101">context</span><span class="sxs-lookup"><span data-stu-id="fb39f-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="fb39f-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="fb39f-102">[Office](Office.md).context</span></span>

<span data-ttu-id="fb39f-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共有 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fb39f-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fb39f-105">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-105">Requirements</span></span>

|<span data-ttu-id="fb39f-106">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-106">Requirement</span></span>| <span data-ttu-id="fb39f-107">値</span><span class="sxs-lookup"><span data-stu-id="fb39f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb39f-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb39f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb39f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fb39f-109">1.0</span></span>|
|[<span data-ttu-id="fb39f-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb39f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fb39f-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fb39f-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="fb39f-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="fb39f-112">Namespaces</span></span>

<span data-ttu-id="fb39f-113">[メールボックス](office.context.mailbox.md): Microsoft Outlook と web 上の Microsoft Outlook で Outlook アドインのオブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="fb39f-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="fb39f-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="fb39f-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="fb39f-115">displayLanguage :String</span></span>

<span data-ttu-id="fb39f-116">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="fb39f-117">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="fb39f-118">型:</span><span class="sxs-lookup"><span data-stu-id="fb39f-118">Type:</span></span>

*   <span data-ttu-id="fb39f-119">String</span><span class="sxs-lookup"><span data-stu-id="fb39f-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fb39f-120">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-120">Requirements</span></span>

|<span data-ttu-id="fb39f-121">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-121">Requirement</span></span>| <span data-ttu-id="fb39f-122">値</span><span class="sxs-lookup"><span data-stu-id="fb39f-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb39f-123">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb39f-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb39f-124">1.0</span><span class="sxs-lookup"><span data-stu-id="fb39f-124">1.0</span></span>|
|[<span data-ttu-id="fb39f-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb39f-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fb39f-126">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fb39f-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fb39f-127">例</span><span class="sxs-lookup"><span data-stu-id="fb39f-127">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="fb39f-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="fb39f-128">officeTheme :Object</span></span>

<span data-ttu-id="fb39f-129">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="fb39f-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="fb39f-130">IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fb39f-130">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fb39f-p102">Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。</span><span class="sxs-lookup"><span data-stu-id="fb39f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="fb39f-133">型:</span><span class="sxs-lookup"><span data-stu-id="fb39f-133">Type:</span></span>

*   <span data-ttu-id="fb39f-134">Object</span><span class="sxs-lookup"><span data-stu-id="fb39f-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="fb39f-135">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="fb39f-135">Properties:</span></span>

|<span data-ttu-id="fb39f-136">名前</span><span class="sxs-lookup"><span data-stu-id="fb39f-136">Name</span></span>| <span data-ttu-id="fb39f-137">型</span><span class="sxs-lookup"><span data-stu-id="fb39f-137">Type</span></span>| <span data-ttu-id="fb39f-138">説明</span><span class="sxs-lookup"><span data-stu-id="fb39f-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="fb39f-139">String</span><span class="sxs-lookup"><span data-stu-id="fb39f-139">String</span></span>|<span data-ttu-id="fb39f-140">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="fb39f-141">String</span><span class="sxs-lookup"><span data-stu-id="fb39f-141">String</span></span>|<span data-ttu-id="fb39f-142">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="fb39f-143">String</span><span class="sxs-lookup"><span data-stu-id="fb39f-143">String</span></span>|<span data-ttu-id="fb39f-144">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="fb39f-145">String</span><span class="sxs-lookup"><span data-stu-id="fb39f-145">String</span></span>|<span data-ttu-id="fb39f-146">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fb39f-147">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-147">Requirements</span></span>

|<span data-ttu-id="fb39f-148">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-148">Requirement</span></span>| <span data-ttu-id="fb39f-149">値</span><span class="sxs-lookup"><span data-stu-id="fb39f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb39f-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb39f-150">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb39f-151">1.3</span><span class="sxs-lookup"><span data-stu-id="fb39f-151">1.3</span></span>|
|[<span data-ttu-id="fb39f-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb39f-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fb39f-153">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fb39f-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fb39f-154">例</span><span class="sxs-lookup"><span data-stu-id="fb39f-154">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="fb39f-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="fb39f-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="fb39f-156">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="fb39f-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="fb39f-157">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="fb39f-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="fb39f-158">型:</span><span class="sxs-lookup"><span data-stu-id="fb39f-158">Type:</span></span>

*   [<span data-ttu-id="fb39f-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="fb39f-159">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="fb39f-160">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-160">Requirements</span></span>

|<span data-ttu-id="fb39f-161">要件</span><span class="sxs-lookup"><span data-stu-id="fb39f-161">Requirement</span></span>| <span data-ttu-id="fb39f-162">値</span><span class="sxs-lookup"><span data-stu-id="fb39f-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="fb39f-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fb39f-163">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fb39f-164">1.0</span><span class="sxs-lookup"><span data-stu-id="fb39f-164">1.0</span></span>|
|[<span data-ttu-id="fb39f-165">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fb39f-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fb39f-166">制限あり</span><span class="sxs-lookup"><span data-stu-id="fb39f-166">Restricted</span></span>|
|[<span data-ttu-id="fb39f-167">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fb39f-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fb39f-168">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fb39f-168">Compose or read</span></span>|