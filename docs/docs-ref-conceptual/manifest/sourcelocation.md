# <a name="sourcelocation-element"></a><span data-ttu-id="47f74-101">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="47f74-101">SourceLocation element</span></span>

<span data-ttu-id="47f74-p101">Office アドインのソース ファイルの場所を、1 から 2018 文字までの長さの URL として指定します。ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="47f74-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="47f74-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="47f74-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="47f74-105">構文</span><span class="sxs-lookup"><span data-stu-id="47f74-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="47f74-106">含まれています。</span><span class="sxs-lookup"><span data-stu-id="47f74-106">Contained in</span></span>

- <span data-ttu-id="47f74-107">[DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)</span><span class="sxs-lookup"><span data-stu-id="47f74-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="47f74-108">[FormSettings](formsettings.md) (メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="47f74-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="47f74-109">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン)</span><span class="sxs-lookup"><span data-stu-id="47f74-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="47f74-110">含めることができます。</span><span class="sxs-lookup"><span data-stu-id="47f74-110">Can contain</span></span>

[<span data-ttu-id="47f74-111">Override</span><span class="sxs-lookup"><span data-stu-id="47f74-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="47f74-112">属性</span><span class="sxs-lookup"><span data-stu-id="47f74-112">Attributes</span></span>

|<span data-ttu-id="47f74-113">**属性**</span><span class="sxs-lookup"><span data-stu-id="47f74-113">**Attribute**</span></span>|<span data-ttu-id="47f74-114">**型**</span><span class="sxs-lookup"><span data-stu-id="47f74-114">**Type**</span></span>|<span data-ttu-id="47f74-115">**必須**</span><span class="sxs-lookup"><span data-stu-id="47f74-115">**Required**</span></span>|<span data-ttu-id="47f74-116">**説明**</span><span class="sxs-lookup"><span data-stu-id="47f74-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="47f74-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="47f74-117">DefaultValue</span></span>|<span data-ttu-id="47f74-118">URL</span><span class="sxs-lookup"><span data-stu-id="47f74-118">URL</span></span>|<span data-ttu-id="47f74-119">必須</span><span class="sxs-lookup"><span data-stu-id="47f74-119">required</span></span>|<span data-ttu-id="47f74-120">[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。</span><span class="sxs-lookup"><span data-stu-id="47f74-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
