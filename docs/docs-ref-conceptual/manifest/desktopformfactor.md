# <a name="desktopformfactor-element"></a><span data-ttu-id="be628-101">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="be628-101">DesktopFormFactor element</span></span>

<span data-ttu-id="be628-p101">デスクトップ フォーム ファクターについてアドインの設定を指定します。デスクトップ フォーム ファクターには、Office for Windows、Office for Mac、Office Online が含まれています。**Resources** ノードを除くデスクトップ フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="be628-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="be628-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="be628-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="be628-107">SupportsSharedFolders 要素には、Outlook のアドイン プレビュー要件セットに Exchange Online をできるだけです。</span><span class="sxs-lookup"><span data-stu-id="be628-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="be628-108">この要素を使用するアドインは、Office ストアまたは集中型の展開で許可されていません。</span><span class="sxs-lookup"><span data-stu-id="be628-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="be628-109">子要素</span><span class="sxs-lookup"><span data-stu-id="be628-109">Child elements</span></span>

| <span data-ttu-id="be628-110">要素</span><span class="sxs-lookup"><span data-stu-id="be628-110">Element</span></span>                               | <span data-ttu-id="be628-111">必須</span><span class="sxs-lookup"><span data-stu-id="be628-111">Required</span></span> | <span data-ttu-id="be628-112">説明</span><span class="sxs-lookup"><span data-stu-id="be628-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="be628-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="be628-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="be628-114">はい</span><span class="sxs-lookup"><span data-stu-id="be628-114">Yes</span></span>      | <span data-ttu-id="be628-115">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="be628-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="be628-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="be628-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="be628-117">はい</span><span class="sxs-lookup"><span data-stu-id="be628-117">Yes</span></span>      | <span data-ttu-id="be628-118">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="be628-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="be628-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="be628-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="be628-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="be628-120">No</span></span>       | <span data-ttu-id="be628-121">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="be628-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="be628-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="be628-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="be628-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="be628-123">No</span></span>       | <span data-ttu-id="be628-124">かどうか、Outlook アドインが既定で*false*に設定し、委任のシナリオでは使用を定義します。</span><span class="sxs-lookup"><span data-stu-id="be628-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="be628-125">要件のセットをプレビューします。</span><span class="sxs-lookup"><span data-stu-id="be628-125">Preview requirement set.</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="be628-126">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="be628-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
