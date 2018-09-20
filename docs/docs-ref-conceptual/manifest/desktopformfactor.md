# <a name="desktopformfactor-element"></a><span data-ttu-id="5c929-101">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="5c929-101">DesktopFormFactor element</span></span>

<span data-ttu-id="5c929-p101">デスクトップ フォーム ファクターについてアドインの設定を指定します。デスクトップ フォーム ファクターには、Office for Windows、Office for Mac、Office Online が含まれています。**Resources** ノードを除くデスクトップ フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="5c929-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="5c929-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5c929-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="5c929-107">子要素</span><span class="sxs-lookup"><span data-stu-id="5c929-107">Child elements</span></span>

| <span data-ttu-id="5c929-108">要素</span><span class="sxs-lookup"><span data-stu-id="5c929-108">Element</span></span>                               | <span data-ttu-id="5c929-109">必須</span><span class="sxs-lookup"><span data-stu-id="5c929-109">Required</span></span> | <span data-ttu-id="5c929-110">説明</span><span class="sxs-lookup"><span data-stu-id="5c929-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="5c929-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="5c929-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="5c929-112">はい</span><span class="sxs-lookup"><span data-stu-id="5c929-112">Yes</span></span>      | <span data-ttu-id="5c929-113">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="5c929-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="5c929-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="5c929-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="5c929-115">はい</span><span class="sxs-lookup"><span data-stu-id="5c929-115">Yes</span></span>      | <span data-ttu-id="5c929-116">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="5c929-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="5c929-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="5c929-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="5c929-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="5c929-118">No</span></span>       | <span data-ttu-id="5c929-119">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="5c929-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="5c929-120">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="5c929-120">DesktopFormFactor example</span></span>

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
