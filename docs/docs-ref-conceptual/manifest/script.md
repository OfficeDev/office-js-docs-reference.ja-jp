# <a name="script-element"></a><span data-ttu-id="5588b-101">Script 要素</span><span class="sxs-lookup"><span data-stu-id="5588b-101">Script element</span></span>

<span data-ttu-id="5588b-102">Excel でカスタム関数によって使用されるスクリプトの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="5588b-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5588b-103">属性</span><span class="sxs-lookup"><span data-stu-id="5588b-103">Attributes</span></span>

<span data-ttu-id="5588b-104">なし</span><span class="sxs-lookup"><span data-stu-id="5588b-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5588b-105">子要素</span><span class="sxs-lookup"><span data-stu-id="5588b-105">Child elements</span></span>

|<span data-ttu-id="5588b-106">要素</span><span class="sxs-lookup"><span data-stu-id="5588b-106">Elements</span></span>  |  <span data-ttu-id="5588b-107">必須</span><span class="sxs-lookup"><span data-stu-id="5588b-107">Required</span></span>  |  <span data-ttu-id="5588b-108">説明</span><span class="sxs-lookup"><span data-stu-id="5588b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5588b-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5588b-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5588b-110">はい</span><span class="sxs-lookup"><span data-stu-id="5588b-110">Yes</span></span>  | <span data-ttu-id="5588b-111">カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="5588b-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="5588b-112">例</span><span class="sxs-lookup"><span data-stu-id="5588b-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
