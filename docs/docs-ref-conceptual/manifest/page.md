# <a name="page-element"></a><span data-ttu-id="569cd-101">Page 要素</span><span class="sxs-lookup"><span data-stu-id="569cd-101">Page element</span></span>

<span data-ttu-id="569cd-102">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="569cd-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="569cd-103">属性</span><span class="sxs-lookup"><span data-stu-id="569cd-103">Attributes</span></span>

<span data-ttu-id="569cd-104">なし</span><span class="sxs-lookup"><span data-stu-id="569cd-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="569cd-105">子要素</span><span class="sxs-lookup"><span data-stu-id="569cd-105">Child elements</span></span>

|  <span data-ttu-id="569cd-106">要素</span><span class="sxs-lookup"><span data-stu-id="569cd-106">Element</span></span>  |  <span data-ttu-id="569cd-107">必須</span><span class="sxs-lookup"><span data-stu-id="569cd-107">Required</span></span>  |  <span data-ttu-id="569cd-108">説明</span><span class="sxs-lookup"><span data-stu-id="569cd-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="569cd-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="569cd-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="569cd-110">はい</span><span class="sxs-lookup"><span data-stu-id="569cd-110">Yes</span></span>  | <span data-ttu-id="569cd-111">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="569cd-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="569cd-112">例</span><span class="sxs-lookup"><span data-stu-id="569cd-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
