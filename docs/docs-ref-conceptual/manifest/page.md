# <a name="page-element"></a><span data-ttu-id="5b8c0-101">Page 要素</span><span class="sxs-lookup"><span data-stu-id="5b8c0-101">Page element</span></span>

<span data-ttu-id="5b8c0-102">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="5b8c0-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5b8c0-103">属性</span><span class="sxs-lookup"><span data-stu-id="5b8c0-103">Attributes</span></span>

<span data-ttu-id="5b8c0-104">なし</span><span class="sxs-lookup"><span data-stu-id="5b8c0-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5b8c0-105">子要素</span><span class="sxs-lookup"><span data-stu-id="5b8c0-105">Child elements</span></span>

|  <span data-ttu-id="5b8c0-106">要素</span><span class="sxs-lookup"><span data-stu-id="5b8c0-106">Element</span></span>  |  <span data-ttu-id="5b8c0-107">必須</span><span class="sxs-lookup"><span data-stu-id="5b8c0-107">Required</span></span>  |  <span data-ttu-id="5b8c0-108">説明</span><span class="sxs-lookup"><span data-stu-id="5b8c0-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5b8c0-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5b8c0-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5b8c0-110">はい</span><span class="sxs-lookup"><span data-stu-id="5b8c0-110">Yes</span></span>  | <span data-ttu-id="5b8c0-111">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="5b8c0-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="5b8c0-112">例</span><span class="sxs-lookup"><span data-stu-id="5b8c0-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
