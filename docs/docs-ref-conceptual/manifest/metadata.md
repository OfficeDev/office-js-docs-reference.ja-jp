# <a name="metadata-element"></a><span data-ttu-id="a3b62-101">メタデータ要素</span><span class="sxs-lookup"><span data-stu-id="a3b62-101">Metadata element</span></span>

<span data-ttu-id="a3b62-102">Excel でユーザー定義関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="a3b62-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a3b62-103">属性</span><span class="sxs-lookup"><span data-stu-id="a3b62-103">Attributes</span></span>

<span data-ttu-id="a3b62-104">なし</span><span class="sxs-lookup"><span data-stu-id="a3b62-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a3b62-105">子要素</span><span class="sxs-lookup"><span data-stu-id="a3b62-105">Child elements</span></span>

|  <span data-ttu-id="a3b62-106">要素</span><span class="sxs-lookup"><span data-stu-id="a3b62-106">Element</span></span>  |  <span data-ttu-id="a3b62-107">必須</span><span class="sxs-lookup"><span data-stu-id="a3b62-107">Required</span></span>  |  <span data-ttu-id="a3b62-108">説明</span><span class="sxs-lookup"><span data-stu-id="a3b62-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a3b62-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a3b62-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a3b62-110">はい</span><span class="sxs-lookup"><span data-stu-id="a3b62-110">Yes</span></span>  | <span data-ttu-id="a3b62-111">カスタム関数で使用される JSON ファイルのリソース id を持つ文字列です。</span><span class="sxs-lookup"><span data-stu-id="a3b62-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a3b62-112">例</span><span class="sxs-lookup"><span data-stu-id="a3b62-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
