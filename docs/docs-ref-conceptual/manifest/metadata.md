# <a name="metadata-element"></a><span data-ttu-id="31f91-101">メタデータ要素</span><span class="sxs-lookup"><span data-stu-id="31f91-101">Metadata element</span></span>

<span data-ttu-id="31f91-102">Excel でユーザー定義関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="31f91-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="31f91-103">属性</span><span class="sxs-lookup"><span data-stu-id="31f91-103">Attributes</span></span>

<span data-ttu-id="31f91-104">なし</span><span class="sxs-lookup"><span data-stu-id="31f91-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="31f91-105">子要素</span><span class="sxs-lookup"><span data-stu-id="31f91-105">Child elements</span></span>

|  <span data-ttu-id="31f91-106">要素</span><span class="sxs-lookup"><span data-stu-id="31f91-106">Element</span></span>  |  <span data-ttu-id="31f91-107">必須</span><span class="sxs-lookup"><span data-stu-id="31f91-107">Required</span></span>  |  <span data-ttu-id="31f91-108">説明</span><span class="sxs-lookup"><span data-stu-id="31f91-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="31f91-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="31f91-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="31f91-110">はい</span><span class="sxs-lookup"><span data-stu-id="31f91-110">Yes</span></span>  | <span data-ttu-id="31f91-111">カスタム関数で使用される JSON ファイルのリソース id を持つ文字列です。</span><span class="sxs-lookup"><span data-stu-id="31f91-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="31f91-112">例</span><span class="sxs-lookup"><span data-stu-id="31f91-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
