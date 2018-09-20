# <a name="supertip"></a><span data-ttu-id="fd50b-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="fd50b-101">Supertip</span></span>

<span data-ttu-id="fd50b-p101">豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="fd50b-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="fd50b-104">子要素</span><span class="sxs-lookup"><span data-stu-id="fd50b-104">Child elements</span></span>

|  <span data-ttu-id="fd50b-105">要素</span><span class="sxs-lookup"><span data-stu-id="fd50b-105">Element</span></span> |  <span data-ttu-id="fd50b-106">必須</span><span class="sxs-lookup"><span data-stu-id="fd50b-106">Required</span></span>  |  <span data-ttu-id="fd50b-107">説明</span><span class="sxs-lookup"><span data-stu-id="fd50b-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fd50b-108">タイトル</span><span class="sxs-lookup"><span data-stu-id="fd50b-108">Title</span></span>](#title)        | <span data-ttu-id="fd50b-109">はい</span><span class="sxs-lookup"><span data-stu-id="fd50b-109">Yes</span></span> |   <span data-ttu-id="fd50b-110">ヒントのテキストです。</span><span class="sxs-lookup"><span data-stu-id="fd50b-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="fd50b-111">説明</span><span class="sxs-lookup"><span data-stu-id="fd50b-111">Description</span></span>](#description)  | <span data-ttu-id="fd50b-112">はい</span><span class="sxs-lookup"><span data-stu-id="fd50b-112">Yes</span></span> |  <span data-ttu-id="fd50b-113">ヒントの説明です。</span><span class="sxs-lookup"><span data-stu-id="fd50b-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="fd50b-114">タイトル</span><span class="sxs-lookup"><span data-stu-id="fd50b-114">Title</span></span>

<span data-ttu-id="fd50b-p102">必ず指定します。ヒントのテキストです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd50b-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="fd50b-118">説明</span><span class="sxs-lookup"><span data-stu-id="fd50b-118">Description</span></span>

<span data-ttu-id="fd50b-p103">必ず指定します。ヒントの記述です。 **resid** 属性には、 **Resources** 要素の **LongStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd50b-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="fd50b-122">例</span><span class="sxs-lookup"><span data-stu-id="fd50b-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
