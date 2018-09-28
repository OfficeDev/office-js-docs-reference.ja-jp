# <a name="supporturl-element"></a><span data-ttu-id="93a44-101">SupportUrl 要素</span><span class="sxs-lookup"><span data-stu-id="93a44-101">SupportUrl element</span></span>

<span data-ttu-id="93a44-102">アドインのサポート情報を提供するページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="93a44-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="93a44-103">構文</span><span class="sxs-lookup"><span data-stu-id="93a44-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="93a44-104">含まれています。</span><span class="sxs-lookup"><span data-stu-id="93a44-104">Contained in</span></span>

[<span data-ttu-id="93a44-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="93a44-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="93a44-106">含めることができます。</span><span class="sxs-lookup"><span data-stu-id="93a44-106">Can contain</span></span>

|  <span data-ttu-id="93a44-107">要素</span><span class="sxs-lookup"><span data-stu-id="93a44-107">Element</span></span> | <span data-ttu-id="93a44-108">必須</span><span class="sxs-lookup"><span data-stu-id="93a44-108">Required</span></span> | <span data-ttu-id="93a44-109">説明</span><span class="sxs-lookup"><span data-stu-id="93a44-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="93a44-110">Override</span><span class="sxs-lookup"><span data-stu-id="93a44-110">Override</span></span>](override.md)   | <span data-ttu-id="93a44-111">なし</span><span class="sxs-lookup"><span data-stu-id="93a44-111">No</span></span> | <span data-ttu-id="93a44-112">追加のロケール URL の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="93a44-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="93a44-113">属性</span><span class="sxs-lookup"><span data-stu-id="93a44-113">Attributes</span></span>

|<span data-ttu-id="93a44-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="93a44-114">**Attribute**</span></span>|<span data-ttu-id="93a44-115">**型**</span><span class="sxs-lookup"><span data-stu-id="93a44-115">**Type**</span></span>|<span data-ttu-id="93a44-116">**必須**</span><span class="sxs-lookup"><span data-stu-id="93a44-116">**Required**</span></span>|<span data-ttu-id="93a44-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="93a44-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="93a44-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="93a44-118">DefaultValue</span></span>|<span data-ttu-id="93a44-119">URL</span><span class="sxs-lookup"><span data-stu-id="93a44-119">URL</span></span>|<span data-ttu-id="93a44-120">必須</span><span class="sxs-lookup"><span data-stu-id="93a44-120">required</span></span>|<span data-ttu-id="93a44-121">この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。</span><span class="sxs-lookup"><span data-stu-id="93a44-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
