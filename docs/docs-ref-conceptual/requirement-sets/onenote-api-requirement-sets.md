# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="e76d1-101">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e76d1-101">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="e76d1-102">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="e76d1-102">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="e76d1-103">Office アドインでは、マニフェストで指定されている要件のセットを使用して、またはランタイム チェックを使用して、Office ホストがアドインを必要とする Api をサポートしているかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="e76d1-103">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="e76d1-104">詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e76d1-104">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="e76d1-105">次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="e76d1-105">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="e76d1-106">要件セット</span><span class="sxs-lookup"><span data-stu-id="e76d1-106">Requirement set</span></span>  |  <span data-ttu-id="e76d1-107">Office Online</span><span class="sxs-lookup"><span data-stu-id="e76d1-107">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="e76d1-108">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="e76d1-108">OneNoteApi 1.1</span></span>  | <span data-ttu-id="e76d1-109">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="e76d1-109">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="e76d1-110">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e76d1-110">Office common API requirement sets</span></span>

<span data-ttu-id="e76d1-111">共通 API の要件セットについて詳しくは、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e76d1-111">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="e76d1-112">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="e76d1-112">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="e76d1-113">OneNote JavaScript API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="e76d1-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="e76d1-114">API についての詳細は、 [OneNote の JavaScript API のプログラミングの概要](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e76d1-114">For details about the API, see the [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="e76d1-115">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="e76d1-115">Runtime requirement support check</span></span>

<span data-ttu-id="e76d1-116">実行時に、アドインは次のチェックを行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="e76d1-116">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="e76d1-117">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="e76d1-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="e76d1-p103">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで Requirements 要素を使用します。Office ホストまたはプラットフォームが、Requirements 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="e76d1-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="e76d1-120">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="e76d1-120">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="e76d1-121">関連項目</span><span class="sxs-lookup"><span data-stu-id="e76d1-121">See also</span></span>

- [<span data-ttu-id="e76d1-122">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="e76d1-122">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="e76d1-123">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="e76d1-123">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="e76d1-124">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="e76d1-124">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
