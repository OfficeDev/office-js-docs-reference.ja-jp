---
title: Office JavaScript API リファレンス
description: ホスト要件セットごとの Office JavaScript Api
ms.date: 11/19/2019
ms.openlocfilehash: f4072c23cb0d6e0d5375cf79d92b4f6dd9b35f0f
ms.sourcegitcommit: d37268ff5254061632a886b196ec28f2f4087377
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/21/2019
ms.locfileid: "38758618"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="7cbf1-103">Office JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7cbf1-103">Office JavaScript API reference</span></span>

<span data-ttu-id="7cbf1-104">JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="7cbf1-105">このセクションを使用して、Office アドインの構築に使用できるクラス、メソッド、およびその他の種類の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="7cbf1-106">ホスト固有の要件セット (およびクロスホスト共通 Api) の一覧を次に示します。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="7cbf1-107">各アイテムは、その要件セットでサポートされているバージョンの API リファレンスドキュメント (ExcelApi 1.3 に、ExcelApi 1.1、1.2、1.3 および Common API の Api を示します) へのリンクを掲載しています。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="7cbf1-108">`ExcelApiOnline 1.1`特別な要件セットです。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="7cbf1-109">Web 上の Excel 用の最新の Api が含まれていますが、これらの Api はすべてのプラットフォームで完全にサポートされているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="7cbf1-110">詳細については、「 [Excel JAVASCRIPT API online 専用の要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="7cbf1-111">このページのリンクを選択して、指定された要件セットでサポートされている Api の参照ドキュメントを表示するか、または目次の上にある [フィルターの選択] ドロップダウンメニューを使用して、要件セットをいつでも変更できます。</span><span class="sxs-lookup"><span data-stu-id="7cbf1-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="7cbf1-112">Excel</span><span class="sxs-lookup"><span data-stu-id="7cbf1-112">Excel</span></span>

- [<span data-ttu-id="7cbf1-113">ExcelApi プレビュー</span><span class="sxs-lookup"><span data-stu-id="7cbf1-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="7cbf1-114">ExcelApiOnline 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="7cbf1-115">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="7cbf1-115">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="7cbf1-116">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="7cbf1-116">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="7cbf1-117">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="7cbf1-117">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="7cbf1-118">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="7cbf1-118">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="7cbf1-119">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="7cbf1-119">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="7cbf1-120">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="7cbf1-120">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="7cbf1-121">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="7cbf1-121">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="7cbf1-122">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="7cbf1-122">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="7cbf1-123">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="7cbf1-123">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="7cbf1-124">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-124">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="7cbf1-125">OneNote</span><span class="sxs-lookup"><span data-stu-id="7cbf1-125">OneNote</span></span>

- [<span data-ttu-id="7cbf1-126">OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-126">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="7cbf1-127">Outlook</span><span class="sxs-lookup"><span data-stu-id="7cbf1-127">Outlook</span></span>

- [<span data-ttu-id="7cbf1-128">メールボックスプレビュー</span><span class="sxs-lookup"><span data-stu-id="7cbf1-128">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="7cbf1-129">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="7cbf1-129">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="7cbf1-130">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="7cbf1-130">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="7cbf1-131">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="7cbf1-131">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="7cbf1-132">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="7cbf1-132">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="7cbf1-133">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="7cbf1-133">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="7cbf1-134">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="7cbf1-134">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="7cbf1-135">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="7cbf1-135">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="7cbf1-136">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-136">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="7cbf1-137">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7cbf1-137">PowerPoint</span></span>

- [<span data-ttu-id="7cbf1-138">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-138">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="7cbf1-139">Visio</span><span class="sxs-lookup"><span data-stu-id="7cbf1-139">Visio</span></span>

- [<span data-ttu-id="7cbf1-140">VisioApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-140">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="7cbf1-141">Word</span><span class="sxs-lookup"><span data-stu-id="7cbf1-141">Word</span></span>

- [<span data-ttu-id="7cbf1-142">Word プレビュー</span><span class="sxs-lookup"><span data-stu-id="7cbf1-142">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="7cbf1-143">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="7cbf1-143">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="7cbf1-144">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="7cbf1-144">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="7cbf1-145">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7cbf1-145">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="7cbf1-146">共通 API</span><span class="sxs-lookup"><span data-stu-id="7cbf1-146">Common API</span></span>

- [<span data-ttu-id="7cbf1-147">共通 API</span><span class="sxs-lookup"><span data-stu-id="7cbf1-147">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="7cbf1-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="7cbf1-148">See also</span></span>

- [<span data-ttu-id="7cbf1-149">Office アドインについて</span><span class="sxs-lookup"><span data-stu-id="7cbf1-149">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="7cbf1-150">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="7cbf1-150">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="7cbf1-151">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="7cbf1-151">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
