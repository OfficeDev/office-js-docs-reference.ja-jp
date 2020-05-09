---
title: Office JavaScript API リファレンス
description: Office JavaScript Api の要件は、ホストによって設定されます。
ms.date: 05/05/2020
ms.openlocfilehash: 3a32c47b23fd6635c4c2b44b58ee9b351fffd8d5
ms.sourcegitcommit: 23d9a58660cb1dedf0bc414849a5aec519b419b3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/07/2020
ms.locfileid: "44146410"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="582bd-103">Office JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="582bd-103">Office JavaScript API reference</span></span>

<span data-ttu-id="582bd-104">JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="582bd-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="582bd-105">このセクションを使用して、Office アドインの構築に使用できるクラス、メソッド、およびその他の種類の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="582bd-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="582bd-106">ホスト固有の要件セット (およびクロスホスト共通 Api) の一覧を次に示します。</span><span class="sxs-lookup"><span data-stu-id="582bd-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="582bd-107">各アイテムは、その要件セットでサポートされているバージョンの API リファレンスドキュメント (ExcelApi 1.3 に、ExcelApi 1.1、1.2、1.3 および Common API の Api を示します) へのリンクを掲載しています。</span><span class="sxs-lookup"><span data-stu-id="582bd-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="582bd-108">`ExcelApiOnline 1.1`特別な要件セットです。</span><span class="sxs-lookup"><span data-stu-id="582bd-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="582bd-109">Web 上の Excel 用の最新の Api が含まれていますが、これらの Api はすべてのプラットフォームで完全にサポートされているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="582bd-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="582bd-110">詳細については、「 [Excel JAVASCRIPT API online 専用の要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="582bd-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="582bd-111">このページのリンクを選択して、指定された要件セットでサポートされている Api の参照ドキュメントを表示するか、または目次の上にある [フィルターの選択] ドロップダウンメニューを使用して、要件セットをいつでも変更できます。</span><span class="sxs-lookup"><span data-stu-id="582bd-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="582bd-112">Excel</span><span class="sxs-lookup"><span data-stu-id="582bd-112">Excel</span></span>

- [<span data-ttu-id="582bd-113">ExcelApi プレビュー</span><span class="sxs-lookup"><span data-stu-id="582bd-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="582bd-114">ExcelApiOnline 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="582bd-115">ExcelApi 1.11</span><span class="sxs-lookup"><span data-stu-id="582bd-115">ExcelApi 1.11</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="582bd-116">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="582bd-116">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="582bd-117">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="582bd-117">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="582bd-118">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="582bd-118">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="582bd-119">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="582bd-119">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="582bd-120">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="582bd-120">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="582bd-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="582bd-121">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="582bd-122">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="582bd-122">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="582bd-123">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="582bd-123">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="582bd-124">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="582bd-124">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="582bd-125">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-125">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="582bd-126">OneNote</span><span class="sxs-lookup"><span data-stu-id="582bd-126">OneNote</span></span>

- [<span data-ttu-id="582bd-127">OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-127">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="582bd-128">Outlook</span><span class="sxs-lookup"><span data-stu-id="582bd-128">Outlook</span></span>

- [<span data-ttu-id="582bd-129">メールボックスプレビュー</span><span class="sxs-lookup"><span data-stu-id="582bd-129">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="582bd-130">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="582bd-130">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="582bd-131">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="582bd-131">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="582bd-132">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="582bd-132">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="582bd-133">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="582bd-133">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="582bd-134">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="582bd-134">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="582bd-135">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="582bd-135">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="582bd-136">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="582bd-136">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="582bd-137">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-137">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="582bd-138">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="582bd-138">PowerPoint</span></span>

- [<span data-ttu-id="582bd-139">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-139">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="582bd-140">Visio</span><span class="sxs-lookup"><span data-stu-id="582bd-140">Visio</span></span>

- [<span data-ttu-id="582bd-141">VisioApi 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-141">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="582bd-142">Word</span><span class="sxs-lookup"><span data-stu-id="582bd-142">Word</span></span>

- [<span data-ttu-id="582bd-143">Word プレビュー</span><span class="sxs-lookup"><span data-stu-id="582bd-143">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="582bd-144">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="582bd-144">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="582bd-145">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="582bd-145">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="582bd-146">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="582bd-146">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="582bd-147">共通 API</span><span class="sxs-lookup"><span data-stu-id="582bd-147">Common API</span></span>

- [<span data-ttu-id="582bd-148">共通 API</span><span class="sxs-lookup"><span data-stu-id="582bd-148">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="582bd-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="582bd-149">See also</span></span>

- [<span data-ttu-id="582bd-150">Office アドインについて</span><span class="sxs-lookup"><span data-stu-id="582bd-150">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="582bd-151">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="582bd-151">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="582bd-152">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="582bd-152">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="582bd-153">Script Lab を使用して Office JavaScript API を探索する</span><span class="sxs-lookup"><span data-stu-id="582bd-153">Explore Office JavaScript API using Script Lab</span></span>](/office/dev/add-ins/overview/explore-with-script-lab)
