---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API リファレンス
description: ホストとバージョン別の Office JavaScript API。
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 06/17/2020
ms.openlocfilehash: aff744f62d55449200a821634510ac3da5ea41c0
ms.sourcegitcommit: e94c95582f58781bf193461f1b8148fac833dba0
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2020
ms.locfileid: "45365091"
---
# <a name="office-add-ins-javascript-api-reference"></a>Office アドインの JavaScript API リファレンス

JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。 このセクションでは、Office アドインの構築に使用できるクラス、メソッド、その他の型について説明します。

以下は、[サポートされている Office ホスト アプリケーション](/office/dev/add-ins/overview/office-add-in-availability)の API の一覧です。 共通 API リンクには、特定のホストに関連付けられていないすべての API が含まれます ([Office 共通 API の要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets))。 他のアイテムは、要求セットに基づいて、そのホストの API リファレンス ドキュメントのバージョンにリンクしています。 リファレンス ドキュメントはバージョン管理され、その要件セットまでのすべての API が含まれます (たとえば、ExcelApi 1.3 は、ExcelApi 1.1、1.2、1.3 の API、および共通 API を示します)。

`ExcelApiOnline 1.1` は特別な要件セットです。 このセットには、Excel on the web の最新の API が含まれていますが、これらの API はまだすべてのプラットフォーム間で完全にサポートされていない可能性があります。 詳細については、「[Excel JavaScript API のオンラインのみの要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set)」を参照してください。

> [!TIP]
> リファレンス ページのバージョンは、目次の上にあるフィルター選択のドロップダウン メニューを使用していつでも変更できます。 特定のバージョンにページが存在しない場合は、現在のバージョンに戻ります。

<h2>Office ホスト</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Excel API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">ExcelApi プレビュー</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">共通 API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Outlook API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Mailbox プレビュー</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Mailbox 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Mailbox 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Mailbox 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Mailbox 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Mailbox 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Mailbox 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Mailbox 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Mailbox 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">共通 API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Word API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">WordApi プレビュー</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">共通 API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>OneNote API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">共通 API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Visio API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>PowerPoint API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-1.1">共通 API</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Project API</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">共通 API のみ</a></li>
                </ul>
            </div>
        </a>
    </li>
</ul>

> [!NOTE]
> Office スクリプトを開発するための JavaScript API をお探しの場合は、「[Office スクリプト API リファレンス](/javascript/api/office-scripts/overview)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインについて](/office/dev/add-ins/overview)
- [Office アドインを使用できるホストおよびプラットフォーム](/office/dev/add-ins/overview/office-add-in-availability)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Script Lab を使用して Office JavaScript API を探索する](/office/dev/add-ins/overview/explore-with-script-lab)
