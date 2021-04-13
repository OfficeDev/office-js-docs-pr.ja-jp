---
title: Excel JavaScript API の概要
description: Excel JavaScript API の詳細情報
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 80340b4990b56b2ba4d51f2a028480af3e267828
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650812"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="703e2-103">Excel JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="703e2-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="703e2-104">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="703e2-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="703e2-105">**Excel JavaScript API**: これは、Excel 用の [アプリケーション固有 API](../../develop/application-specific-api-model.md) です。</span><span class="sxs-lookup"><span data-stu-id="703e2-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="703e2-106">Office 2016 で導入された [Excel JavaScript API](/javascript/api/excel) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="703e2-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="703e2-107">**共通 API**: Office 2013 で導入された [共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="703e2-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="703e2-108">ドキュメントのこのセクションでは、Excel JavaScript API に焦点を当てて、そしてそれを Excel on the web または Excel 2016 以降を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="703e2-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="703e2-109">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="703e2-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-object-model-concepts"></a><span data-ttu-id="703e2-110">オブジェクト モデルの概念について</span><span class="sxs-lookup"><span data-stu-id="703e2-110">Learn object model concepts</span></span>

<span data-ttu-id="703e2-111">重要なオブジェクト モデルの概念については、「[Office アドインの Excel JavaScript オブジェクト モデル](../../excel/excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="703e2-111">See [Excel JavaScript object model in Office Add-ins](../../excel/excel-add-ins-core-concepts.md) for information about important object model concepts.</span></span>

<span data-ttu-id="703e2-112">Excel JavaScript API を使用して Excel のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Excel アドインのチュートリアル](../../tutorials/excel-tutorial.md)」を完了してください。</span><span class="sxs-lookup"><span data-stu-id="703e2-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="703e2-113">API 機能について</span><span class="sxs-lookup"><span data-stu-id="703e2-113">Learn API capabilities</span></span>

<span data-ttu-id="703e2-114">主要な Excel API 機能にはそれぞれ、その機能が実行できることと関連するオブジェクト モデルについての記事または記事のセットがあります。</span><span class="sxs-lookup"><span data-stu-id="703e2-114">Each major Excel API feature has an article or set of articles exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="703e2-115">グラフ</span><span class="sxs-lookup"><span data-stu-id="703e2-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="703e2-116">コメント</span><span class="sxs-lookup"><span data-stu-id="703e2-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="703e2-117">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="703e2-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="703e2-118">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="703e2-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="703e2-119">データ検証</span><span class="sxs-lookup"><span data-stu-id="703e2-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="703e2-120">イベント</span><span class="sxs-lookup"><span data-stu-id="703e2-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="703e2-121">PivotTables</span><span class="sxs-lookup"><span data-stu-id="703e2-121">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="703e2-122">[Range](../../excel/excel-add-ins-ranges-get.md) および [Cells](../../excel/excel-add-ins-cells.md)</span><span class="sxs-lookup"><span data-stu-id="703e2-122">[Ranges](../../excel/excel-add-ins-ranges-get.md) and [Cells](../../excel/excel-add-ins-cells.md)</span></span>
* [<span data-ttu-id="703e2-123">RangeAreas (複数の範囲)</span><span class="sxs-lookup"><span data-stu-id="703e2-123">RangeAreas (Multiple ranges)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="703e2-124">図形</span><span class="sxs-lookup"><span data-stu-id="703e2-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="703e2-125">表</span><span class="sxs-lookup"><span data-stu-id="703e2-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="703e2-126">ブックとアプリケーションレベルの API</span><span class="sxs-lookup"><span data-stu-id="703e2-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="703e2-127">ワークシート</span><span class="sxs-lookup"><span data-stu-id="703e2-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="703e2-128">Excel JavaScript API オブジェクト モデルに関する詳細情報については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="703e2-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="703e2-129">Script Lab でコード サンプルを試してみる</span><span class="sxs-lookup"><span data-stu-id="703e2-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="703e2-130">[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="703e2-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="703e2-131">Script Lab のサンプルを実行すると、作業ウィンドウまたはワークシートですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="703e2-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="703e2-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="703e2-132">See also</span></span>

* [<span data-ttu-id="703e2-133">Excel アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="703e2-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="703e2-134">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="703e2-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="703e2-135">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="703e2-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="703e2-136">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="703e2-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="703e2-137">アプリケーション固有の API モデルの使用</span><span class="sxs-lookup"><span data-stu-id="703e2-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
