---
title: Excel JavaScript API の概要
description: Excel JavaScript API の詳細情報
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293661"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="bc3f5-103">Excel JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="bc3f5-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="bc3f5-104">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="bc3f5-105">**Excel JavaScript API**: これは、Excel 用の [アプリケーション固有 API](../../develop/application-specific-api-model.md) です。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="bc3f5-106">Office 2016 で導入された [Excel JavaScript API](/javascript/api/excel) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="bc3f5-107">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="bc3f5-108">ドキュメントのこのセクションでは、Excel JavaScript API に焦点を当てて、そしてそれを Excel on the web または Excel 2016 以降を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="bc3f5-109">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="bc3f5-110">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="bc3f5-110">Learn programming concepts</span></span>

<span data-ttu-id="bc3f5-111">重要なプログラミング概念の詳細については、「[Excel JavaScript API を使用した基本的なプログラミングの概念](../../excel/excel-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-111">See [Fundamental programming concepts with the Excel JavaScript API](../../excel/excel-add-ins-core-concepts.md) for information about important programming concepts.</span></span>

<span data-ttu-id="bc3f5-112">Excel JavaScript API を使用して Excel のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Excel アドインのチュートリアル](../../tutorials/excel-tutorial.md)」を完了してください。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="bc3f5-113">API 機能について</span><span class="sxs-lookup"><span data-stu-id="bc3f5-113">Learn API capabilities</span></span>

<span data-ttu-id="bc3f5-114">主要な Excel API 機能にはそれぞれ、その機能が実行できることと関連するオブジェクト モデルについての記事があります。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-114">Each major Excel API feature has an article exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="bc3f5-115">グラフ</span><span class="sxs-lookup"><span data-stu-id="bc3f5-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="bc3f5-116">コメント</span><span class="sxs-lookup"><span data-stu-id="bc3f5-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="bc3f5-117">条件付き書式</span><span class="sxs-lookup"><span data-stu-id="bc3f5-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="bc3f5-118">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="bc3f5-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="bc3f5-119">データ検証</span><span class="sxs-lookup"><span data-stu-id="bc3f5-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="bc3f5-120">イベント</span><span class="sxs-lookup"><span data-stu-id="bc3f5-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="bc3f5-121">複数の範囲 (範囲領域)</span><span class="sxs-lookup"><span data-stu-id="bc3f5-121">Multiple ranges (RangeArea)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="bc3f5-122">ピボットテーブル</span><span class="sxs-lookup"><span data-stu-id="bc3f5-122">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="bc3f5-123">[範囲](../../excel/excel-add-ins-ranges.md) および [高度な範囲 API](../../excel/excel-add-ins-ranges-advanced.md)</span><span class="sxs-lookup"><span data-stu-id="bc3f5-123">[Ranges](../../excel/excel-add-ins-ranges.md) and [Advanced Range APIs](../../excel/excel-add-ins-ranges-advanced.md)</span></span>
* [<span data-ttu-id="bc3f5-124">図形</span><span class="sxs-lookup"><span data-stu-id="bc3f5-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="bc3f5-125">表</span><span class="sxs-lookup"><span data-stu-id="bc3f5-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="bc3f5-126">ブックとアプリケーションレベルの API</span><span class="sxs-lookup"><span data-stu-id="bc3f5-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="bc3f5-127">ワークシート</span><span class="sxs-lookup"><span data-stu-id="bc3f5-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="bc3f5-128">Excel JavaScript API オブジェクト モデルに関する詳細情報については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="bc3f5-129">Script Lab でコード サンプルを試してみる</span><span class="sxs-lookup"><span data-stu-id="bc3f5-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="bc3f5-130">[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="bc3f5-131">Script Lab のサンプルを実行すると、作業ウィンドウまたはワークシートですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="bc3f5-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="bc3f5-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="bc3f5-132">See also</span></span>

* [<span data-ttu-id="bc3f5-133">Excel アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="bc3f5-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="bc3f5-134">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="bc3f5-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="bc3f5-135">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="bc3f5-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="bc3f5-136">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="bc3f5-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="bc3f5-137">アプリケーション固有の API モデルの使用</span><span class="sxs-lookup"><span data-stu-id="bc3f5-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
