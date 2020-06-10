---
title: Excel JavaScript API の概要
description: Excel JavaScript API の詳細情報
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 218f328468994065acda91c11b38659d7a20fe15
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679347"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="6613e-103">Excel JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="6613e-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="6613e-104">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="6613e-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="6613e-105">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](/javascript/api/excel) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="6613e-105">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="6613e-106">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="6613e-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="6613e-107">ドキュメントのこのセクションでは、Excel JavaScript API に焦点を当てて、そしてそれを Excel on the web または Excel 2016 以降を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="6613e-107">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="6613e-108">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6613e-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="6613e-109">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="6613e-109">Learn programming concepts</span></span>

<span data-ttu-id="6613e-110">重要なプログラミングの概念に関する詳細情報については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6613e-110">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="6613e-111">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="6613e-111">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="6613e-112">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="6613e-112">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="6613e-113">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="6613e-113">Learn about API capabilities</span></span>

<span data-ttu-id="6613e-114">ドキュメントのこのセクションにあるその他の記事を使用して、[イベント](../../excel/excel-add-ins-events.md)、[グラフ](../../excel/excel-add-ins-charts.md)、[範囲](../../excel/excel-add-ins-ranges.md)、[テーブル](../../excel/excel-add-ins-tables.md)、[ワークシート](../../excel/excel-add-ins-worksheets.md)などの操作について学びます。</span><span class="sxs-lookup"><span data-stu-id="6613e-114">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="6613e-115">また、このセクションでは[Excel アドインの共同編集](../../excel/co-authoring-in-excel-add-ins.md)、[データ検証](../../excel/excel-add-ins-data-validation.md)、[エラー処理](../../excel/excel-add-ins-error-handling.md)、[パフォーマンスの最適化](../../excel/performance.md)などの Excel JavaScript API の概念についてのガイダンスを確認できます。</span><span class="sxs-lookup"><span data-stu-id="6613e-115">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="6613e-116">すべての提供可能な記事の一覧については、目次でご確認ください。</span><span class="sxs-lookup"><span data-stu-id="6613e-116">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="6613e-117">Excel JavaScript API を使用して Excel のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Excel アドインのチュートリアル](../../tutorials/excel-tutorial.md)」を完了してください。</span><span class="sxs-lookup"><span data-stu-id="6613e-117">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="6613e-118">Excel JavaScript API オブジェクト モデルに関する詳細情報については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="6613e-118">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="6613e-119">Script Lab でコード サンプルを試してみる</span><span class="sxs-lookup"><span data-stu-id="6613e-119">Try out code samples in Script Lab</span></span>

<span data-ttu-id="6613e-120">[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="6613e-120">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="6613e-121">Script Lab のサンプルを実行すると、作業ウィンドウまたはワークシートですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="6613e-121">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="6613e-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="6613e-122">See also</span></span>

- [<span data-ttu-id="6613e-123">Excel アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="6613e-123">Excel add-ins documentation</span></span>](../../excel/index.yml)
- [<span data-ttu-id="6613e-124">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="6613e-124">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="6613e-125">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="6613e-125">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="6613e-126">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="6613e-126">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
