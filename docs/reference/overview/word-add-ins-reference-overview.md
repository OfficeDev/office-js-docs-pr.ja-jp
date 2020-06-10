---
title: Word JavaScript API の概要
description: Word JavaScript API の概要
ms.date: 02/19/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: efcf963e311f57aab781e1b71c0cd45d4655928f
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679298"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="323b5-103">Word JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="323b5-103">Word JavaScript API overview</span></span>

<span data-ttu-id="323b5-104">Word アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Word のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="323b5-104">An Word add-in interacts with objects in Word by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="323b5-105">**Word JavaScript API**: Office 2016 で導入された [Word JavaScript API](/javascript/api/word) には、Word 文書内のオブジェクトとメタデータへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="323b5-105">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="323b5-106">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="323b5-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="323b5-107">ドキュメントのこのセクションでは、Word JavaScript API に焦点を当てて、そしてそれを Word on the web または Word 2016 以降を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="323b5-107">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="323b5-108">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="323b5-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="323b5-109">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="323b5-109">Learn programming concepts</span></span>

<span data-ttu-id="323b5-110">重要なプログラミング概念の詳細については、「[Word JavaScript API を使用した基本的なプログラミングの概念](../../word/word-add-ins-core-concepts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="323b5-110">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="323b5-111">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="323b5-111">Learn about API capabilities</span></span>

<span data-ttu-id="323b5-112">ドキュメントのこのセクションに記載されている他の記事を参照すると、[アドインからドキュメント全体を取得する](../../word/get-the-whole-document-from-an-add-in-for-word.md)方法、[検索オプションを使用して Word アドインでテキストを検索する](../../word/search-option-guidance.md)方法などを学習できます。</span><span class="sxs-lookup"><span data-stu-id="323b5-112">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="323b5-113">すべての提供可能な記事の一覧については、目次でご確認ください。</span><span class="sxs-lookup"><span data-stu-id="323b5-113">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="323b5-114">Word JavaScript API を使用して Word のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Word アドインのチュートリアル](../../tutorials/word-tutorial.md)」を完了してください。</span><span class="sxs-lookup"><span data-stu-id="323b5-114">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="323b5-115">Word JavaScript API オブジェクト モデルの詳細については、[Word JavaScript API リファレンス ドキュメント](/javascript/api/word)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="323b5-115">For detailed information about the Word JavaScript API object model, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="323b5-116">Script Lab でコード サンプルを試してみる</span><span class="sxs-lookup"><span data-stu-id="323b5-116">Try out code samples in Script Lab</span></span>

<span data-ttu-id="323b5-117">[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。</span><span class="sxs-lookup"><span data-stu-id="323b5-117">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="323b5-118">Script Lab のサンプルを実行すると、作業ウィンドウまたはドキュメントですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。</span><span class="sxs-lookup"><span data-stu-id="323b5-118">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="323b5-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="323b5-119">See also</span></span>

- [<span data-ttu-id="323b5-120">Word アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="323b5-120">Word add-ins documentation</span></span>](../../word/index.yml)
- [<span data-ttu-id="323b5-121">Word アドインの概要</span><span class="sxs-lookup"><span data-stu-id="323b5-121">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="323b5-122">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="323b5-122">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="323b5-123">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="323b5-123">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
