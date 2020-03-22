---
title: OneNote JavaScript API の概要
description: OneNote JavaScript API の詳細情報
ms.date: 02/19/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 70c3bca323084630f1926b501900bca26cf54304
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719923"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="c0a14-103">OneNote JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="c0a14-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="c0a14-104">OneNote アドインでは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して OneNote on the web のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="c0a14-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="c0a14-105">**OneNote JavaScript API**: Office 2016 で導入された [OneNote JavaScript API](/javascript/api/onenote) には、OneNote on the web へのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="c0a14-105">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="c0a14-106">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c0a14-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="c0a14-107">ドキュメントのこのセクションでは、OneNote JavaScript API に焦点を当てて、そしてそれを OneNote on the web を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="c0a14-107">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="c0a14-108">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0a14-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="c0a14-109">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="c0a14-109">Learn programming concepts</span></span>

<span data-ttu-id="c0a14-110">重要なプログラミングの概念に関する詳細情報については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0a14-110">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="c0a14-111">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="c0a14-111">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="c0a14-112">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="c0a14-112">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="c0a14-113">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="c0a14-113">Learn about API capabilities</span></span>

<span data-ttu-id="c0a14-114">OneNote JavaScript API を使用して OneNote on the web のコンテンツと対話するには、「[OneNote アドインのクイック スタート](../../quickstarts/onenote-quickstart.md)」を実行します。</span><span class="sxs-lookup"><span data-stu-id="c0a14-114">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="c0a14-115">OneNote JavaScript API オブジェクト モデルの詳細については、[OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="c0a14-115">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="c0a14-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="c0a14-116">See also</span></span>

- [<span data-ttu-id="c0a14-117">OneNote アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="c0a14-117">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="c0a14-118">OneNote アドインの概要</span><span class="sxs-lookup"><span data-stu-id="c0a14-118">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="c0a14-119">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="c0a14-119">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="c0a14-120">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="c0a14-120">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

