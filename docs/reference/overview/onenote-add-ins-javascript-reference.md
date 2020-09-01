---
title: OneNote JavaScript API の概要
description: OneNote JavaScript API の詳細情報
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 08e98e81e46ca62178235454d3ba44f35be2eec2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293633"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="e9566-103">OneNote JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="e9566-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="e9566-104">OneNote アドインでは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して OneNote on the web のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="e9566-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e9566-105">**OneNote JavaScript API**: これは、OneNote 用の [アプリケーション固有 API](../../develop/application-specific-api-model.md) です。</span><span class="sxs-lookup"><span data-stu-id="e9566-105">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote.</span></span> <span data-ttu-id="e9566-106">Office 2016 で導入された [OneNote JavaScript API](/javascript/api/onenote) には、OneNote on the web へのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="e9566-106">Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="e9566-107">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="e9566-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="e9566-108">ドキュメントのこのセクションでは、OneNote JavaScript API に焦点を当てて、そしてそれを OneNote on the web を対象としたアドインの大部分の機能開発に使用します。</span><span class="sxs-lookup"><span data-stu-id="e9566-108">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="e9566-109">共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9566-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="e9566-110">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="e9566-110">Learn programming concepts</span></span>

<span data-ttu-id="e9566-111">重要なプログラミングの概念に関する詳細情報については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9566-111">See the following articles for information about important programming concepts:</span></span>

* [<span data-ttu-id="e9566-112">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="e9566-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="e9566-113">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="e9566-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="e9566-114">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="e9566-114">Learn about API capabilities</span></span>

<span data-ttu-id="e9566-115">OneNote JavaScript API を使用して OneNote on the web のコンテンツと対話するには、「[OneNote アドインのクイック スタート](../../quickstarts/onenote-quickstart.md)」を実行します。</span><span class="sxs-lookup"><span data-stu-id="e9566-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="e9566-116">OneNote JavaScript API オブジェクト モデルの詳細については、[OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9566-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="e9566-117">関連項目</span><span class="sxs-lookup"><span data-stu-id="e9566-117">See also</span></span>

* [<span data-ttu-id="e9566-118">OneNote アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="e9566-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="e9566-119">OneNote アドインの概要</span><span class="sxs-lookup"><span data-stu-id="e9566-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="e9566-120">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="e9566-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="e9566-121">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="e9566-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
