---
title: OneNote JavaScript API の概要
description: OneNote JavaScript API の詳細情報
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: d917d71cd9d3f4fadbab91a434a177c45b54c6f2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349113"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="cb921-103">OneNote JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="cb921-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="cb921-104">OneNote アドインでは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して OneNote on the web のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="cb921-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="cb921-p101">**OneNote JavaScript API**: OneNote には [application-specific APIs](../../develop/application-specific-api-model.md)があり、Office 2016 で導入された [OneNote JavaScript API](/javascript/api/onenote) には、OneNote on the web のオブジェクトへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="cb921-p101">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote. Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="cb921-107">**共通 API**: Office 2013 で導入された [共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="cb921-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="cb921-p102">ドキュメントのこのセクションでは、OneNote JavaScript API に焦点を当てますが、それを OneNote on the Web を対象としたアドインの大部分の機能開発に使用します。共通 API の詳細情報については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cb921-p102">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web. For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="cb921-110">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="cb921-110">Learn programming concepts</span></span>

<span data-ttu-id="cb921-111">重要なプログラミングの概念に関する詳細情報については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb921-111">See the following articles for information about important programming concepts.</span></span>

* [<span data-ttu-id="cb921-112">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="cb921-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="cb921-113">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="cb921-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="cb921-114">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="cb921-114">Learn about API capabilities</span></span>

<span data-ttu-id="cb921-115">OneNote JavaScript API を使用して OneNote on the web のコンテンツと対話するには、「[OneNote アドインのクイック スタート](../../quickstarts/onenote-quickstart.md)」を実行します。</span><span class="sxs-lookup"><span data-stu-id="cb921-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="cb921-116">OneNote JavaScript API オブジェクト モデルの詳細については、[OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb921-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="cb921-117">関連項目</span><span class="sxs-lookup"><span data-stu-id="cb921-117">See also</span></span>

* [<span data-ttu-id="cb921-118">OneNote アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="cb921-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="cb921-119">OneNote アドインの概要</span><span class="sxs-lookup"><span data-stu-id="cb921-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="cb921-120">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="cb921-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="cb921-121">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="cb921-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
