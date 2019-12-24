---
title: OneNote JavaScript API の概要
description: ''
ms.date: 07/05/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 3bd90f1bea7c9b3e7f6689341247d66801357f85
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851524"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="f619e-102">OneNote JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="f619e-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="f619e-103">OneNote アドインは、次の2つの JavaScript オブジェクトモデルを含む JavaScript API for Office を使用して、web 上の OneNote のオブジェクトと対話します。</span><span class="sxs-lookup"><span data-stu-id="f619e-103">A OneNote add-in interacts with objects in OneNote on the web by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="f619e-104">**Onenote JAVASCRIPT api**: Office 2016 で導入された[onenote javascript api](/javascript/api/onenote)には、web 上の onenote のオブジェクトへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="f619e-104">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="f619e-105">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f619e-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="f619e-106">ドキュメントのこのセクションでは OneNote JavaScript API に重点を置いています。これは、web 上の OneNote を対象とするアドインで大部分の機能を開発するために使用します。</span><span class="sxs-lookup"><span data-stu-id="f619e-106">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="f619e-107">一般的な API の詳細については、「 [Office JAVASCRIPT API オブジェクトモデル](../../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f619e-107">For information about the Common API, see [Office JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="f619e-108">プログラミングの概念を学ぶ</span><span class="sxs-lookup"><span data-stu-id="f619e-108">Learn programming concepts</span></span>

<span data-ttu-id="f619e-109">重要なプログラミングの概念に関する詳細情報については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f619e-109">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="f619e-110">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="f619e-110">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="f619e-111">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="f619e-111">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="f619e-112">API 機能について学ぶ</span><span class="sxs-lookup"><span data-stu-id="f619e-112">Learn about API capabilities</span></span>

<span data-ttu-id="f619e-113">Onenote JavaScript API を使用して web 上の OneNote のコンテンツを操作する作業を行うには、 [onenote アドインのクイックスタート](../../quickstarts/onenote-quickstart.md)を完了します。</span><span class="sxs-lookup"><span data-stu-id="f619e-113">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="f619e-114">OneNote JavaScript API オブジェクトモデルの詳細については、 [Onenote JAVASCRIPT api リファレンスドキュメント](/javascript/api/onenote)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f619e-114">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="f619e-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="f619e-115">See also</span></span>

- [<span data-ttu-id="f619e-116">OneNote アドイン ドキュメント</span><span class="sxs-lookup"><span data-stu-id="f619e-116">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="f619e-117">OneNote アドインの概要</span><span class="sxs-lookup"><span data-stu-id="f619e-117">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="f619e-118">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f619e-118">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="f619e-119">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="f619e-119">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

