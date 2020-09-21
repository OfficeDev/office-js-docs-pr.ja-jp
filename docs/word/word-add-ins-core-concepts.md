---
title: Office アドインの Word JavaScript オブジェクト モデル
description: Word 固有の JavaScript オブジェクト モデルの最も重要なクラスについて説明します。
ms.date: 09/04/2020
localization_priority: Priority
ms.openlocfilehash: 7424ee83bde0c19a574233c64811ecbb55763d93
ms.sourcegitcommit: 0844ca7589ad3a6b0432fe126ca4e0ac9dbb80ce
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/18/2020
ms.locfileid: "47963840"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="31707-103">Office アドインの Word JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="31707-103">Word JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="31707-104">この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用してアドインを構築するための基本的な概念について説明します。API を使用するための基本的なコア コンセプトを紹介します。</span><span class="sxs-lookup"><span data-stu-id="31707-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins. It introduces core concepts that are fundamental to using the API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="31707-105">Word API の非同期性と、ドキュメントでの動作方法については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31707-105">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.</span></span>

## <a name="officejs-apis-for-word"></a><span data-ttu-id="31707-106">Word 用の Office.js API</span><span class="sxs-lookup"><span data-stu-id="31707-106">Office.js APIs for Word</span></span>

<span data-ttu-id="31707-107">Word アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します:</span><span class="sxs-lookup"><span data-stu-id="31707-107">A Word add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="31707-108">**Word JavaScript API**: [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) には、ドキュメント、範囲、テーブル、リスト、フォーマットなどにアクセスするために使用できる厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="31707-108">**Word JavaScript API**: The [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access the document, ranges, tables, lists, formatting, and more.</span></span>

* <span data-ttu-id="31707-109">**共通 API**: [共通 API](/javascript/api/office) を使用して、UI、ダイアログ、クライアント設定など、複数のタイプの Office アプリケーションに共通の機能にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="31707-109">**Common APIs**: The [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="31707-110">Word を対象にしたアドインでは、機能の大部分を Word JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。</span><span class="sxs-lookup"><span data-stu-id="31707-110">While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API.</span></span> <span data-ttu-id="31707-111">例:</span><span class="sxs-lookup"><span data-stu-id="31707-111">For example:</span></span>

* <span data-ttu-id="31707-112">[コンテキスト](/javascript/api/office/office.context): `Context` オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="31707-112">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="31707-113">これは `contentLanguage` や `officeTheme` などのドキュメント構成の詳細で構成され、`host` や `platform` などのアドインのランタイム環境に関する情報も提供します。</span><span class="sxs-lookup"><span data-stu-id="31707-113">It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="31707-114">さらに、`requirements.isSetSupported()` メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="31707-114">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="31707-115">[ドキュメント](/javascript/api/office/office.document): `Document` オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Word ファイルをダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="31707-115">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running.</span></span>

![Word JS API と共通 API の違いを示す画像](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a><span data-ttu-id="31707-117">Word 固有のオブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="31707-117">Word-specific object model</span></span>

<span data-ttu-id="31707-118">Word API について理解するには、ドキュメントの構成要素が互いにどのように関連しているかを理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="31707-118">To understand the Word APIs, you must understand how the components of a document are related to one another.</span></span>

* <span data-ttu-id="31707-119">**ドキュメント** には **セクション** と、設定やカスタム XML パーツなどのドキュメントレベルのエンティティが含まれます。</span><span class="sxs-lookup"><span data-stu-id="31707-119">The **Document** contains the **Section**s, and document-level entities such as settings and custom XML parts.</span></span>
* <span data-ttu-id="31707-120">**セクション** には **本文** が含まれます。</span><span class="sxs-lookup"><span data-stu-id="31707-120">A **Section** contains a **Body**.</span></span>
* <span data-ttu-id="31707-121">**本文** は、特に **パラグラフ**、**ContentControl**、および **範囲** オブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="31707-121">A **Body** gives access to **Paragraph**s, **ContentControl**s, and **Range** objects, among others.</span></span>
* <span data-ttu-id="31707-122">**範囲** は、テキスト、空白、**テーブル**、画像など、コンテンツの連続した領域を表します。</span><span class="sxs-lookup"><span data-stu-id="31707-122">A **Range** represents a contiguous area of content, including text, white space, **Table**s, and images.</span></span> <span data-ttu-id="31707-123">また、テキストの操作方法のほとんどが含まれます。</span><span class="sxs-lookup"><span data-stu-id="31707-123">It also contains most of the text manipulation methods.</span></span>
* <span data-ttu-id="31707-124">**リスト** は、番号付きまたは箇条書きのリスト内のテキストを表します。</span><span class="sxs-lookup"><span data-stu-id="31707-124">A **List** represents text in a numbered or bulleted list.</span></span>

## <a name="see-also"></a><span data-ttu-id="31707-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="31707-125">See also</span></span>

- [<span data-ttu-id="31707-126">Word JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="31707-126">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="31707-127">最初の Word アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="31707-127">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="31707-128">Word アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="31707-128">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="31707-129">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="31707-129">Word JavaScript API reference</span></span>](/javascript/api/word)
