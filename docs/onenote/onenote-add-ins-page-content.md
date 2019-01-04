---
title: OneNote ページ コンテンツを使用する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: aef9d80ebb37dacd2c3b5f2ec9d33cb0164d8452
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457615"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="24824-102">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="24824-102">Work with OneNote page content</span></span> 

<span data-ttu-id="24824-103">OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。</span><span class="sxs-lookup"><span data-stu-id="24824-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote ページのオブジェクト モデル図](../images/one-note-om-page.png)

- <span data-ttu-id="24824-105">ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="24824-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="24824-106">PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="24824-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="24824-107">アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="24824-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="24824-108">Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="24824-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="24824-109">空の OneNote ページを作成するには、次の方法のいずれかを使用します。</span><span class="sxs-lookup"><span data-stu-id="24824-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="24824-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="24824-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="24824-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="24824-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="24824-112">その後、次のオブジェクトのメソッドを使用して、Page.addOutline や Outline.appendHtml などのページのコンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="24824-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="24824-113">Page</span><span class="sxs-lookup"><span data-stu-id="24824-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="24824-114">Outline</span><span class="sxs-lookup"><span data-stu-id="24824-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="24824-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="24824-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="24824-p101">OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="24824-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="24824-118">サポートされている HTML</span><span class="sxs-lookup"><span data-stu-id="24824-118">Supported HTML</span></span>

<span data-ttu-id="24824-119">ページ コンテンツを作成して更新するために、OneNote アドインの JavaScript API では次の HTML がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="24824-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="24824-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="24824-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="24824-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="24824-121"></span></span> 
- <span data-ttu-id="24824-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="24824-122"></span></span>
- <span data-ttu-id="24824-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="24824-123"></span></span>
- <span data-ttu-id="24824-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="24824-124"></span></span>

> [!NOTE]
> <span data-ttu-id="24824-125">HTML を OneNote にインポートすると、空白文字が統合されます。</span><span class="sxs-lookup"><span data-stu-id="24824-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="24824-126">結果のコンテンツは、1 つのアウトラインに貼り付けられます。</span><span class="sxs-lookup"><span data-stu-id="24824-126">The resulting content is pasted into one outline.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="24824-127">ページ コンテンツへのアクセス</span><span class="sxs-lookup"><span data-stu-id="24824-127">Accessing page contents</span></span>

<span data-ttu-id="24824-p103">現在アクティブなページの `Page#load` による*ページ コンテンツ*へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="24824-p103">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="24824-130">タイトルなどのメタデータは、どのページでも照会できます。</span><span class="sxs-lookup"><span data-stu-id="24824-130">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="24824-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="24824-131">See also</span></span>

- [<span data-ttu-id="24824-132">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="24824-132">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="24824-133">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="24824-133">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="24824-134">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="24824-134">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="24824-135">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="24824-135">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
