---
title: OneNote ページ コンテンツを使用する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 3ceb693b85490e5b7046880a79ae46753a1d3238
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944128"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="72206-102">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="72206-102">Work with OneNote page content</span></span> 

<span data-ttu-id="72206-103">OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。</span><span class="sxs-lookup"><span data-stu-id="72206-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote ページのオブジェクト モデル図](../images/one-note-om-page.png)

- <span data-ttu-id="72206-105">ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="72206-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="72206-106">PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="72206-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="72206-107">アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="72206-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="72206-108">Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="72206-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="72206-109">空の OneNote ページを作成するには、次の方法のいずれかを使用します。</span><span class="sxs-lookup"><span data-stu-id="72206-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="72206-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="72206-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#addpage-title-)
- [<span data-ttu-id="72206-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="72206-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section?view=office-js#insertsectionassibling-location--title-)

<span data-ttu-id="72206-112">その後、次のオブジェクトのメソッドを使用して、Page.addOutline や Outline.appendHtml などのページのコンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="72206-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="72206-113">ページ</span><span class="sxs-lookup"><span data-stu-id="72206-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page?view=office-js)
- [<span data-ttu-id="72206-114">アウトライン</span><span class="sxs-lookup"><span data-stu-id="72206-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline?view=office-js)
- [<span data-ttu-id="72206-115">段落</span><span class="sxs-lookup"><span data-stu-id="72206-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph?view=office-js)

<span data-ttu-id="72206-p101">OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="72206-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="72206-118">サポートされている HTML</span><span class="sxs-lookup"><span data-stu-id="72206-118">Supported HTML</span></span>

<span data-ttu-id="72206-119">ページ コンテンツを作成して更新するために、OneNote アドインの JavaScript API では次の HTML がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="72206-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="72206-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="72206-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="72206-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="72206-121">`<ul>`, `<ol>`, `<li>`</span></span> 
- <span data-ttu-id="72206-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="72206-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="72206-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="72206-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="72206-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="72206-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="72206-125">ページ コンテンツへのアクセス</span><span class="sxs-lookup"><span data-stu-id="72206-125">Accessing page contents</span></span>

<span data-ttu-id="72206-p102">現在アクティブなページの `Page#load` による*ページ コンテンツ*へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="72206-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="72206-128">タイトルなどのメタデータは、どのページでも照会できます。</span><span class="sxs-lookup"><span data-stu-id="72206-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="72206-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="72206-129">See also</span></span>

- [<span data-ttu-id="72206-130">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="72206-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="72206-131">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="72206-131">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/onenote-add-ins-javascript-reference?view=office-js)
- [<span data-ttu-id="72206-132">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="72206-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="72206-133">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="72206-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
