---
title: OneNote ページ コンテンツを使用する
description: JavaScript API を使用して OneNote ページコンテンツを操作する方法について説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ec8a6a92bf6bf58fac9c3c2d22987bc027414
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720940"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="16ea4-103">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="16ea4-103">Work with OneNote page content</span></span>

<span data-ttu-id="16ea4-104">OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。</span><span class="sxs-lookup"><span data-stu-id="16ea4-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote ページのオブジェクト モデル図](../images/one-note-om-page.png)

- <span data-ttu-id="16ea4-106">ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="16ea4-107">PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="16ea4-108">アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="16ea4-109">Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="16ea4-110">空の OneNote ページを作成するには、次の方法のいずれかを使用します。</span><span class="sxs-lookup"><span data-stu-id="16ea4-110">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="16ea4-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="16ea4-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="16ea4-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="16ea4-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="16ea4-113">その後、次のオブジェクトのメソッドを使用して、`Page.addOutline` や `Outline.appendHtml` などのページ コンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="16ea4-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="16ea4-114">Page</span><span class="sxs-lookup"><span data-stu-id="16ea4-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="16ea4-115">Outline</span><span class="sxs-lookup"><span data-stu-id="16ea4-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="16ea4-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="16ea4-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="16ea4-p101">OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="16ea4-119">サポートされている HTML</span><span class="sxs-lookup"><span data-stu-id="16ea4-119">Supported HTML</span></span>

<span data-ttu-id="16ea4-120">ページ コンテンツを作成して更新するために、OneNote アドインの JavaScript API では次の HTML がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="16ea4-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="16ea4-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="16ea4-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="16ea4-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="16ea4-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="16ea4-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="16ea4-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="16ea4-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="16ea4-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="16ea4-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="16ea4-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="16ea4-126">HTML を OneNote にインポートすると、空白文字が統合されます。</span><span class="sxs-lookup"><span data-stu-id="16ea4-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="16ea4-127">結果のコンテンツは、1 つのアウトラインに貼り付けられます。</span><span class="sxs-lookup"><span data-stu-id="16ea4-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="16ea4-128">OneNote では、ユーザーのセキュリティを確保しながら、HTML をページ コンテンツに変換します。</span><span class="sxs-lookup"><span data-stu-id="16ea4-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="16ea4-129">HTML と CSS の基準は OneNote のコンテンツ モデルと完全に一致しないため、特に CSS スタイルでは外観が異なります。</span><span class="sxs-lookup"><span data-stu-id="16ea4-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="16ea4-130">特定の書式設定が必要な場合は、JavaScript オブジェクトを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="16ea4-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="16ea4-131">ページ コンテンツへのアクセス</span><span class="sxs-lookup"><span data-stu-id="16ea4-131">Accessing page contents</span></span>

<span data-ttu-id="16ea4-p104">現在アクティブなページの `Page#load` による*ページ コンテンツ*へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="16ea4-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="16ea4-134">タイトルなどのメタデータは、どのページでも照会できます。</span><span class="sxs-lookup"><span data-stu-id="16ea4-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="16ea4-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="16ea4-135">See also</span></span>

- [<span data-ttu-id="16ea4-136">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="16ea4-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="16ea4-137">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="16ea4-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="16ea4-138">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="16ea4-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="16ea4-139">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="16ea4-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
