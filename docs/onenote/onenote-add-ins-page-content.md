---
title: OneNote ページ コンテンツを使用する
description: JavaScript API を使用してページ OneNoteを使用する方法について説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f506617bfdbc97e94f8fb16930dfc2a935385d5f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349050"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="082f3-103">OneNote ページ コンテンツを使用する</span><span class="sxs-lookup"><span data-stu-id="082f3-103">Work with OneNote page content</span></span>

<span data-ttu-id="082f3-104">OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。</span><span class="sxs-lookup"><span data-stu-id="082f3-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote オブジェクト モデル図を参照します。](../images/one-note-om-page.png)

- <span data-ttu-id="082f3-106">ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="082f3-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="082f3-107">PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="082f3-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="082f3-108">アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="082f3-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="082f3-109">Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。</span><span class="sxs-lookup"><span data-stu-id="082f3-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="082f3-110">空のページを作成OneNote、次のいずれかの方法を使用します。</span><span class="sxs-lookup"><span data-stu-id="082f3-110">To create an empty OneNote page, use one of the following methods.</span></span>

- [<span data-ttu-id="082f3-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="082f3-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="082f3-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="082f3-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="082f3-113">その後、次のオブジェクトのメソッドを使用して、`Page.addOutline` や `Outline.appendHtml` などのページ コンテンツを操作します。</span><span class="sxs-lookup"><span data-stu-id="082f3-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="082f3-114">Page</span><span class="sxs-lookup"><span data-stu-id="082f3-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="082f3-115">Outline</span><span class="sxs-lookup"><span data-stu-id="082f3-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="082f3-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="082f3-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="082f3-p101">OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="082f3-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="082f3-119">サポートされている HTML</span><span class="sxs-lookup"><span data-stu-id="082f3-119">Supported HTML</span></span>

<span data-ttu-id="082f3-120">このOneNote JavaScript API では、ページ コンテンツを作成および更新するための次の HTML がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="082f3-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content.</span></span>

- <span data-ttu-id="082f3-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="082f3-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="082f3-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="082f3-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="082f3-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="082f3-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="082f3-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="082f3-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="082f3-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="082f3-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="082f3-126">HTML を OneNote にインポートすると、空白文字が統合されます。</span><span class="sxs-lookup"><span data-stu-id="082f3-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="082f3-127">結果のコンテンツは、1 つのアウトラインに貼り付けられます。</span><span class="sxs-lookup"><span data-stu-id="082f3-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="082f3-128">OneNote では、ユーザーのセキュリティを確保しながら、HTML をページ コンテンツに変換します。</span><span class="sxs-lookup"><span data-stu-id="082f3-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="082f3-129">HTML と CSS の基準は OneNote のコンテンツ モデルと完全に一致しないため、特に CSS スタイルでは外観が異なります。</span><span class="sxs-lookup"><span data-stu-id="082f3-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="082f3-130">特定の書式設定が必要な場合は、JavaScript オブジェクトを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="082f3-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="082f3-131">ページ コンテンツへのアクセス</span><span class="sxs-lookup"><span data-stu-id="082f3-131">Accessing page contents</span></span>

<span data-ttu-id="082f3-p104">現在アクティブなページの `Page#load` による *ページ コンテンツ* へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="082f3-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="082f3-134">タイトルなどのメタデータは、どのページでも照会できます。</span><span class="sxs-lookup"><span data-stu-id="082f3-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="082f3-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="082f3-135">See also</span></span>

- [<span data-ttu-id="082f3-136">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="082f3-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="082f3-137">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="082f3-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="082f3-138">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="082f3-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="082f3-139">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="082f3-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
