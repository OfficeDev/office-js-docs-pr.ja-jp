---
title: OneNote JavaScript API の概要
description: ''
ms.date: 06/20/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 68ac6f94921ba3b1ea14f364988b57ef86809890
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127129"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="0a200-102">OneNote JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="0a200-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="0a200-103">適用対象: web 上の OneNote</span><span class="sxs-lookup"><span data-stu-id="0a200-103">Applies to: OneNote on the web</span></span>

<span data-ttu-id="0a200-104">以下のリンクは、API で使用できる高レベルの OneNote オブジェクトを示しています。</span><span class="sxs-lookup"><span data-stu-id="0a200-104">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="0a200-105">オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、イベント、メソッドの説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="0a200-105">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="0a200-106">リンクを参照して、詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="0a200-106">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="0a200-107">[Application](/javascript/api/onenote/onenote.application): グローバルにアドレス可能な OneNote オブジェクト (アクティブなノートブック、アクティブなセクションなど) すべてへのアクセスに使用する最上位のオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0a200-107">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="0a200-p102">[Notebook](/javascript/api/onenote/onenote.notebook): ノートブックです。ノートブックには、セクション グループとセクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="0a200-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection):ノートブックのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="0a200-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup):セクション グループです。セクション グループには、セクション グループとセクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="0a200-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection):セクション グループのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="0a200-p104">[Section](/javascript/api/onenote/onenote.section):セクションです。セクションには、ページが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="0a200-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection):セクションのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="0a200-p105">[Page](/javascript/api/onenote/onenote.page):ページです。ページには、PageContent オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="0a200-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection):ページのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="0a200-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent):Outline や Image などのコンテンツの種類を含むページの最上位の領域です。PageContent オブジェクトは、ページ上の位置を指定できます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="0a200-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection):PageContent オブジェクトのコレクションで、ページのコンテンツを表します。</span><span class="sxs-lookup"><span data-stu-id="0a200-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="0a200-p107">[Outline](/javascript/api/onenote/onenote.outline):Paragraph オブジェクトのコンテナーです。Outline は、PageContent オブジェクトの直接の子です。</span><span class="sxs-lookup"><span data-stu-id="0a200-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="0a200-p108">[Image](/javascript/api/onenote/onenote.image):Image オブジェクトです。Image は、PageContent オブジェクトまたは Paragraph の直接の子にすることができます。</span><span class="sxs-lookup"><span data-stu-id="0a200-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="0a200-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph):ページに表示されるコンテンツのコンテナーです。Paragraph は、Outline の直接の子です。</span><span class="sxs-lookup"><span data-stu-id="0a200-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="0a200-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection):Outline 内の Paragraph オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="0a200-130">[RichText](/javascript/api/onenote/onenote.richtext):RichText オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0a200-130">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="0a200-131">[Table](/javascript/api/onenote/onenote.table):TableRow オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="0a200-131">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="0a200-132">[TableRow](/javascript/api/onenote/onenote.tablerow):TableCell オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="0a200-132">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="0a200-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection):Table 内の TableRow オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="0a200-134">[TableCell](/javascript/api/onenote/onenote.tablecell):Paragraph オブジェクトのコンテナーです。</span><span class="sxs-lookup"><span data-stu-id="0a200-134">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="0a200-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): TableRow 内の TableCell オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0a200-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="0a200-136">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="0a200-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="0a200-137">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="0a200-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="0a200-138">Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="0a200-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="0a200-139">OneNote JavaScript API 要件セットの詳細については、「[OneNote JavaScript API の要件セット](../requirement-sets/onenote-api-requirement-sets.md)」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a200-139">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="0a200-140">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="0a200-140">OneNote JavaScript API reference</span></span>

<span data-ttu-id="0a200-141">OneNote JavaScript API の詳細については、[OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a200-141">For detailed information about the OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="0a200-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="0a200-142">See also</span></span>

- [<span data-ttu-id="0a200-143">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="0a200-143">OneNote JavaScript API programming overview</span></span>](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="0a200-144">最初の OneNote 用アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="0a200-144">Build your first OneNote add-in</span></span>](../../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="0a200-145">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="0a200-145">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="0a200-146">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="0a200-146">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
