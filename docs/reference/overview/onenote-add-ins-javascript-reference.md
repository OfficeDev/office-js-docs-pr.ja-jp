---
title: OneNote JavaScript API の概要
description: ''
ms.date: 10/09/2018
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: f8fed0104412f60ec59146ef7820be958047d1f3
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052743"
---
# <a name="onenote-javascript-api-overview"></a>OneNote JavaScript API の概要

適用対象: OneNote Online

以下のリンクは、API で使用できる高レベルの OneNote オブジェクトを示しています。 オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、イベント、メソッドの説明が含まれています。 リンクを参照して、詳細を確認してください。 
    
- [Application](/javascript/api/onenote/onenote.application): グローバルにアドレス可能な OneNote オブジェクト (アクティブなノートブック、アクティブなセクションなど) すべてへのアクセスに使用する最上位のオブジェクトです。

- [Notebook](/javascript/api/onenote/onenote.notebook): ノートブックです。ノートブックには、セクション グループとセクションが含まれます。
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection):ノートブックのコレクションです。

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup):セクション グループです。セクション グループには、セクション グループとセクションが含まれます。
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection):セクション グループのコレクションです。

- [Section](/javascript/api/onenote/onenote.section):セクションです。セクションには、ページが含まれます。
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection):セクションのコレクションです。

- [Page](/javascript/api/onenote/onenote.page):ページです。ページには、PageContent オブジェクトが含まれます。
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection):ページのコレクションです。

- [PageContent](/javascript/api/onenote/onenote.pagecontent):Outline や Image などのコンテンツの種類を含むページの最上位の領域です。PageContent オブジェクトは、ページ上の位置を指定できます。
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection):PageContent オブジェクトのコレクションで、ページのコンテンツを表します。

- [Outline](/javascript/api/onenote/onenote.outline):Paragraph オブジェクトのコンテナーです。Outline は、PageContent オブジェクトの直接の子です。

- [Image](/javascript/api/onenote/onenote.image):Image オブジェクトです。Image は、PageContent オブジェクトまたは Paragraph の直接の子にすることができます。

- [Paragraph](/javascript/api/onenote/onenote.paragraph):ページに表示されるコンテンツのコンテナーです。Paragraph は、Outline の直接の子です。
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection):Outline 内の Paragraph オブジェクトのコレクションです。

- [RichText](/javascript/api/onenote/onenote.richtext):RichText オブジェクトです。

- [Table](/javascript/api/onenote/onenote.table):TableRow オブジェクトのコンテナーです。

- [TableRow](/javascript/api/onenote/onenote.tablerow):TableCell オブジェクトのコンテナーです。
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection):Table 内の TableRow オブジェクトのコレクションです。
 
- [TableCell](/javascript/api/onenote/onenote.tablecell):Paragraph オブジェクトのコンテナーです。
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): TableRow 内の TableCell オブジェクトのコレクションです。

## <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。 OneNote JavaScript API 要件セットの詳細については、「[OneNote JavaScript API の要件セット](../requirement-sets/onenote-api-requirement-sets.md)」の記事を参照してください。

## <a name="onenote-javascript-api-reference"></a>OneNote JavaScript API リファレンス

OneNote JavaScript API の詳細については、[OneNote JavaScript API リファレンス ドキュメント](/javascript/api/onenote)に関するページを参照してください。

## <a name="see-also"></a>関連項目

- [OneNote の JavaScript API のプログラミングの概要](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [最初の OneNote 用アドインをビルドする](../../quickstarts/onenote-quickstart.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
