---
title: OneNote ページ コンテンツを使用する
description: ''
ms.date: 1/10/2019
ms.openlocfilehash: 617c30f2a9a0c72b1c309ce299f388b5a16b983f
ms.sourcegitcommit: 384e217fd51d73d13ccfa013bfc6e049b66bd98c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/11/2019
ms.locfileid: "27896337"
---
# <a name="work-with-onenote-page-content"></a>OneNote ページ コンテンツを使用する

OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。

  ![OneNote ページのオブジェクト モデル図](../images/one-note-om-page.png)

- ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。
- PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。
- アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。
- Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。

空の OneNote ページを作成するには、次の方法のいずれかを使用します。

- [Section.addPage](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [Page.insertPageAsSibling](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

その後、次のオブジェクトのメソッドを使用して、`Page.addOutline` や `Outline.appendHtml` などのページ コンテンツを操作します。

- [Page](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [Outline](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [Paragraph](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。

## <a name="supported-html"></a>サポートされている HTML

ページ コンテンツを作成して更新するために、OneNote アドインの JavaScript API では次の HTML がサポートされています。

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>`
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>`
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

> [!NOTE]
> HTML を OneNote にインポートすると、空白文字が統合されます。 結果のコンテンツは、1 つのアウトラインに貼り付けられます。

OneNote では、ユーザーのセキュリティを確保しながら、HTML をページ コンテンツに変換します。 HTML と CSS の基準は OneNote のコンテンツ モデルと完全に一致しないため、特に CSS スタイルでは外観が異なります。 特定の書式設定が必要な場合は、JavaScript オブジェクトを使用することをお勧めします。

## <a name="accessing-page-contents"></a>ページ コンテンツへのアクセス

現在アクティブなページの `Page#load` による*ページ コンテンツ*へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。

タイトルなどのメタデータは、どのページでも照会できます。

## <a name="see-also"></a>関連項目

- [OneNote の JavaScript API のプログラミングの概要](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
