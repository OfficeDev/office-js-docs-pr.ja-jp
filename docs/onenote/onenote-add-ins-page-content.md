---
title: OneNote ページ コンテンツを使用する
description: JavaScript API を使用してページ OneNoteを使用する方法について説明します。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 72a5402d16f8d8a39903b3285c62ade48a409578
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154288"
---
# <a name="work-with-onenote-page-content"></a>OneNote ページ コンテンツを使用する

OneNote アドインの JavaScript API では、ページ コンテンツは次のようなオブジェクト モデルで表されます。

  ![OneNote オブジェクト モデル図を参照します。](../images/one-note-om-page.png)

- ページ オブジェクトには、PageContent オブジェクトのコレクションが含まれています。
- PageContent オブジェクトには、アウトライン、イメージ、その他のコンテンツ タイプが含まれています。
- アウトライン オブジェクトには、Paragraph オブジェクトのコレクションが含まれています。
- Paragraph オブジェクトには、RichText、Image、Table、Other のコンテンツ タイプが含まれています。

空のページを作成OneNote、次のいずれかの方法を使用します。

- [Section.addPage](/javascript/api/onenote/onenote.section#addPage_title_)
- [Page.insertPageAsSibling](/javascript/api/onenote/onenote.section#insertSectionAsSibling_location__title_)

その後、次のオブジェクトのメソッドを使用して、`Page.addOutline` や `Outline.appendHtml` などのページ コンテンツを操作します。

- [Page](/javascript/api/onenote/onenote.page)
- [Outline](/javascript/api/onenote/onenote.outline)
- [Paragraph](/javascript/api/onenote/onenote.paragraph)

OneNote ページのコンテンツと構造は、HTML で表されます。次に説明するように、ページ コンテンツの作成や更新には、HTML のサブセットだけがサポートされています。

## <a name="supported-html"></a>サポートされている HTML

このOneNote JavaScript API では、ページ コンテンツを作成および更新するための次の HTML がサポートされています。

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

現在アクティブなページの `Page#load` による *ページ コンテンツ* へのアクセスだけが可能です。アクティブなページを変更するには、`navigateToPage($page)` を呼び出します。

タイトルなどのメタデータは、どのページでも照会できます。

## <a name="see-also"></a>関連項目

- [OneNote の JavaScript API のプログラミングの概要](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
