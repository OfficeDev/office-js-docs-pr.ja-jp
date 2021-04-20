---
title: Word JavaScript API の概要
description: Word JavaScript API の概要
ms.date: 07/28/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: a3bc6e1bc19fdc149506301068969366fb141e52
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293626"
---
# <a name="word-javascript-api-overview"></a>Word JavaScript API の概要

Word アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Word のオブジェクトを操作します。

* **Word JavaScript API**: これは、Word 用の [アプリケーション固有 API](../../develop/application-specific-api-model.md) です。 Office 2016 で導入された [Word JavaScript API](/javascript/api/word) には、Word 文書内のオブジェクトとメタデータへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。

* **共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。

ドキュメントのこのセクションでは、Word JavaScript API に焦点を当てて、そしてそれを Word on the web または Word 2016 以降を対象としたアドインの大部分の機能開発に使用します。 共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。

## <a name="learn-programming-concepts"></a>プログラミングの概念を学ぶ

重要なプログラミング概念の詳細については、「[Word JavaScript API を使用した基本的なプログラミングの概念](../../word/word-add-ins-core-concepts.md)」を参照してください。

## <a name="learn-about-api-capabilities"></a>API 機能について学ぶ

ドキュメントのこのセクションに記載されている他の記事を参照すると、[アドインからドキュメント全体を取得する](../../word/get-the-whole-document-from-an-add-in-for-word.md)方法、[検索オプションを使用して Word アドインでテキストを検索する](../../word/search-option-guidance.md)方法などを学習できます。 すべての提供可能な記事の一覧については、目次でご確認ください。

Word JavaScript API を使用して Word のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Word アドインのチュートリアル](../../tutorials/word-tutorial.md)」を完了してください。

Word JavaScript API オブジェクト モデルの詳細については、[Word JavaScript API リファレンス ドキュメント](/javascript/api/word)に関するページを参照してください。

## <a name="try-out-code-samples-in-script-lab"></a>Script Lab でコード サンプルを試してみる

[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。 Script Lab のサンプルを実行すると、作業ウィンドウまたはドキュメントですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。

## <a name="see-also"></a>関連項目

* [Word アドイン ドキュメント](../../word/index.yml)
* [Word アドインの概要](../../word/word-add-ins-programming-overview.md)
* [Word JavaScript API リファレンス](/javascript/api/word)
* [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](../../overview/office-add-in-availability.md)
