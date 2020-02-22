---
title: Word JavaScript API の概要
description: ''
ms.date: 02/19/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 90dd7c787086a67dd8607479bbc46c957192d5c3
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163970"
---
# <a name="word-javascript-api-overview"></a>Word JavaScript API の概要

Word アドインは、次の 2 つの JavaScript オブジェクト モデルを含む JavaScript API for Office を使用して、Word のオブジェクトを操作します。

* **Word JavaScript API**: Office 2016 で導入された [Word JavaScript API](/javascript/api/word) には、Word 文書内のオブジェクトとメタデータへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。 

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

- [Word アドイン ドキュメント](../../word/index.md)
- [Word アドインの概要](../../word/word-add-ins-programming-overview.md)
- [Word JavaScript API リファレンス](/javascript/api/word)
- [Office アドインのホストとプラットフォームの可用性](../../overview/office-add-in-availability.md)
- [API オープン仕様](../openspec/openspec.md)
