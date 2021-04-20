---
title: Excel JavaScript API の概要
description: Excel JavaScript API の詳細情報
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293661"
---
# <a name="excel-javascript-api-overview"></a>Excel JavaScript API の概要

Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。

* **Excel JavaScript API**: これは、Excel 用の [アプリケーション固有 API](../../develop/application-specific-api-model.md) です。 Office 2016 で導入された [Excel JavaScript API](/javascript/api/excel) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。

* **共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。

ドキュメントのこのセクションでは、Excel JavaScript API に焦点を当てて、そしてそれを Excel on the web または Excel 2016 以降を対象としたアドインの大部分の機能開発に使用します。 共通 API の詳細については、「[共通 JavaScript API オブジェクト モデル](../../develop/office-javascript-api-object-model.md)」を参照してください。

## <a name="learn-programming-concepts"></a>プログラミングの概念を学ぶ

重要なプログラミング概念の詳細については、「[Excel JavaScript API を使用した基本的なプログラミングの概念](../../excel/excel-add-ins-core-concepts.md)」を参照してください。

Excel JavaScript API を使用して Excel のオブジェクトにアクセスするための実践的なエクスペリエンスに関しては、「[Excel アドインのチュートリアル](../../tutorials/excel-tutorial.md)」を完了してください。

## <a name="learn-api-capabilities"></a>API 機能について

主要な Excel API 機能にはそれぞれ、その機能が実行できることと関連するオブジェクト モデルについての記事があります。

* [グラフ](../../excel/excel-add-ins-charts.md)
* [コメント](../../excel/excel-add-ins-comments.md)
* [条件付き書式](../../excel/excel-add-ins-conditional-formatting.md)
* [カスタム関数](../../excel/custom-functions-overview.md)
* [データ検証](../../excel/excel-add-ins-data-validation.md)
* [イベント](../../excel/excel-add-ins-events.md)
* [複数の範囲 (範囲領域)](../../excel/excel-add-ins-multiple-ranges.md)
* [ピボットテーブル](../../excel/excel-add-ins-pivottables.md)
* [範囲](../../excel/excel-add-ins-ranges.md) および [高度な範囲 API](../../excel/excel-add-ins-ranges-advanced.md)
* [図形](../../excel/excel-add-ins-shapes.md)
* [表](../../excel/excel-add-ins-tables.md)
* [ブックとアプリケーションレベルの API](../../excel/excel-add-ins-workbooks.md)
* [ワークシート](../../excel/excel-add-ins-worksheets.md)

Excel JavaScript API オブジェクト モデルに関する詳細情報については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。

## <a name="try-out-code-samples-in-script-lab"></a>Script Lab でコード サンプルを試してみる

[Script Lab](../../overview/explore-with-script-lab.md) を使用すると、API を使用してタスクを完了する方法を示す組み込みのサンプルのコレクションを使用して操作をすぐに開始できます。 Script Lab のサンプルを実行すると、作業ウィンドウまたはワークシートですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。

## <a name="see-also"></a>関連項目

* [Excel アドイン ドキュメント](../../excel/index.yml)
* [Excel アドインの概要](../../excel/excel-add-ins-overview.md)
* [Excel JavaScript API リファレンス](/javascript/api/excel)
* [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](../../overview/office-add-in-availability.md)
* [アプリケーション固有の API モデルの使用](../../develop/application-specific-api-model.md)
