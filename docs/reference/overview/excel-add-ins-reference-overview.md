---
title: Excel JavaScript API の概要
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: bf1d4642a7ceeb34eab51722a398887bb5c03fec
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450171"
---
# <a name="excel-javascript-api-overview"></a>Excel JavaScript API の概要

Excel の JavaScript API を使用して、Excel 2016 以降のアドインをビルドします。 API で使用できる Excel オブジェクトの概要を次に示します。 オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、イベント、メソッドの説明が含まれています。 メニューからのリンクを調べて、詳細を確認してください。

便宜上、Excel の主要なオブジェクトの一部を以下に示します。 

- [ブック](/javascript/api/excel/excel.workbook): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。

- [Worksheet](/javascript/api/excel/excel.worksheet):ブック内のワークシートを表します。 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): ブック内の **Worksheet** オブジェクトのコレクション。
    - [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): **Worksheet** オブジェクトの保護を表します。

- [Range](/javascript/api/excel/excel.range): 1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。
    - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat): ルールの条件が満たされたときに範囲に適用されるルールと形式を定義するオブジェクトです。
    - [DataValidation](/javascript/api/excel/excel.datavalidation): さまざまな基準に基づいて範囲へのユーザー入力を制限するオブジェクトです。
    - [RangeSort](/javascript/api/excel/excel.rangesort): 範囲の並べ替え操作を管理するオブジェクトを表します。

- [Table](/javascript/api/excel/excel.table): データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。
    - [TableCollection](/javascript/api/excel/excel.tablecollection):ブックまたはワークシート内のテーブルのコレクション。
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection):テーブル内のすべての列のコレクション。
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): テーブル内のすべての行のコレクションです。
    - [TableSort](/javascript/api/excel/excel.tablesort): テーブルの並べ替え操作を管理するオブジェクトを表します。

- [Chart](/javascript/api/excel/excel.chart): 基になるデータを視覚的に表示する、ワークシート内の Chart オブジェクトを表します。
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): ワークシート内のグラフのコレクションです。
    
- [PivotTable](/javascript/api/excel/excel.pivottable): データの階層型のグループ化とプレゼンテーションを行う Excel のピボットテーブルを表します。 
    - [PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): ワークシート内のピボットテーブルのコレクションです。

- [Filter](/javascript/api/excel/excel.filter): テーブルの列のフィルター処理を管理するオブジェクトを表します。

- [NamedItem](/javascript/api/excel/excel.nameditem): セルまたは値の範囲の定義済みの名前を表します。 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection):ブック内の **NamedItem** オブジェクトのコレクション。

- [バインド](/javascript/api/excel/excel.binding): ブックのセクションへのバインドを表す抽象クラス。
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): ブック内の **Binding** オブジェクトのコレクションです。

## <a name="excel-javascript-api-open-specifications"></a>Excel JavaScript API オープン仕様

新しい Excel アドイン用の API の設計と開発にあたり、[Open API の仕様](../openspec.md) ページでこれらに対するフィードバックの提供が可能になります。 Excel JavaScript API 用のパイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。 Excel JavaScript API 要件セットの詳細については、「[Excel JavaScript API の要件セット](../requirement-sets/excel-api-requirement-sets.md)」の記事を参照してください。

## <a name="excel-javascript-api-reference"></a>Excel JavaScript API リファレンス

Excel JavaScript API の詳細については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。

## <a name="see-also"></a>関連項目

- [Excel アドインの概要](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office アドイン プラットフォームの概要](/office/dev/add-ins/overview/office-add-ins)
- [GitHub の Excel アドインのサンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
