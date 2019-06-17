---
title: Excel JavaScript API の概要
description: ''
ms.date: 06/10/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa9574a93252c0011b211c39e37cc013beb64432
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910148"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="c1a73-102">Excel JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="c1a73-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="c1a73-103">Excel の JavaScript API を使用して、Excel 2016 以降のアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="c1a73-103">You can use the Excel JavaScript API to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="c1a73-104">API で使用できる Excel オブジェクトの概要を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-104">The following list shows the high-level Excel objects that are available in the API.</span></span> <span data-ttu-id="c1a73-105">オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、イベント、メソッドの説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c1a73-105">Each object page link contains a description of the properties, events, and methods that are available on the object.</span></span> <span data-ttu-id="c1a73-106">メニューからのリンクを調べて、詳細を確認してください。</span><span class="sxs-lookup"><span data-stu-id="c1a73-106">Explore the links from the menu to learn more.</span></span>

<span data-ttu-id="c1a73-107">便宜上、Excel の主要なオブジェクトの一部を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-107">Some of the core Excel objects are listed below for convenience:</span></span>

- <span data-ttu-id="c1a73-108">[ブック](/javascript/api/excel/excel.workbook): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。</span><span class="sxs-lookup"><span data-stu-id="c1a73-108">[Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.</span></span>

- <span data-ttu-id="c1a73-109">[Worksheet](/javascript/api/excel/excel.worksheet):ブック内のワークシートを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-109">[Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook.</span></span>
  - <span data-ttu-id="c1a73-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): ブック内の **Worksheet** オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="c1a73-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.</span></span>
  - <span data-ttu-id="c1a73-111">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): **Worksheet** オブジェクトの保護を表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-111">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.</span></span>

- <span data-ttu-id="c1a73-112">[Range](/javascript/api/excel/excel.range): 1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-112">[Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.</span></span>
  - <span data-ttu-id="c1a73-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat): ルールの条件が満たされたときに範囲に適用されるルールと形式を定義するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat): An object defining a rule and a format applied to the range when the rule's condition is met.</span></span>
  - <span data-ttu-id="c1a73-114">[DataValidation](/javascript/api/excel/excel.datavalidation): さまざまな基準に基づいて範囲へのユーザー入力を制限するオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-114">[DataValidation](/javascript/api/excel/excel.datavalidation): An object that restricts user input to a range based on a variety of criteria.</span></span>
  - <span data-ttu-id="c1a73-115">[RangeSort](/javascript/api/excel/excel.rangesort): 範囲の並べ替え操作を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-115">[RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on a range.</span></span>

- <span data-ttu-id="c1a73-116">[Table](/javascript/api/excel/excel.table): データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-116">[Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.</span></span>
  - <span data-ttu-id="c1a73-117">[TableCollection](/javascript/api/excel/excel.tablecollection):ブックまたはワークシート内のテーブルのコレクション。</span><span class="sxs-lookup"><span data-stu-id="c1a73-117">[TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.</span></span>
  - <span data-ttu-id="c1a73-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection):テーブル内のすべての列のコレクション。</span><span class="sxs-lookup"><span data-stu-id="c1a73-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.</span></span>
  - <span data-ttu-id="c1a73-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): テーブル内のすべての行のコレクションです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.</span></span>
  - <span data-ttu-id="c1a73-120">[TableSort](/javascript/api/excel/excel.tablesort): テーブルの並べ替え操作を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-120">[TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on a table.</span></span>

- <span data-ttu-id="c1a73-121">[Chart](/javascript/api/excel/excel.chart): 基になるデータを視覚的に表示する、ワークシート内の Chart オブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-121">[Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.</span></span>
  - <span data-ttu-id="c1a73-122">[ChartCollection](/javascript/api/excel/excel.chartcollection): ワークシート内のグラフのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-122">[ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.</span></span>

- <span data-ttu-id="c1a73-123">[PivotTable](/javascript/api/excel/excel.pivottable): データの階層型のグループ化とプレゼンテーションを行う Excel のピボットテーブルを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-123">[PivotTable](/javascript/api/excel/excel.pivottable): Represents an Excel PivotTable, which is a hierarchical grouping and presentation of data.</span></span>
  - <span data-ttu-id="c1a73-124">[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): ワークシート内のピボットテーブルのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-124">[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): A collection of PivotTables in a worksheet.</span></span>

- <span data-ttu-id="c1a73-125">[Filter](/javascript/api/excel/excel.filter): テーブルの列のフィルター処理を管理するオブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-125">[Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.</span></span>

- <span data-ttu-id="c1a73-126">[NamedItem](/javascript/api/excel/excel.nameditem): セルまたは値の範囲の定義済みの名前を表します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-126">[NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value.</span></span>
  - <span data-ttu-id="c1a73-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection):ブック内の **NamedItem** オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="c1a73-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.</span></span>

- <span data-ttu-id="c1a73-128">[バインド](/javascript/api/excel/excel.binding): ブックのセクションへのバインドを表す抽象クラス。</span><span class="sxs-lookup"><span data-stu-id="c1a73-128">[Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.</span></span>
  - <span data-ttu-id="c1a73-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection): ブック内の **Binding** オブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.</span></span>

## <a name="excel-javascript-api-requirement-sets"></a><span data-ttu-id="c1a73-130">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="c1a73-130">Excel JavaScript API requirement sets</span></span>

<span data-ttu-id="c1a73-131">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="c1a73-131">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="c1a73-132">Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="c1a73-132">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="c1a73-133">Excel JavaScript API 要件セットの詳細については、「[Excel JavaScript API の要件セット](../requirement-sets/excel-api-requirement-sets.md)」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1a73-133">For detailed information about Excel JavaScript API requirement sets, see the [Excel JavaScript API requirement sets](../requirement-sets/excel-api-requirement-sets.md) article.</span></span>

## <a name="excel-javascript-api-reference"></a><span data-ttu-id="c1a73-134">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="c1a73-134">Excel JavaScript API reference</span></span>

<span data-ttu-id="c1a73-135">Excel JavaScript API の詳細については、[Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="c1a73-135">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="see-also"></a><span data-ttu-id="c1a73-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="c1a73-136">See also</span></span>

- [<span data-ttu-id="c1a73-137">Excel アドインの概要</span><span class="sxs-lookup"><span data-stu-id="c1a73-137">Excel add-ins overview</span></span>](/office/dev/add-ins/excel/excel-add-ins-overview)
- [<span data-ttu-id="c1a73-138">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="c1a73-138">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
- [<span data-ttu-id="c1a73-139">GitHub の Excel アドインのサンプル</span><span class="sxs-lookup"><span data-stu-id="c1a73-139">Excel add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
- [<span data-ttu-id="c1a73-140">API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="c1a73-140">API open specifications</span></span>](../openspec/openspec.md)
