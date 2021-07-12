---
title: Office アドインの Excel JavaScript オブジェクト モデル
description: Excel JavaScript API の主要なオブジェクトの種類と、それらを使用して Excel のアドインを構築する方法を説明します。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6c88dc84796d9fd898bee880035ed964ab6cd7c8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349561"
---
# <a name="excel-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="fd7c5-103">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="fd7c5-103">Excel JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="fd7c5-104">この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用して Excel 2016 以降のアドインをビルドする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="fd7c5-105">ここでは API の使用の基本となる中心概念について説明し、広い範囲に対する読み取り、書き込み、一定範囲内すべてのセルの更新など、特定のタスクを実行するためのガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fd7c5-106">Excel API の非同期性とブックでの動作方法については、「[Using the application-specific API model (アプリケーション固有の API モデルの使用)](../develop/application-specific-api-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="fd7c5-107">Excel 用の Office.js API</span><span class="sxs-lookup"><span data-stu-id="fd7c5-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="fd7c5-108">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="fd7c5-109">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="fd7c5-110">**共通 API**: Office 2013 で導入された [共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="fd7c5-p102">Excel 2016 以降を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-p102">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API. For example:</span></span>

* <span data-ttu-id="fd7c5-p103">[Context](/javascript/api/office/office.context): `Context`Context`contentLanguage` オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。 これは `officeTheme` や `host` などのブック構成の詳細で構成され、`platform` や `requirements.isSetSupported()` などのアドインのランタイム環境に関する情報も提供します。 さらに、 メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-p103">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="fd7c5-116">[Document](/javascript/api/office/office.document): `Document` オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="fd7c5-117">次の図は、Excel JavaScript API または共通 API を使用するタイミングを示しています。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Excel JS API と共通 API の違い。](../images/excel-js-api-common-api.png)

## <a name="excel-specific-object-model"></a><span data-ttu-id="fd7c5-119">Excel 固有のオブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="fd7c5-119">Excel-specific object model</span></span>

<span data-ttu-id="fd7c5-120">Excel API について理解するには、ブックの構成要素が互いにどのように関連しているかを理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="fd7c5-121">**ブック** には、1 つ以上の **ワークシート** が含まれます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="fd7c5-122">**ワークシート** には、個々のシートに存在するデータ オブジェクトのコレクションが含まれており、**Range** オブジェクトを介してセルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet, and gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="fd7c5-123">**Range** は、連続したセルのグループを表します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="fd7c5-124">**Range** は、**表**、**グラフ**、**図形**、およびその他のデータ可視化や組織オブジェクトを作成して配置するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="fd7c5-125">**ブック** には、**ブック** 全体のデータ オブジェクト (**表** など) の一部のコレクションが含まれます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-125">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

### <a name="ranges"></a><span data-ttu-id="fd7c5-126">範囲</span><span class="sxs-lookup"><span data-stu-id="fd7c5-126">Ranges</span></span>

<span data-ttu-id="fd7c5-127">範囲とは、ブック内の連続したセルのグループのことです。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-127">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="fd7c5-128">アドインでは、範囲を定義するのに通常 A1 形式の表記が使用されます (例: **B3** は、列 **B**、行 **3** の単一のセルで、**C2:F4** は、列 **C** から **F**、行 **2** から **4** までのセル)。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-128">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="fd7c5-129">範囲には `values`、`formulas`、`format` の 3 つの主要なプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-129">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="fd7c5-130">これらのプロパティで、セルの値、評価する数式、およびセルの視覚的な書式設定を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-130">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="fd7c5-131">サンプル範囲</span><span class="sxs-lookup"><span data-stu-id="fd7c5-131">Range sample</span></span>

<span data-ttu-id="fd7c5-132">次のサンプルで、売上記録の作成方法を示します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-132">The following sample shows how to create sales records.</span></span> <span data-ttu-id="fd7c5-133">この関数は、`Range` オブジェクトを使用して、値、数式、書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-133">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

<span data-ttu-id="fd7c5-134">このサンプルは、現在のワークシートに次のデータを作成します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-134">This sample creates the following data in the current worksheet.</span></span>

![値の行、数式の列、書式設定されたヘッダーを示す売上記録。](../images/excel-overview-range-sample.png)

<span data-ttu-id="fd7c5-136">詳細については、「[Excel JavaScript API を使用した範囲値、テキスト、または数式の設定と取得](excel-add-ins-ranges-set-get-values.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-136">For more information, see [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md).</span></span>

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="fd7c5-137">グラフ、表、およびその他のデータ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fd7c5-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="fd7c5-138">Excel JavaScript API を使用することにより、Excel 内でデータ構造やビジュアル化を作成および操作できます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="fd7c5-139">表とグラフの 2 つのオブジェクトが頻繁に使用されますが、API はピボットテーブル、図形、画像などもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="fd7c5-140">表の作成</span><span class="sxs-lookup"><span data-stu-id="fd7c5-140">Creating a table</span></span>

<span data-ttu-id="fd7c5-p108">データ入力範囲を使ってテーブルを作成します。書式設定とテーブル コントロール (フィルターなど) が自動的に範囲に適用されます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-p108">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="fd7c5-143">次のサンプルでは、前のサンプルの範囲を使用して表を作成します。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="fd7c5-144">前のデータを含むワークシート上でこのサンプル コードを使用すると、次のテーブルが作成されます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-144">Using this sample code on the worksheet with the previous data creates the following table.</span></span>

![前の売上記録から作成された表。](../images/excel-overview-table-sample.png)

<span data-ttu-id="fd7c5-146">詳細については、「[Excel JavaScript API を使用して表を操作する](excel-add-ins-tables.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-146">For more information, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>

#### <a name="creating-a-chart"></a><span data-ttu-id="fd7c5-147">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="fd7c5-147">Creating a chart</span></span>

<span data-ttu-id="fd7c5-148">グラフを作成すると、範囲内のデータを視覚化できます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-148">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="fd7c5-149">この API は、さまざまな種類のグラフをサポートしています。いずれのグラフも、必要に応じてカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-149">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="fd7c5-150">次のサンプルでは 3 つの品目の簡単な縦棒グラフが作成され、ワークシートの上端から 100 ピクセル下に配置されます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-150">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="fd7c5-151">前の表を含むワークシート上でこのサンプルを実行すると、次のグラフが作成されます。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-151">Running this sample on the worksheet with the previous table creates the following chart.</span></span>

![前の売上記録の 3 つの品目の数量が表示されている縦棒グラフ。](../images/excel-overview-chart-sample.png)

<span data-ttu-id="fd7c5-153">詳細については、「[Excel JavaScript API を使用してグラフを操作する](excel-add-ins-charts.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fd7c5-153">For more information, see [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fd7c5-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="fd7c5-154">See also</span></span>

* [<span data-ttu-id="fd7c5-155">最初の Excel アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="fd7c5-155">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="fd7c5-156">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="fd7c5-156">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="fd7c5-157">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="fd7c5-157">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="fd7c5-158">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="fd7c5-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
