---
title: Excel JavaScript API を使用してグラフを操作する
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b804e2130e30626a9caf21bca1f3955c57a3f94c
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457552"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="3ea22-102">Excel JavaScript API を使用してグラフを操作する</span><span class="sxs-lookup"><span data-stu-id="3ea22-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="3ea22-103">この記事では、Excel JavaScript API を使用して、グラフの一般的なタスクを実行する方法のサンプル コードを提供します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-103">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.</span></span> <span data-ttu-id="3ea22-104">**Chart** オブジェクトと **ChartCollection** オブジェクトをサポートするプロパティとメソッドの完全なリストについては、「[Chart Object オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart)」および「[Chart Collection オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3ea22-104">For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="3ea22-105">グラフを作成する</span><span class="sxs-lookup"><span data-stu-id="3ea22-105">Create a chart</span></span>

<span data-ttu-id="3ea22-106">次のコード サンプルでは、**Sample** というワークシートにグラフを作成します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-106">The following code sample creates a chart in the worksheet named **Sample**.</span></span> <span data-ttu-id="3ea22-107">グラフは、範囲 **A1:B13** のデータに基づいた**折れ線**グラフです。</span><span class="sxs-lookup"><span data-stu-id="3ea22-107">The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-108">**新しい折れ線グラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-108">**New line chart**</span></span>

![Excel での新しい折れ線グラフ](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="3ea22-110">データ系列をグラフに追加する</span><span class="sxs-lookup"><span data-stu-id="3ea22-110">Add a data series to a chart</span></span>

<span data-ttu-id="3ea22-111">次のコード サンプルは、ワークシートの最初のグラフにデータ系列を追加します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-111">The following code sample adds a data series to the first chart in the worksheet.</span></span> <span data-ttu-id="3ea22-112">新しいデータ系列は **2016** という名前の列に対応し、範囲 **D2:D5** のデータに基づいています。</span><span class="sxs-lookup"><span data-stu-id="3ea22-112">The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-113">**2016 データ系列が追加される前のグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-113">**Chart before the 2016 data series is added**</span></span>

![2016 データ系列が追加される前の Excel のグラフ](../images/excel-charts-data-series-before.png)

<span data-ttu-id="3ea22-115">**2016 データ系列が追加された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-115">**Chart after the 2016 data series is added**</span></span>

![2016 データ系列が追加された後の Excel のグラフ](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="3ea22-117">グラフ タイトルを設定する</span><span class="sxs-lookup"><span data-stu-id="3ea22-117">Set chart title</span></span>

<span data-ttu-id="3ea22-118">次のコード サンプルは、ワークシートの最初のグラフのタイトルを **Sales Data by Year** に設定します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-118">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-119">**タイトル設定後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-119">**Chart after title is set**</span></span>

![タイトルが付いた Excel のグラフ](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="3ea22-121">グラフの軸のプロパティを設定する</span><span class="sxs-lookup"><span data-stu-id="3ea22-121">Set properties of an axis in a chart</span></span>

<span data-ttu-id="3ea22-122">縦棒グラフ、横棒グラフ、散布図などの[デカルト座標系](https://en.wikipedia.org/wiki/Cartesian_coordinate_system)を使用するグラフには、項目軸と数値軸が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3ea22-122">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="3ea22-123">次の例で、タイトルを設定し、グラフの軸の単位を表示する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-123">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="3ea22-124">軸のタイトルを設定する</span><span class="sxs-lookup"><span data-stu-id="3ea22-124">Set axis title</span></span>

<span data-ttu-id="3ea22-125">次のコード サンプルは、ワークシートの最初のグラフの、項目軸のタイトルを **Product** に設定します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-125">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-126">**項目軸のタイトルが設定された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-126">**Chart after title of category axis is set**</span></span>

![軸のタイトルが付いた Excel のグラフ](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="3ea22-128">軸の表示単位を設定する</span><span class="sxs-lookup"><span data-stu-id="3ea22-128">Set axis display unit</span></span>

<span data-ttu-id="3ea22-129">次のコード サンプルは、ワークシートの最初のグラフの、数値軸の表示単位を **Hundreds** に設定します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-129">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-130">**数値軸の表示単位が設定された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-130">**Chart after display unit of value axis is set**</span></span>

![軸の表示単位が付いた Excel のグラフ](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="3ea22-132">グラフの枠線の表示/非表示を設定する</span><span class="sxs-lookup"><span data-stu-id="3ea22-132">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="3ea22-133">次のコード サンプルは、ワークシートの最初のグラフの、数値軸の主な枠線を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="3ea22-133">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="3ea22-134">`chart.axes.valueAxis.majorGridlines.visible` を **true** に設定すると、グラフの数値軸の主な枠線を表示できます。</span><span class="sxs-lookup"><span data-stu-id="3ea22-134">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-135">**枠線が非表示にされたグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-135">**Chart with gridlines hidden**</span></span>

![枠線が非表示にされた Excel のグラフ](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="3ea22-137">グラフの近似曲線</span><span class="sxs-lookup"><span data-stu-id="3ea22-137">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="3ea22-138">近似曲線を追加する</span><span class="sxs-lookup"><span data-stu-id="3ea22-138">Add a trendline</span></span>

<span data-ttu-id="3ea22-p106">次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に移動平均の近似曲線を追加します。近似曲線は 5 期間にわたる移動平均を示します。</span><span class="sxs-lookup"><span data-stu-id="3ea22-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-141">**移動平均の近似曲線が記入されたグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-141">**Chart with moving average trendline**</span></span>

![移動平均の近似曲線が記入された Excel のグラフ](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="3ea22-143">近似曲線を更新する</span><span class="sxs-lookup"><span data-stu-id="3ea22-143">Update a trendline</span></span>

<span data-ttu-id="3ea22-144">次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に対して、近似曲線の種類を**線形**に設定しています。</span><span class="sxs-lookup"><span data-stu-id="3ea22-144">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3ea22-145">**線形の近似曲線が記入されたグラフ**</span><span class="sxs-lookup"><span data-stu-id="3ea22-145">**Chart with linear trendline**</span></span>

![線形の近似曲線が記入された Excel のグラフ](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="3ea22-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="3ea22-147">See also</span></span>

- [<span data-ttu-id="3ea22-148">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="3ea22-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
