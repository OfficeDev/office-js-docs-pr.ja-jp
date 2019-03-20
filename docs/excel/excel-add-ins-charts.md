---
title: Excel JavaScript API を使用してグラフを操作する
description: ''
ms.date: 03/11/2019
localization_priority: Priority
ms.openlocfilehash: f058110c7c150a75c847a07df83aa2795c891025
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600264"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="f70dd-102">Excel JavaScript API を使用してグラフを操作する</span><span class="sxs-lookup"><span data-stu-id="f70dd-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="f70dd-p101">この記事では、Excel JavaScript API を使用して、グラフの一般的なタスクを実行する方法のサンプル コードを提供します。 **Chart** オブジェクトと **ChartCollection** オブジェクトをサポートするプロパティとメソッドの完全なリストについては、「[Chart Object オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart)」および「[Chart Collection オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p101">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API. For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="f70dd-105">グラフの作成</span><span class="sxs-lookup"><span data-stu-id="f70dd-105">Create a chart</span></span>

<span data-ttu-id="f70dd-p102">次のコード サンプルでは、**Sample** というワークシートにグラフを作成します。 グラフは、範囲 **A1:B13** のデータに基づいた**折れ線**グラフです。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p102">The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

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

<span data-ttu-id="f70dd-108">**新しい折れ線グラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-108">**New line chart**</span></span>

![Excel での新しい折れ線グラフ](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="f70dd-110">データ系列をグラフに追加する</span><span class="sxs-lookup"><span data-stu-id="f70dd-110">Add a data series to a chart</span></span>

<span data-ttu-id="f70dd-p103">次のコード サンプルは、ワークシートの最初のグラフにデータ系列を追加します。 新しいデータ系列は **2016** という名前の列に対応し、範囲 **D2:D5** のデータに基づいています。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p103">The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

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

<span data-ttu-id="f70dd-113">**2016 データ系列が追加される前のグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-113">**Chart before the 2016 data series is added**</span></span>

![2016 データ系列が追加される前の Excel のグラフ](../images/excel-charts-data-series-before.png)

<span data-ttu-id="f70dd-115">**2016 データ系列が追加された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-115">**Chart after the 2016 data series is added**</span></span>

![2016 データ系列が追加された後の Excel のグラフ](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="f70dd-117">グラフ タイトルを設定する</span><span class="sxs-lookup"><span data-stu-id="f70dd-117">Set chart title</span></span>

<span data-ttu-id="f70dd-118">次のコード サンプルは、ワークシートの最初のグラフのタイトルを **Sales Data by Year** に設定します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-118">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-119">**タイトル設定後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-119">**Chart after title is set**</span></span>

![タイトルが付いた Excel のグラフ](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="f70dd-121">グラフの軸のプロパティを設定する</span><span class="sxs-lookup"><span data-stu-id="f70dd-121">Set properties of an axis in a chart</span></span>

<span data-ttu-id="f70dd-p104">縦棒グラフ、横棒グラフ、散布図などの[デカルト座標系](https://en.wikipedia.org/wiki/Cartesian_coordinate_system)を使用するグラフには、項目軸と数値軸が含まれています。 次の例で、タイトルを設定し、グラフの軸の単位を表示する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p104">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis. These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="f70dd-124">軸のタイトルを設定する</span><span class="sxs-lookup"><span data-stu-id="f70dd-124">Set axis title</span></span>

<span data-ttu-id="f70dd-125">次のコード サンプルは、ワークシートの最初のグラフの、項目軸のタイトルを **Product** に設定します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-125">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-126">**項目軸のタイトルが設定された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-126">**Chart after title of category axis is set**</span></span>

![軸のタイトルが付いた Excel のグラフ](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="f70dd-128">軸の表示単位を設定する</span><span class="sxs-lookup"><span data-stu-id="f70dd-128">Set axis display unit</span></span>

<span data-ttu-id="f70dd-129">次のコード サンプルは、ワークシートの最初のグラフの、数値軸の表示単位を **Hundreds** に設定します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-129">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-130">**数値軸の表示単位が設定された後のグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-130">**Chart after display unit of value axis is set**</span></span>

![軸の表示単位が付いた Excel のグラフ](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="f70dd-132">グラフの枠線の表示/非表示を設定する</span><span class="sxs-lookup"><span data-stu-id="f70dd-132">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="f70dd-p105">次のコード サンプルは、ワークシートの最初のグラフの、数値軸の主な枠線を非表示にします。 `chart.axes.valueAxis.majorGridlines.visible` を **true** に設定すると、グラフの数値軸の主な枠線を表示できます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p105">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet. You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-135">**枠線が非表示にされたグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-135">**Chart with gridlines hidden**</span></span>

![枠線が非表示にされた Excel のグラフ](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="f70dd-137">グラフの近似曲線</span><span class="sxs-lookup"><span data-stu-id="f70dd-137">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="f70dd-138">近似曲線を追加する</span><span class="sxs-lookup"><span data-stu-id="f70dd-138">Add a trendline</span></span>

<span data-ttu-id="f70dd-p106">次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に移動平均の近似曲線を追加します。近似曲線は 5 期間にわたる移動平均を示します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-141">**移動平均の近似曲線が記入されたグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-141">**Chart with moving average trendline**</span></span>

![移動平均の近似曲線が記入された Excel のグラフ](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="f70dd-143">近似曲線を更新する</span><span class="sxs-lookup"><span data-stu-id="f70dd-143">Update a trendline</span></span>

<span data-ttu-id="f70dd-144">次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に対して、近似曲線の種類を**線形**に設定しています。</span><span class="sxs-lookup"><span data-stu-id="f70dd-144">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

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

<span data-ttu-id="f70dd-145">**線形の近似曲線が記入されたグラフ**</span><span class="sxs-lookup"><span data-stu-id="f70dd-145">**Chart with linear trendline**</span></span>

![線形の近似曲線が記入された Excel のグラフ](../images/excel-charts-trendline-linear.png)

## <a name="export-a-chart-as-an-image"></a><span data-ttu-id="f70dd-147">グラフを画像としてエクスポートする</span><span class="sxs-lookup"><span data-stu-id="f70dd-147">Export a chart as an image</span></span>

<span data-ttu-id="f70dd-148">グラフを Excel の外部で画像としてレンダリングできます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-148">Charts can be rendered as images outside of Excel.</span></span> <span data-ttu-id="f70dd-149">`Chart.getImage` からは、グラフを JPEG 画像として表す base 64 エンコード文字列が返されます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-149">`Chart.getImage` returns the chart as a base64-encoded string representing the chart as a JPEG image.</span></span> <span data-ttu-id="f70dd-150">次のコードでは、画像の文字列を取得してコンソールに表示する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-150">The following code shows how to get the image string and log it to the console.</span></span>

```js
Excel.run(function (ctx) {
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    var imageAsString = chart.getImage();
    return context.sync().then(function () {
        console.log(imageAsString.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f70dd-151">`Chart.getImage` は、省略可能なパラメーターとして幅、高さ、自動調整モードの 3 つを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="f70dd-151">`Chart.getImage` takes three optional parameters: width, height, and the fitting mode.</span></span>

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

<span data-ttu-id="f70dd-152">これらのパラメーターにより、画像のサイズが決まります。</span><span class="sxs-lookup"><span data-stu-id="f70dd-152">These parameters determine the size of the image.</span></span> <span data-ttu-id="f70dd-153">画像は常に同じ縦横比でスケーリングされます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-153">Images are always proportionally scaled.</span></span> <span data-ttu-id="f70dd-154">幅と高さのパラメーターにより、スケーリングされた画像の上端または下端が設定されます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-154">The width and height parameters put upper or lower bounds on the scaled image.</span></span> <span data-ttu-id="f70dd-155">`ImageFittingMode` には 3 つの値があり、次のように動作します。</span><span class="sxs-lookup"><span data-stu-id="f70dd-155">`ImageFittingMode` has three values with the following behaviors:</span></span>

- <span data-ttu-id="f70dd-156">`Fill`: 画像の最小の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。</span><span class="sxs-lookup"><span data-stu-id="f70dd-156">`Fill`: The image’s minimum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span> <span data-ttu-id="f70dd-157">これは、自動調整モードが指定されていない場合の既定の動作です。</span><span class="sxs-lookup"><span data-stu-id="f70dd-157">This is the default behavior if no character is specified.</span></span>
- <span data-ttu-id="f70dd-158">`Fit`: 画像の最大の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。</span><span class="sxs-lookup"><span data-stu-id="f70dd-158">`Fit`: The image’s maximum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span>
- <span data-ttu-id="f70dd-159">`FitAndCenter`: 画像の最大の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。</span><span class="sxs-lookup"><span data-stu-id="f70dd-159">`FitAndCenter`: The image’s maximum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span> <span data-ttu-id="f70dd-160">結果の画像は、他の寸法について中央に配置されます。</span><span class="sxs-lookup"><span data-stu-id="f70dd-160">The resulting image is centered relative to the other dimension.</span></span>

## <a name="see-also"></a><span data-ttu-id="f70dd-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="f70dd-161">See also</span></span>

- [<span data-ttu-id="f70dd-162">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="f70dd-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
