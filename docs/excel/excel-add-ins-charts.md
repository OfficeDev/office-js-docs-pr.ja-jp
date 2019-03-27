---
title: Excel JavaScript API を使用してグラフを操作する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: adaae105050a01968525e38deb6db570b962ddde
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870885"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してグラフを操作する

この記事では、Excel JavaScript API を使用して、グラフの一般的なタスクを実行する方法のサンプル コードを提供します。 **Chart** オブジェクトと **ChartCollection** オブジェクトをサポートするプロパティとメソッドの完全なリストについては、「[Chart Object オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.chart)」および「[Chart Collection オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.chartcollection)」を参照してください。

## <a name="create-a-chart"></a>グラフの作成

次のコード サンプルでは、**Sample** というワークシートにグラフを作成します。 グラフは、範囲 **A1:B13** のデータに基づいた**折れ線**グラフです。

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

**新しい折れ線グラフ**

![Excel での新しい折れ線グラフ](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>データ系列をグラフに追加する

次のコード サンプルは、ワークシートの最初のグラフにデータ系列を追加します。 新しいデータ系列は **2016** という名前の列に対応し、範囲 **D2:D5** のデータに基づいています。

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

**2016 データ系列が追加される前のグラフ**

![2016 データ系列が追加される前の Excel のグラフ](../images/excel-charts-data-series-before.png)

**2016 データ系列が追加された後のグラフ**

![2016 データ系列が追加された後の Excel のグラフ](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>グラフ タイトルを設定する

次のコード サンプルは、ワークシートの最初のグラフのタイトルを **Sales Data by Year** に設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**タイトル設定後のグラフ**

![タイトルが付いた Excel のグラフ](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>グラフの軸のプロパティを設定する

縦棒グラフ、横棒グラフ、散布図などの[デカルト座標系](https://en.wikipedia.org/wiki/Cartesian_coordinate_system)を使用するグラフには、項目軸と数値軸が含まれています。 次の例で、タイトルを設定し、グラフの軸の単位を表示する方法を示します。

### <a name="set-axis-title"></a>軸のタイトルを設定する

次のコード サンプルは、ワークシートの最初のグラフの、項目軸のタイトルを **Product** に設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**項目軸のタイトルが設定された後のグラフ**

![軸のタイトルが付いた Excel のグラフ](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>軸の表示単位を設定する

次のコード サンプルは、ワークシートの最初のグラフの、数値軸の表示単位を **Hundreds** に設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**数値軸の表示単位が設定された後のグラフ**

![軸の表示単位が付いた Excel のグラフ](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>グラフの枠線の表示/非表示を設定する

次のコード サンプルは、ワークシートの最初のグラフの、数値軸の主な枠線を非表示にします。 `chart.axes.valueAxis.majorGridlines.visible` を **true** に設定すると、グラフの数値軸の主な枠線を表示できます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**枠線が非表示にされたグラフ**

![枠線が非表示にされた Excel のグラフ](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>グラフの近似曲線

### <a name="add-a-trendline"></a>近似曲線を追加する

次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に移動平均の近似曲線を追加します。近似曲線は 5 期間にわたる移動平均を示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**移動平均の近似曲線が記入されたグラフ**

![移動平均の近似曲線が記入された Excel のグラフ](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>近似曲線を更新する

次のコード サンプルは、**Sample** という名前のワークシートの、最初のグラフの最初の系列に対して、近似曲線の種類を**線形**に設定しています。

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

**線形の近似曲線が記入されたグラフ**

![線形の近似曲線が記入された Excel のグラフ](../images/excel-charts-trendline-linear.png)

## <a name="export-a-chart-as-an-image"></a>グラフを画像としてエクスポートする

グラフを Excel の外部で画像としてレンダリングできます。 `Chart.getImage` からは、グラフを JPEG 画像として表す base 64 エンコード文字列が返されます。 次のコードでは、画像の文字列を取得してコンソールに表示する方法を示します。

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

`Chart.getImage` は、省略可能なパラメーターとして幅、高さ、自動調整モードの 3 つを受け取ります。

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

これらのパラメーターにより、画像のサイズが決まります。 画像は常に同じ縦横比でスケーリングされます。 幅と高さのパラメーターにより、スケーリングされた画像の上端または下端が設定されます。 `ImageFittingMode` には 3 つの値があり、次のように動作します。

- `Fill`: 画像の最小の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。 これは、自動調整モードが指定されていない場合の既定の動作です。
- `Fit`: 画像の最大の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。
- `FitAndCenter`: 画像の最大の高さまたは幅が、指定された高さまたは幅になります (画像をスケーリングしたときに最初に達した方)。 結果の画像は、他の寸法について中央に配置されます。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
