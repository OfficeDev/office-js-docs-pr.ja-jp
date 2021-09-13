---
title: Excel JavaScript API を使用してグラフを操作する
description: JavaScript API を使用してグラフ タスクを示すExcelサンプルです。
ms.date: 07/17/2019
ms.localizationpriority: medium
ms.openlocfilehash: b3cb04ff3bd8b1b0c050741a7238b1e9d6bd498f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151440"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してグラフを操作する

この記事では、Excel JavaScript API を使用して、グラフの一般的なタスクを実行する方法のサンプル コードを提供します。
and オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「Chart `Chart` `ChartCollection` Object [(JavaScript API for Excel)」](/javascript/api/excel/excel.chart)および「Chart [Collection Object (JavaScript API for](/javascript/api/excel/excel.chartcollection)Excel)」を参照してください。

## <a name="create-a-chart"></a>グラフの作成

次のコード サンプルでは、**Sample** というワークシートにグラフを作成します。 グラフは、範囲 **A1:B13** のデータに基づいた **折れ線** グラフです。

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

![グラフの新しいExcel。](../images/excel-charts-create-line.png)


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

![2016 Excel前のグラフを追加しました。](../images/excel-charts-data-series-before.png)

**2016 データ系列が追加された後のグラフ**

![2016 Excelデータ系列が追加された後のグラフ。](../images/excel-charts-data-series-after.png)

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

![グラフにタイトルが含Excel。](../images/excel-charts-title-set.png)

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

![グラフに軸のタイトルが含Excel。](../images/excel-charts-axis-title-set.png)

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

![グラフに軸表示単位をExcel。](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>グラフの枠線の表示/非表示を設定する

次のコード サンプルは、ワークシートの最初のグラフの、数値軸の主な枠線を非表示にします。 に設定すると、グラフの値軸の大きなグリッド線を表示 `chart.axes.valueAxis.majorGridlines.visible` できます `true` 。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**枠線が非表示にされたグラフ**

![グリッド線が非表示のグラフは、Excel。](../images/excel-charts-gridlines-removed.png)

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

![グラフ内の移動平均の傾向線Excel。](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>近似曲線を更新する

次のコード サンプルでは、Sample という名前のワークシートの最初のグラフの最初の系列に対して、傾向線 `Linear` を入力します。 

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

![グラフに線形の傾向線Excel。](../images/excel-charts-trendline-linear.png)

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

これらのパラメーターにより、画像のサイズが決まります。 画像は常に同じ縦横比でスケーリングされます。 幅と高さのパラメーターにより、スケーリングされた画像の上端または下端が設定されます。 `ImageFittingMode` 次の動作を持つ 3 つの値があります。

- `Fill`: イメージの最小の高さまたは幅は、指定された高さまたは幅です (イメージのスケーリング時に最初に到達した方)。 これは、自動調整モードが指定されていない場合の既定の動作です。
- `Fit`: イメージの最大の高さまたは幅は、指定された高さまたは幅です (イメージのスケーリング時に最初に到達した方)。
- `FitAndCenter`: イメージの最大の高さまたは幅は、指定された高さまたは幅です (イメージのスケーリング時に最初に到達した方)。 結果の画像は、他の寸法について中央に配置されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
