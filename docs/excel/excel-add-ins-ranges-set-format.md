---
title: JavaScript API を使用して範囲のExcel設定する
description: JavaScript API の Excelを使用して範囲の形式を設定する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 3e9c8dd58216de695fc0bc634a6eaed7dbc724468bd9e02bc23bf34c394b5e16
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086497"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲Excel設定する

この記事では、JavaScript API を使用して範囲のセルのフォントの色、塗りつぶしの色、および数値Excelします。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>フォントの色と塗りつぶしの色を設定する

次のコード サンプルは、範囲 **B2：E2** のセルのフォントの色と塗りつぶしの色を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>フォントの色と塗りつぶしの色を設定する前の範囲内のデータ

![書式が設定Excel前のデータ。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>フォントの色と塗りつぶしの色を設定した後の範囲内のデータ

![書式が設定Excel後のデータ。](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>数値の書式を設定する

次のコード サンプルは、範囲 **D3：E5** のセルの数値を書式を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a>数値の書式を設定する前の範囲内のデータ

![数値の形式Excel前のデータ。](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>数値の書式を設定した後の範囲内のデータ

![数値の形式Excel後のデータ。](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を設定Excel取得する](excel-add-ins-ranges-set-get.md)
- [JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する](excel-add-ins-ranges-set-get-values.md)
