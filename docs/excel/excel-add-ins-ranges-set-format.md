---
title: Excel JavaScript API を使用して範囲の形式を設定する
description: Excel JavaScript API を使用して範囲の形式を設定する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652910"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して範囲の形式を設定する

この記事では、Excel JavaScript API を使用して範囲内のセルのフォントの色、塗りつぶしの色、および数値の形式を設定するコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。

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

![書式設定する前の Excel のデータ](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>フォントの色と塗りつぶしの色を設定した後の範囲内のデータ

![書式設定した後の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

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

![数値形式が設定される前の Excel のデータ](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>数値の書式を設定した後の範囲内のデータ

![数値形式が設定された後の Excel のデータ](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel JavaScript API を使用して範囲を設定および取得する](excel-add-ins-ranges-set-get.md)
- [Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得する](excel-add-ins-ranges-set-get-values.md)
