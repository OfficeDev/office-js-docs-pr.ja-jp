---
title: JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する
description: JavaScript API の Excelを使用して、範囲の値、テキスト、または数式を設定および取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5d9d1bf3b248585bf27ac591754cfa4eb4dd0fbc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075748"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する

この記事では、JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得するExcelします。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a>値または数式を設定する

次のコード サンプルでは、1 つのセルまたはセル範囲の値と数式を設定します。

### <a name="set-value-for-a-single-cell"></a>1 つのセルの値を設定する

次のコード サンプルでは、セル **C3** の値を "5" に設定し、データに最も適した列の幅を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a>セルの値が更新される前のデータ

![セル値Excel前のデータ。](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a>セルの値が更新された後のデータ

![セルのExcel後のデータ。](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>複数のセルの範囲の値を設定する

次のコード サンプルでは、範囲 **B5：D5** のセルの値を設定し、データに最も適した列の幅を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];

    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-values-are-updated"></a>複数のセルの値が更新される前のデータ

![セル値Excel前のデータ。](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a>複数のセルの値が更新された後のデータ

![セルのExcel後のデータ。](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>1 つのセルの数式を設定する

次のコード サンプルでは、セル **E3** の数式を設定し、データに最も適した列の幅を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a>セルの数式が設定される前のデータ

![セル式Excel前のデータ。](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a>セルの数式が設定された後のデータ

![セル数式Excel後のデータ。](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>セルの範囲の数式を設定する

次のコード サンプルでは、範囲 **E2:E6** のセルの数式を設定し、データに最も適した列の幅を設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formulas-are-set"></a>複数のセルの数式が設定される前のデータ

![セル数式がExcel前のデータ。](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a>複数のセルの数式が設定された後のデータ

![セル数式Excel後のデータ。](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>値、テキスト、または数式を取得する

これらのコード サンプルは、セルの範囲から値、テキスト、および数式を取得します。

### <a name="get-values-from-a-range-of-cells"></a>セルの範囲から値を取得する

次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、コンソール `values` に値を書き込みます。 範囲 `values` のプロパティは、セルに含まれる生の値を指定します。 範囲内の一部のセルに数式が含まれている場合でも、範囲のプロパティは、これらのセルの生の値を指定し、数式 `values` は指定しない。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>範囲内のデータ (列 E の値は数式の結果)

![セル数式Excel後のデータ。](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a>range.values (上記のコード サンプルによりコンソールに記録される)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a>セルの範囲からテキストを取得する

次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、 `text` コンソールに書き込みます。 範囲 `text` のプロパティは、範囲内のセルの表示値を指定します。 範囲内の一部のセルに数式が含まれている場合でも、範囲のプロパティは、これらのセルの表示値を指定します。数式 `text` は指定されません。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>範囲内のデータ (列 E の値は数式の結果)

![セル数式Excel後のデータ。](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a>range.text (上記のコード サンプルによりコンソールに記録される)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a>セルの範囲から数式を取得する

次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、 `formulas` コンソールに書き込みます。 範囲のプロパティは、数式を含むセルの数式と、数式を含むセルの生の値 `formulas` を指定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>範囲内のデータ (列 E の値は数式の結果)

![セル数式Excel後のデータ。](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a>range.formulas (上記のコード サンプルによりコンソールに記録される)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を設定Excel取得する](excel-add-ins-ranges-set-get.md)
- [JavaScript API を使用して範囲Excel設定する](excel-add-ins-ranges-set-format.md)
