---
title: Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得する
description: Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad6e58c6e9fe3246d23d6ef1dd298fc6c18167a2
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652905"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a><span data-ttu-id="f3263-103">Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-103">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>

<span data-ttu-id="f3263-104">この記事では、Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="f3263-104">This article provides code samples that set and get range values, text, or formulas with the Excel JavaScript API.</span></span> <span data-ttu-id="f3263-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="f3263-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a><span data-ttu-id="f3263-106">値または数式を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-106">Set values or formulas</span></span>

<span data-ttu-id="f3263-107">次のコード サンプルでは、1 つのセルまたはセル範囲の値と数式を設定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-107">The following code samples set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="f3263-108">1 つのセルの値を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-108">Set value for a single cell</span></span>

<span data-ttu-id="f3263-109">次のコード サンプルでは、セル **C3** の値を "5" に設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-109">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a><span data-ttu-id="f3263-110">セルの値が更新される前のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-110">Data before cell value is updated</span></span>

![セルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a><span data-ttu-id="f3263-112">セルの値が更新された後のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-112">Data after cell value is updated</span></span>

![セルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="f3263-114">複数のセルの範囲の値を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-114">Set values for a range of cells</span></span>

<span data-ttu-id="f3263-115">次のコード サンプルでは、範囲 **B5：D5** のセルの値を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-115">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

#### <a name="data-before-cell-values-are-updated"></a><span data-ttu-id="f3263-116">複数のセルの値が更新される前のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-116">Data before cell values are updated</span></span>

![複数のセルの値が更新される前の Excel のデータ](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a><span data-ttu-id="f3263-118">複数のセルの値が更新された後のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-118">Data after cell values are updated</span></span>

![複数のセルの値が更新された後の Excel のデータ](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="f3263-120">1 つのセルの数式を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-120">Set formula for a single cell</span></span>

<span data-ttu-id="f3263-121">次のコード サンプルでは、セル **E3** の数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-121">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a><span data-ttu-id="f3263-122">セルの数式が設定される前のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-122">Data before cell formula is set</span></span>

![セルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a><span data-ttu-id="f3263-124">セルの数式が設定された後のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-124">Data after cell formula is set</span></span>

![セルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="f3263-126">セルの範囲の数式を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-126">Set formulas for a range of cells</span></span>

<span data-ttu-id="f3263-127">次のコード サンプルでは、範囲 **E2:E6** のセルの数式を設定し、データに最も適した列の幅を設定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-127">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

#### <a name="data-before-cell-formulas-are-set"></a><span data-ttu-id="f3263-128">複数のセルの数式が設定される前のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-128">Data before cell formulas are set</span></span>

![複数のセルの数式が設定される前の Excel のデータ](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a><span data-ttu-id="f3263-130">複数のセルの数式が設定された後のデータ</span><span class="sxs-lookup"><span data-stu-id="f3263-130">Data after cell formulas are set</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="f3263-132">値、テキスト、または数式を取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-132">Get values, text, or formulas</span></span>

<span data-ttu-id="f3263-133">これらのコード サンプルは、セルの範囲から値、テキスト、および数式を取得します。</span><span class="sxs-lookup"><span data-stu-id="f3263-133">These code samples get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="f3263-134">セルの範囲から値を取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-134">Get values from a range of cells</span></span>

<span data-ttu-id="f3263-135">次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、コンソール `values` に値を書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3263-135">The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console.</span></span> <span data-ttu-id="f3263-136">範囲 `values` のプロパティは、セルに含まれる生の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-136">The `values` property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="f3263-137">範囲内の一部のセルに数式が含まれている場合でも、範囲のプロパティは、これらのセルの生の値を指定し、数式 `values` は指定しない。</span><span class="sxs-lookup"><span data-stu-id="f3263-137">Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="f3263-138">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="f3263-138">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="f3263-140">range.values (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="f3263-140">range.values (as logged to the console by the code sample above)</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="f3263-141">セルの範囲からテキストを取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-141">Get text from a range of cells</span></span>

<span data-ttu-id="f3263-142">次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、 `text` コンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3263-142">The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console.</span></span> <span data-ttu-id="f3263-143">範囲 `text` のプロパティは、範囲内のセルの表示値を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-143">The `text` property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="f3263-144">範囲内の一部のセルに数式が含まれている場合でも、範囲のプロパティは、これらのセルの表示値を指定します。数式 `text` は指定されません。</span><span class="sxs-lookup"><span data-stu-id="f3263-144">Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="f3263-145">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="f3263-145">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="f3263-147">range.text (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="f3263-147">range.text (as logged to the console by the code sample above)</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="f3263-148">セルの範囲から数式を取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-148">Get formulas from a range of cells</span></span>

<span data-ttu-id="f3263-149">次のコード サンプルは、 **範囲 B2:E6** を取得し、プロパティを読み込み、 `formulas` コンソールに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3263-149">The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console.</span></span> <span data-ttu-id="f3263-150">範囲のプロパティは、数式を含むセルの数式と、数式を含むセルの生の値 `formulas` を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3263-150">The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="f3263-151">範囲内のデータ (列 E の値は数式の結果)</span><span class="sxs-lookup"><span data-stu-id="f3263-151">Data in range (values in column E are a result of formulas)</span></span>

![複数のセルの数式が設定された後の Excel のデータ](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="f3263-153">range.formulas (上記のコード サンプルによりコンソールに記録される)</span><span class="sxs-lookup"><span data-stu-id="f3263-153">range.formulas (as logged to the console by the code sample above)</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f3263-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="f3263-154">See also</span></span>

- [<span data-ttu-id="f3263-155">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="f3263-155">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f3263-156">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="f3263-156">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="f3263-157">Excel JavaScript API を使用して範囲を設定および取得する</span><span class="sxs-lookup"><span data-stu-id="f3263-157">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="f3263-158">Excel JavaScript API を使用して範囲の形式を設定する</span><span class="sxs-lookup"><span data-stu-id="f3263-158">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
