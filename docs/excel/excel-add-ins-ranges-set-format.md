---
title: JavaScript API を使用して範囲のExcel設定する
description: JavaScript API の Excelを使用して範囲の形式を設定する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a09d3b4d79584e186c0be37d4a30954c4d4d0086
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075727"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="b82d5-103">JavaScript API を使用して範囲Excel設定する</span><span class="sxs-lookup"><span data-stu-id="b82d5-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="b82d5-104">この記事では、JavaScript API を使用して範囲のセルのフォントの色、塗りつぶしの色、および数値Excelします。</span><span class="sxs-lookup"><span data-stu-id="b82d5-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="b82d5-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="b82d5-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="b82d5-106">フォントの色と塗りつぶしの色を設定する</span><span class="sxs-lookup"><span data-stu-id="b82d5-106">Set font color and fill color</span></span>

<span data-ttu-id="b82d5-107">次のコード サンプルは、範囲 **B2：E2** のセルのフォントの色と塗りつぶしの色を設定します。</span><span class="sxs-lookup"><span data-stu-id="b82d5-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="b82d5-108">フォントの色と塗りつぶしの色を設定する前の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="b82d5-108">Data in range before font color and fill color are set</span></span>

![書式が設定Excel前のデータ。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="b82d5-110">フォントの色と塗りつぶしの色を設定した後の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="b82d5-110">Data in range after font color and fill color are set</span></span>

![書式が設定Excel後のデータ。](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="b82d5-112">数値の書式を設定する</span><span class="sxs-lookup"><span data-stu-id="b82d5-112">Set number format</span></span>

<span data-ttu-id="b82d5-113">次のコード サンプルは、範囲 **D3：E5** のセルの数値を書式を設定します。</span><span class="sxs-lookup"><span data-stu-id="b82d5-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="b82d5-114">数値の書式を設定する前の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="b82d5-114">Data in range before number format is set</span></span>

![数値の形式Excel前のデータ。](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="b82d5-116">数値の書式を設定した後の範囲内のデータ</span><span class="sxs-lookup"><span data-stu-id="b82d5-116">Data in range after number format is set</span></span>

![数値の形式Excel後のデータ。](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="b82d5-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="b82d5-118">See also</span></span>

- [<span data-ttu-id="b82d5-119">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="b82d5-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b82d5-120">JavaScript API を使用してセルExcelする</span><span class="sxs-lookup"><span data-stu-id="b82d5-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="b82d5-121">JavaScript API を使用して範囲を設定Excel取得する</span><span class="sxs-lookup"><span data-stu-id="b82d5-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="b82d5-122">JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する</span><span class="sxs-lookup"><span data-stu-id="b82d5-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
