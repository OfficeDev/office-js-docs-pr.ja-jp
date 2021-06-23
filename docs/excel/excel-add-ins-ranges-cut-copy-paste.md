---
title: JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける
description: JavaScript API を使用して範囲を切り取り、コピー、貼り付けるExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2112702110b72e0020ed72090ce495abb3ff5366
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075825"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="a7fae-103">JavaScript API を使用して範囲を切り取り、コピー Excel貼り付ける</span><span class="sxs-lookup"><span data-stu-id="a7fae-103">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="a7fae-104">この記事では、JavaScript API を使用して範囲を切り取り、コピー、貼り付けるExcel説明します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-104">This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="a7fae-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="a7fae-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a><span data-ttu-id="a7fae-106">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="a7fae-106">Copy and paste</span></span>

<span data-ttu-id="a7fae-107">[Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)メソッドは、ユーザー UI の **コピー** と **貼** り付Excelします。</span><span class="sxs-lookup"><span data-stu-id="a7fae-107">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="a7fae-108">宛先は、 `Range` 呼び出 `copyFrom` されるオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="a7fae-108">The destination is the `Range` object that `copyFrom` is called on.</span></span> <span data-ttu-id="a7fae-109">コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="a7fae-109">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="a7fae-110">次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。</span><span class="sxs-lookup"><span data-stu-id="a7fae-110">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a7fae-111">`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。</span><span class="sxs-lookup"><span data-stu-id="a7fae-111">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="a7fae-112">`copyType` では、ソースからコピー先にコピーされるデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-112">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="a7fae-113">`Excel.RangeCopyType.formulas` ソース セル内の数式を転送し、それらの数式の範囲の相対位置を保持します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-113">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="a7fae-114">任意の数式以外のエントリはそのままコピーされます。</span><span class="sxs-lookup"><span data-stu-id="a7fae-114">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="a7fae-115">`Excel.RangeCopyType.values` では、データ値と、数式の場合は数式の結果をコピーします。</span><span class="sxs-lookup"><span data-stu-id="a7fae-115">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="a7fae-116">`Excel.RangeCopyType.formats` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。</span><span class="sxs-lookup"><span data-stu-id="a7fae-116">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="a7fae-117">`Excel.RangeCopyType.all` (既定のオプション) は、データと書式設定の両方をコピーし、セルの数式が見つかった場合は保持します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-117">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="a7fae-118">`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-118">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="a7fae-119">true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。</span><span class="sxs-lookup"><span data-stu-id="a7fae-119">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="a7fae-120">スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。</span><span class="sxs-lookup"><span data-stu-id="a7fae-120">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="a7fae-121">既定値は false です。</span><span class="sxs-lookup"><span data-stu-id="a7fae-121">The default is false.</span></span>

<span data-ttu-id="a7fae-122">`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-122">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="a7fae-123">行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。</span><span class="sxs-lookup"><span data-stu-id="a7fae-123">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="a7fae-124">次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。</span><span class="sxs-lookup"><span data-stu-id="a7fae-124">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-copied-and-pasted"></a><span data-ttu-id="a7fae-125">範囲がコピーおよび貼り付けされる前のデータ</span><span class="sxs-lookup"><span data-stu-id="a7fae-125">Data before range is copied and pasted</span></span>

![範囲のコピー Excel実行する前のデータ。](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a><span data-ttu-id="a7fae-127">範囲がコピーおよび貼り付けされた後のデータ</span><span class="sxs-lookup"><span data-stu-id="a7fae-127">Data after range is copied and pasted</span></span>

![範囲のコピー Excelが実行された後のデータ。](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a><span data-ttu-id="a7fae-129">セルの切り取りと貼り付け (移動)</span><span class="sxs-lookup"><span data-stu-id="a7fae-129">Cut and paste (move) cells</span></span>

<span data-ttu-id="a7fae-130">[Range.moveTo メソッドは](/javascript/api/excel/excel.range#moveto-destinationrange-)、ブック内の新しい場所にセルを移動します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-130">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="a7fae-131">このセルの移動動作は、セルを移動するときに、範囲 [](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e)の境界線をドラッグするか、切り取りおよび貼り付けアクションを実行する場合 **と\*\*\*\*同じように動作** します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-131">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="a7fae-132">範囲の書式設定と値の両方が、パラメーターとして指定された場所に移動 `destinationRange` されます。</span><span class="sxs-lookup"><span data-stu-id="a7fae-132">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="a7fae-133">次のコード サンプルでは、メソッドを使用して範囲を移動 `Range.moveTo` します。</span><span class="sxs-lookup"><span data-stu-id="a7fae-133">The following code sample moves a range with the `Range.moveTo` method.</span></span> <span data-ttu-id="a7fae-134">移動先の範囲がソースより小さい場合は、ソース コンテンツを含む範囲に拡張されます。</span><span class="sxs-lookup"><span data-stu-id="a7fae-134">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="a7fae-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="a7fae-135">See also</span></span>

- [<span data-ttu-id="a7fae-136">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="a7fae-136">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a7fae-137">JavaScript API を使用してセルExcelする</span><span class="sxs-lookup"><span data-stu-id="a7fae-137">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="a7fae-138">JavaScript API を使用して重複Excel削除する</span><span class="sxs-lookup"><span data-stu-id="a7fae-138">Remove duplicates using the Excel JavaScript API</span></span>](excel-add-ins-ranges-remove-duplicates.md)
- [<span data-ttu-id="a7fae-139">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="a7fae-139">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
