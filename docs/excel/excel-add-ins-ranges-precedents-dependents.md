---
title: JavaScript API を使用して数式の前例と依存Excel処理する
description: JavaScript API の Excelを使用して、数式の前例と依存を取得する方法について説明します。
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6021e383f02ca0de15210638b991dfe8b109ab63
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075797"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="c8f11-103">JavaScript API を使用して数式の前例と依存Excel取得する</span><span class="sxs-lookup"><span data-stu-id="c8f11-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="c8f11-104">Excelは、多くの場合、他のセルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="c8f11-105">これらのクロスセル参照は、"前例" および "依存" と呼ばれる。</span><span class="sxs-lookup"><span data-stu-id="c8f11-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="c8f11-106">前例は、数式にデータを提供するセルです。</span><span class="sxs-lookup"><span data-stu-id="c8f11-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="c8f11-107">従属とは、他のセルを参照する数式を含むセルです。</span><span class="sxs-lookup"><span data-stu-id="c8f11-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="c8f11-108">セル間のリレーションシップにExcelする機能の詳細については、「数式とセル間のリレーションシップを表示する[」を参照してください](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。</span><span class="sxs-lookup"><span data-stu-id="c8f11-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="c8f11-109">セルには前例のセルを含め、その前のセルには独自の前例セルを含めできます。</span><span class="sxs-lookup"><span data-stu-id="c8f11-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="c8f11-110">"直接の前例" は、親子関係の親の概念と同様に、このシーケンス内のセルの最初の前のグループです。</span><span class="sxs-lookup"><span data-stu-id="c8f11-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="c8f11-111">"直接依存" は、親子関係の子と同様に、シーケンス内のセルの最初の依存グループです。</span><span class="sxs-lookup"><span data-stu-id="c8f11-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="c8f11-112">ブック内の他のセルを参照するが、リレーションシップが親子関係ではないセルは、直接依存または直接の前例ではありません。</span><span class="sxs-lookup"><span data-stu-id="c8f11-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="c8f11-113">この記事では、JavaScript API を使用して数式の直接の前例と直接依存を取得するコード サンプルExcel示します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="c8f11-114">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="c8f11-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="c8f11-115">数式の直接の前例を取得する</span><span class="sxs-lookup"><span data-stu-id="c8f11-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="c8f11-116">[Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--)を使用して数式の直接の先行セルを検索します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="c8f11-117">`Range.getDirectPrecedents` オブジェクトを返 `WorkbookRangeAreas` します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="c8f11-118">このオブジェクトには、ブック内のすべての直接の前例のアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c8f11-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="c8f11-119">このオブジェクトには、少なくとも 1 つの数式の前例を含 `RangeAreas` むワークシートごとに個別のオブジェクトがあります。</span><span class="sxs-lookup"><span data-stu-id="c8f11-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="c8f11-120">オブジェクトの操作の詳細については、「複数の範囲を同時に操作する」を参照Excel `RangeAreas` [アドインを参照してください](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="c8f11-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="c8f11-121">次のスクリーンショットは、UI の [前例のトレース] ボタンを選択した結果Excel示しています。</span><span class="sxs-lookup"><span data-stu-id="c8f11-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="c8f11-122">このボタンは、前のセルから選択したセルに矢印を描画します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="c8f11-123">選択したセル **E3** には数式 "=C3 \* D3" が含まれているので **、C3** と **D3** の両方が先行セルです。</span><span class="sxs-lookup"><span data-stu-id="c8f11-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="c8f11-124">UI ボタンExcel異なり、 `getDirectPrecedents` メソッドは矢印を描画しない。</span><span class="sxs-lookup"><span data-stu-id="c8f11-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![UI の矢印トレースの先行セルExcelします。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="c8f11-126">メソッド `getDirectPrecedents` は、ブック間で先行セルを取得できない。</span><span class="sxs-lookup"><span data-stu-id="c8f11-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="c8f11-127">次のコード サンプルでは、アクティブな範囲の直接の前例を取得し、それらの前のセルの背景色を黄色に変更します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula-preview"></a><span data-ttu-id="c8f11-128">数式の直接依存を取得する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c8f11-128">Get the direct dependents of a formula (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="c8f11-129">この `Range.getDirectDependents` メソッドは現在、パブリック プレビューでのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="c8f11-129">The `Range.getDirectDependents` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="c8f11-130">[Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__)を使用して数式の直接依存セルを検索します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-130">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="c8f11-131">同様 `Range.getDirectPrecedents` に `Range.getDirectDependents` 、オブジェクトも返 `WorkbookRangeAreas` します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-131">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="c8f11-132">このオブジェクトには、ブック内のすべての直接依存のアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c8f11-132">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="c8f11-133">このオブジェクトには、少なくとも 1 つの数式に依存 `RangeAreas` するワークシートごとに個別のオブジェクトがあります。</span><span class="sxs-lookup"><span data-stu-id="c8f11-133">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="c8f11-134">オブジェクトの操作の詳細については、「複数の範囲を同時に操作する」を参照Excel `RangeAreas` [アドインを参照してください](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="c8f11-134">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="c8f11-135">次のスクリーンショットは、UI の [トレース依存] ボタンを選択した結果Excel示しています。</span><span class="sxs-lookup"><span data-stu-id="c8f11-135">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="c8f11-136">このボタンは、依存セルから選択したセルに矢印を描画します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-136">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="c8f11-137">選択したセル **D3** には、セル **E3** が従属セルとして含されます。</span><span class="sxs-lookup"><span data-stu-id="c8f11-137">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="c8f11-138">**E3 には** 、"=C3 \* D3" という数式が含まれる。</span><span class="sxs-lookup"><span data-stu-id="c8f11-138">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="c8f11-139">UI ボタンExcel異なり、 `getDirectDependents` メソッドは矢印を描画しない。</span><span class="sxs-lookup"><span data-stu-id="c8f11-139">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![UI 内の依存セルをExcelします。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="c8f11-141">メソッド `getDirectDependents` は、ブック間で依存セルを取得できない。</span><span class="sxs-lookup"><span data-stu-id="c8f11-141">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="c8f11-142">次のコード サンプルは、アクティブな範囲の直接の依存を取得し、それらの依存セルの背景色を黄色に変更します。</span><span class="sxs-lookup"><span data-stu-id="c8f11-142">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="c8f11-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="c8f11-143">See also</span></span>

- [<span data-ttu-id="c8f11-144">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="c8f11-144">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c8f11-145">JavaScript API を使用してセルExcelする</span><span class="sxs-lookup"><span data-stu-id="c8f11-145">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="c8f11-146">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="c8f11-146">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
