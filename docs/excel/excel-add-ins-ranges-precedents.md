---
title: Excel JavaScript API を使用して数式の前例を使用する
description: Excel JavaScript API を使用して数式の前例を取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652916"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a><span data-ttu-id="f9a36-103">Excel JavaScript API を使用して数式の前例を取得する</span><span class="sxs-lookup"><span data-stu-id="f9a36-103">Get formula precedents using the Excel JavaScript API</span></span>

<span data-ttu-id="f9a36-104">この記事では、Excel JavaScript API を使用して数式の前例を取得するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="f9a36-104">This article provides a code sample that retrieves formula precedents using the Excel JavaScript API.</span></span> <span data-ttu-id="f9a36-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="f9a36-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="get-formula-precedents"></a><span data-ttu-id="f9a36-106">数式の前例を取得する</span><span class="sxs-lookup"><span data-stu-id="f9a36-106">Get formula precedents</span></span>

<span data-ttu-id="f9a36-107">Excel の数式は、多くの場合、他のセルを参照します。</span><span class="sxs-lookup"><span data-stu-id="f9a36-107">An Excel formula often refers to other cells.</span></span> <span data-ttu-id="f9a36-108">セルが数式にデータを提供する場合、セルは数式 "前例" と呼ばれる。</span><span class="sxs-lookup"><span data-stu-id="f9a36-108">When a cell provides data to a formula, it is known as a formula "precedent".</span></span> <span data-ttu-id="f9a36-109">セル間のリレーションシップに関連する Excel 機能の詳細については、「数式とセル間のリレーションシップを表示する [」を参照してください](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。</span><span class="sxs-lookup"><span data-stu-id="f9a36-109">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span> 

<span data-ttu-id="f9a36-110">[Range.getDirectPrecedents を](/javascript/api/excel/excel.range#getdirectprecedents--)使用すると、アドインは数式の直接の先行セルを検索できます。</span><span class="sxs-lookup"><span data-stu-id="f9a36-110">With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells.</span></span> <span data-ttu-id="f9a36-111">`Range.getDirectPrecedents` オブジェクトを返 `WorkbookRangeAreas` します。</span><span class="sxs-lookup"><span data-stu-id="f9a36-111">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="f9a36-112">このオブジェクトには、ブック内のすべての前例のアドレスが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f9a36-112">This object contains the addresses of all the precedents in the workbook.</span></span> <span data-ttu-id="f9a36-113">このオブジェクトには、少なくとも 1 つの数式の前例を含 `RangeAreas` むワークシートごとに個別のオブジェクトがあります。</span><span class="sxs-lookup"><span data-stu-id="f9a36-113">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="f9a36-114">オブジェクトの操作の詳細については、「Excel アドインで複数の範囲を同時に操作する `RangeAreas` [」を参照してください](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="f9a36-114">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="f9a36-115">Excel UI では、[前 **例のトレース** ] ボタンは、前のセルから選択した数式に矢印を描画します。</span><span class="sxs-lookup"><span data-stu-id="f9a36-115">In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula.</span></span> <span data-ttu-id="f9a36-116">Excel UI ボタンとは異なり、 `getDirectPrecedents` メソッドは矢印を描画しない。</span><span class="sxs-lookup"><span data-stu-id="f9a36-116">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="f9a36-117">メソッド `getDirectPrecedents` は、ブック間で先行セルを取得できない。</span><span class="sxs-lookup"><span data-stu-id="f9a36-117">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span> 

<span data-ttu-id="f9a36-118">次のコード サンプルでは、アクティブな範囲の直接の前例を取得し、それらの前のセルの背景色を黄色に変更します。</span><span class="sxs-lookup"><span data-stu-id="f9a36-118">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span> 

> [!NOTE]
> <span data-ttu-id="f9a36-119">強調表示が適切に機能するには、アクティブな範囲に同じブック内の他のセルを参照する数式が含まれている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f9a36-119">The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.</span></span> 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="f9a36-120">関連項目</span><span class="sxs-lookup"><span data-stu-id="f9a36-120">See also</span></span>

- [<span data-ttu-id="f9a36-121">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="f9a36-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f9a36-122">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="f9a36-122">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="f9a36-123">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="f9a36-123">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
