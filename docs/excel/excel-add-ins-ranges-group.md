---
title: Excel JavaScript API を使用して範囲をグループ化する
description: Excel JavaScript API を使用して範囲の行または列をグループ化してアウトラインを作成する方法について説明します。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 32f65cf88c23bd6368b37318d3ba20fde95b8436
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652923"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a><span data-ttu-id="e49d6-103">Excel JavaScript API を使用してアウトラインの範囲をグループ化する</span><span class="sxs-lookup"><span data-stu-id="e49d6-103">Group ranges for an outline using the Excel JavaScript API</span></span>

<span data-ttu-id="e49d6-104">この記事では、Excel JavaScript API を使用してアウトラインの範囲をグループ化する方法を示すコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-104">This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API.</span></span> <span data-ttu-id="e49d6-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="e49d6-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a><span data-ttu-id="e49d6-106">アウトラインの範囲の行または列をグループ化する</span><span class="sxs-lookup"><span data-stu-id="e49d6-106">Group rows or columns of a range for an outline</span></span>

<span data-ttu-id="e49d6-107">範囲の行または列をグループ化してアウトラインを作成 [できます](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。</span><span class="sxs-lookup"><span data-stu-id="e49d6-107">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="e49d6-108">これらのグループを折りたたみ、展開して、対応するセルを非表示にし、表示できます。</span><span class="sxs-lookup"><span data-stu-id="e49d6-108">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="e49d6-109">これにより、トップライン データの迅速な分析が容易になります。</span><span class="sxs-lookup"><span data-stu-id="e49d6-109">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="e49d6-110">[Range.group を使用して](/javascript/api/excel/excel.range#group-groupoption-)、これらのアウトライン グループを作成します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="e49d6-111">アウトラインには階層を含め、小さなグループは大きなグループの下に入れ子にできます。</span><span class="sxs-lookup"><span data-stu-id="e49d6-111">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="e49d6-112">これにより、アウトラインをさまざまなレベルで表示できます。</span><span class="sxs-lookup"><span data-stu-id="e49d6-112">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="e49d6-113">表示されるアウトライン レベルを変更するには [、Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) メソッドを使用してプログラムを使用します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-113">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="e49d6-114">Excel は 8 つのレベルのアウトライン グループのみをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="e49d6-114">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="e49d6-115">次のコード サンプルでは、行と列の両方に 2 つのレベルのグループを含むアウトラインを作成します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-115">The following code sample creates an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="e49d6-116">次の図は、そのアウトラインのグループ化を示しています。</span><span class="sxs-lookup"><span data-stu-id="e49d6-116">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="e49d6-117">コード サンプルでは、グループ化されている範囲にアウトライン コントロールの行または列 (この例の "Totals") は含めされません。</span><span class="sxs-lookup"><span data-stu-id="e49d6-117">In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="e49d6-118">グループは、コントロールの行または列ではなく、折りたたむものを定義します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-118">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![2 つのレベルの 2 次元アウトラインを持つ範囲](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a><span data-ttu-id="e49d6-120">範囲の行または列からグループ化を削除する</span><span class="sxs-lookup"><span data-stu-id="e49d6-120">Remove grouping from rows or columns of a range</span></span>

<span data-ttu-id="e49d6-121">行または列グループのグループ化を解除するには [、Range.ungroup メソッドを使用](/javascript/api/excel/excel.range#ungroup-groupoption-) します。</span><span class="sxs-lookup"><span data-stu-id="e49d6-121">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="e49d6-122">これにより、アウトラインから最も外側のレベルが削除されます。</span><span class="sxs-lookup"><span data-stu-id="e49d6-122">This removes the outermost level from the outline.</span></span> <span data-ttu-id="e49d6-123">同じ行または列の種類の複数のグループが指定した範囲内で同じレベルにある場合、それらのグループはすべてグループ化解除されます。</span><span class="sxs-lookup"><span data-stu-id="e49d6-123">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="e49d6-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="e49d6-124">See also</span></span>

- [<span data-ttu-id="e49d6-125">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="e49d6-125">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="e49d6-126">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="e49d6-126">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="e49d6-127">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="e49d6-127">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
