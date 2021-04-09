---
title: Excel JavaScript API を使用して動的配列と範囲のスピルを処理する
description: Excel JavaScript API を使用して動的配列と範囲のスピルを処理する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652940"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a><span data-ttu-id="d9205-103">Excel JavaScript API を使用して動的配列とスピルを処理する</span><span class="sxs-lookup"><span data-stu-id="d9205-103">Handle dynamic arrays and spilling using the Excel JavaScript API</span></span>

<span data-ttu-id="d9205-104">この記事では、Excel JavaScript API を使用して動的配列と範囲のスピルを処理するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="d9205-104">This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API.</span></span> <span data-ttu-id="d9205-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="d9205-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="dynamic-arrays"></a><span data-ttu-id="d9205-106">動的配列</span><span class="sxs-lookup"><span data-stu-id="d9205-106">Dynamic arrays</span></span>

<span data-ttu-id="d9205-107">一部の Excel 数式は動的 [配列を返します](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。</span><span class="sxs-lookup"><span data-stu-id="d9205-107">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="d9205-108">数式の元のセルの外側にある複数のセルの値を入力します。</span><span class="sxs-lookup"><span data-stu-id="d9205-108">These fill the values of multiple cells outside of the formula's original cell.</span></span> <span data-ttu-id="d9205-109">この値のオーバーフローは"スピル" と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="d9205-109">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="d9205-110">アドインは [、Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) メソッドを使用して流出に使用される範囲を検索できます。</span><span class="sxs-lookup"><span data-stu-id="d9205-110">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="d9205-111">[\*OrNullObject バージョンも用意されています](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="d9205-111">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="d9205-112">次のサンプルは、セルに範囲の内容をコピーする基本的な数式を示しています。これは隣接するセルに流出します。</span><span class="sxs-lookup"><span data-stu-id="d9205-112">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="d9205-113">その後、アドインは流出を含む範囲をログに記録します。</span><span class="sxs-lookup"><span data-stu-id="d9205-113">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a><span data-ttu-id="d9205-114">範囲の流出</span><span class="sxs-lookup"><span data-stu-id="d9205-114">Range spilling</span></span>

<span data-ttu-id="d9205-115">[Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--)メソッドを使用して、特定のセルにこぼれるセルを検索します。</span><span class="sxs-lookup"><span data-stu-id="d9205-115">Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="d9205-116">range オブジェクト `getSpillParent` が 1 つのセルの場合にのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="d9205-116">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="d9205-117">複数 `getSpillParent` のセルを含む範囲を呼び出す場合、エラーがスローされます (または null の範囲が返されます `Range.getSpillParentOrNullObject` )。</span><span class="sxs-lookup"><span data-stu-id="d9205-117">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="d9205-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="d9205-118">See also</span></span>

- [<span data-ttu-id="d9205-119">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="d9205-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d9205-120">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="d9205-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="d9205-121">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="d9205-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
