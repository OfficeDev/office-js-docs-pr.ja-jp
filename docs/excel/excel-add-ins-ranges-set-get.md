---
title: JavaScript API を使用して選択した範囲を設定Excel取得する
description: JavaScript API を使用して、Excel JavaScript API を使用して選択した範囲を設定および取得するExcel説明します。
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 623ba5c1b9e76151d4a2c4b169e655236b37e8c8
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290783"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="f7fbb-103">JavaScript API を使用して選択した範囲を設定Excel取得する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="f7fbb-104">この記事では、JavaScript API を使用して選択した範囲を設定して取得するExcel説明します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="f7fbb-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="f7fbb-106">選択範囲を設定する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-106">Set the selected range</span></span>

<span data-ttu-id="f7fbb-107">次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="f7fbb-108">選択範囲 B2:E6</span><span class="sxs-lookup"><span data-stu-id="f7fbb-108">Selected range B2:E6</span></span>

![[選択した範囲] Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="f7fbb-110">選択範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-110">Get the selected range</span></span>

<span data-ttu-id="f7fbb-111">次のコード サンプルでは、選択した範囲を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="select-the-edge-of-a-used-range"></a><span data-ttu-id="f7fbb-112">使用範囲の端を選択する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-112">Select the edge of a used range</span></span>

<span data-ttu-id="f7fbb-113">[Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_)メソッドと[Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_)メソッドを使用すると、アドインはキーボード選択ショートカットの動作をレプリケートし、現在選択されている範囲に基づいて使用範囲のエッジを選択できます。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-113">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="f7fbb-114">使用範囲の詳細については、「使用範囲の [取得」を参照してください](excel-add-ins-ranges-get.md#get-used-range)。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-114">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="f7fbb-115">次のスクリーンショットでは、使用される範囲は、各セルの **値が C5:F12 のテーブルです**。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-115">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="f7fbb-116">この表の外側の空のセルは、使用範囲の外側です。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-116">The empty cells outside this table are outside the used range.</span></span>

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="f7fbb-118">現在使用されている範囲の端にあるセルを選択する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-118">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="f7fbb-119">次のコード サンプルは、メソッドを使用して、現在使用されている範囲の最も遠い端にあるセルを上方向 `Range.getRangeEdge` に選択する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-119">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="f7fbb-120">このアクションは、範囲が選択されている間に Ctrl + 上矢印キーのキーボード ショートカットを使用した結果と一致します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-120">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="f7fbb-121">使用範囲の端にあるセルを選択する前に</span><span class="sxs-lookup"><span data-stu-id="f7fbb-121">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="f7fbb-122">次のスクリーンショットは、使用範囲と、使用範囲内で選択した範囲を示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-122">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="f7fbb-123">使用範囲は **、C5:F12** のデータを含むテーブルです。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-123">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="f7fbb-124">この表の中で、 **範囲 D8:E9 が** 選択されています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-124">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="f7fbb-125">この選択は、 *メソッドを実行* する前の前の状態 `Range.getRangeEdge` です。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-125">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="f7fbb-128">使用範囲の端にあるセルを選択した後</span><span class="sxs-lookup"><span data-stu-id="f7fbb-128">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="f7fbb-129">次のスクリーンショットは、前のスクリーンショットと同じ表を示し **、C5:F12** の範囲のデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-129">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="f7fbb-130">この表の中で、 **範囲 D5 が** 選択されています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-130">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="f7fbb-131">この選択は *、メソッド* を実行した後の状態の後で、使用範囲の端にあるセルを上方向 `Range.getRangeEdge` に選択します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-131">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="f7fbb-134">現在の範囲から使用範囲の最も遠い端までのすべてのセルを選択する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-134">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="f7fbb-135">次のコード サンプルは、メソッドを使用して、現在選択されている範囲から使用範囲の最も遠い端まで、下方向のすべてのセルを選択する方法 `Range.getExtendedRange` を示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-135">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="f7fbb-136">このアクションは、範囲が選択されている間に Ctrl + Shift +下矢印キーのキーボード ショートカットを使用した結果と一致します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-136">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="f7fbb-137">現在の範囲から使用範囲の端までのすべてのセルを選択する前に</span><span class="sxs-lookup"><span data-stu-id="f7fbb-137">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="f7fbb-138">次のスクリーンショットは、使用範囲と、使用範囲内で選択した範囲を示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-138">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="f7fbb-139">使用範囲は **、C5:F12** のデータを含むテーブルです。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-139">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="f7fbb-140">この表の中で、 **範囲 D8:E9 が** 選択されています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-140">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="f7fbb-141">この選択は、 *メソッドを実行* する前の前の状態 `Range.getExtendedRange` です。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-141">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="f7fbb-144">現在の範囲から使用範囲の端までのすべてのセルを選択した後</span><span class="sxs-lookup"><span data-stu-id="f7fbb-144">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="f7fbb-145">次のスクリーンショットは、前のスクリーンショットと同じ表を示し **、C5:F12** の範囲のデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-145">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="f7fbb-146">この表の中で、 **範囲 D8:E12 が** 選択されています。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-146">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="f7fbb-147">この選択は *、メソッド* を実行した後の状態の後で、現在の範囲から下方向の使用範囲の端までのすべてのセル `Range.getExtendedRange` を選択します。</span><span class="sxs-lookup"><span data-stu-id="f7fbb-147">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![C5:F12 のデータが含Excel。](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="f7fbb-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="f7fbb-150">See also</span></span>

- [<span data-ttu-id="f7fbb-151">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="f7fbb-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f7fbb-152">JavaScript API を使用してセルExcelする</span><span class="sxs-lookup"><span data-stu-id="f7fbb-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="f7fbb-153">JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-153">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="f7fbb-154">JavaScript API を使用して範囲Excel設定する</span><span class="sxs-lookup"><span data-stu-id="f7fbb-154">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
