---
title: JavaScript API を使用して範囲Excel挿入する
description: JavaScript API を使用してセル範囲を挿入するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0571e7d6140f5023008654a1e74d7abf6b3cab0a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075783"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="b6596-103">JavaScript API を使用してセル範囲をExcelする</span><span class="sxs-lookup"><span data-stu-id="b6596-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="b6596-104">この記事では、JavaScript API を使用してセル範囲を挿入するコード サンプルExcel示します。</span><span class="sxs-lookup"><span data-stu-id="b6596-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="b6596-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については、 `Range` 次の[Excel。Range クラス](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="b6596-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="b6596-106">セルの範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="b6596-106">Insert a range of cells</span></span>

<span data-ttu-id="b6596-107">次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。</span><span class="sxs-lookup"><span data-stu-id="b6596-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="b6596-108">範囲を挿入する前のデータ</span><span class="sxs-lookup"><span data-stu-id="b6596-108">Data before range is inserted</span></span>

![範囲が挿入Excel前のデータ。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="b6596-110">範囲を挿入した後のデータ</span><span class="sxs-lookup"><span data-stu-id="b6596-110">Data after range is inserted</span></span>

![範囲が挿入Excel後のデータ。](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="b6596-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="b6596-112">See also</span></span>

- [<span data-ttu-id="b6596-113">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="b6596-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b6596-114">JavaScript API を使用してセルExcelする</span><span class="sxs-lookup"><span data-stu-id="b6596-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="b6596-115">JavaScript API を使用して範囲をクリアまたはExcelする</span><span class="sxs-lookup"><span data-stu-id="b6596-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)
