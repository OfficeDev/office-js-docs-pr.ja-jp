---
title: Excel JavaScript API を使用して範囲を挿入する
description: Excel JavaScript API を使用してセル範囲を挿入する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652922"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="f577d-103">Excel JavaScript API を使用してセル範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="f577d-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="f577d-104">この記事では、Excel JavaScript API を使用してセル範囲を挿入するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="f577d-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="f577d-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、Excel.Range クラスを参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="f577d-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="f577d-106">セルの範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="f577d-106">Insert a range of cells</span></span>

<span data-ttu-id="f577d-107">次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。</span><span class="sxs-lookup"><span data-stu-id="f577d-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="f577d-108">範囲を挿入する前のデータ</span><span class="sxs-lookup"><span data-stu-id="f577d-108">Data before range is inserted</span></span>

![範囲を挿入する前の Excel のデータ](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="f577d-110">範囲を挿入した後のデータ</span><span class="sxs-lookup"><span data-stu-id="f577d-110">Data after range is inserted</span></span>

![範囲を挿入した後の Excel のデータ](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="f577d-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="f577d-112">See also</span></span>

- [<span data-ttu-id="f577d-113">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="f577d-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f577d-114">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="f577d-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="f577d-115">Excel JavaScript API を使用して範囲をクリアまたは削除する</span><span class="sxs-lookup"><span data-stu-id="f577d-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)
