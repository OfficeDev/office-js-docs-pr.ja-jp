---
title: Excel JavaScript API を使用して選択した範囲を設定および取得する
description: Excel JavaScript API を使用して、Excel JavaScript API を使用して範囲を設定および取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652881"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="2178d-103">Excel JavaScript API を使用して範囲を設定および取得する</span><span class="sxs-lookup"><span data-stu-id="2178d-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="2178d-104">この記事では、Excel JavaScript API を使用して範囲を設定および取得するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="2178d-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="2178d-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="2178d-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="2178d-106">選択範囲を設定する</span><span class="sxs-lookup"><span data-stu-id="2178d-106">Set the selected range</span></span>

<span data-ttu-id="2178d-107">次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。</span><span class="sxs-lookup"><span data-stu-id="2178d-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="2178d-108">選択範囲 B2:E6</span><span class="sxs-lookup"><span data-stu-id="2178d-108">Selected range B2:E6</span></span>

![Excel の選択範囲](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="2178d-110">選択範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="2178d-110">Get the selected range</span></span>

<span data-ttu-id="2178d-111">次のコード サンプルでは、選択した範囲を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="2178d-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2178d-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="2178d-112">See also</span></span>

- [<span data-ttu-id="2178d-113">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="2178d-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="2178d-114">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="2178d-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="2178d-115">Excel JavaScript API を使用して範囲の値、テキスト、または数式を設定および取得する</span><span class="sxs-lookup"><span data-stu-id="2178d-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="2178d-116">Excel JavaScript API を使用して範囲の形式を設定する</span><span class="sxs-lookup"><span data-stu-id="2178d-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
