---
title: Excel JavaScript API を使用して、非バウンド範囲に対する読み取りまたは書き込み
description: Excel JavaScript API を使用して、非バウンド範囲への読み取りまたは書き込みを行う方法について説明します。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652886"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a><span data-ttu-id="6ca8b-103">Excel JavaScript API を使用して、非バウンド範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="6ca8b-103">Read or write to an unbounded range using the Excel JavaScript API</span></span>

<span data-ttu-id="6ca8b-104">この記事では、Excel JavaScript API を使用して、非バウンド範囲の読み取りおよび書き込みを行う方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-104">This article describes how to read and write to an unbounded range with the Excel JavaScript API.</span></span> <span data-ttu-id="6ca8b-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

<span data-ttu-id="6ca8b-106">非バウンド範囲アドレスは、列全体または行全体を指定する範囲アドレスです。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-106">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="6ca8b-107">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-107">For example:</span></span>

- <span data-ttu-id="6ca8b-108">列全体で構成される範囲アドレス:</span><span class="sxs-lookup"><span data-stu-id="6ca8b-108">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="6ca8b-109">行全体で構成される範囲アドレス:</span><span class="sxs-lookup"><span data-stu-id="6ca8b-109">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a><span data-ttu-id="6ca8b-110">無制限の範囲の読み取り</span><span class="sxs-lookup"><span data-stu-id="6ca8b-110">Read an unbounded range</span></span>

<span data-ttu-id="6ca8b-p103">API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`null`、`values`、`text`、または `numberFormat` などのセル レベルのプロパティに `formula` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-p103">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

## <a name="write-to-an-unbounded-range"></a><span data-ttu-id="6ca8b-113">無制限の範囲への書き込み</span><span class="sxs-lookup"><span data-stu-id="6ca8b-113">Write to an unbounded range</span></span>

<span data-ttu-id="6ca8b-114">入力要求が大きすぎるため、セル レベルのプロパティ (、 など) を非バウンド `values` `numberFormat` `formula` 範囲に設定することはできません。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-114">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="6ca8b-115">たとえば、次のコード例は、非バウンド範囲を指定しようとして `values` 無効です。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-115">For example, the following code example is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="6ca8b-116">非バウンド範囲のセル レベルのプロパティを設定しようとすると、API はエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="6ca8b-116">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a><span data-ttu-id="6ca8b-117">関連項目</span><span class="sxs-lookup"><span data-stu-id="6ca8b-117">See also</span></span>

- [<span data-ttu-id="6ca8b-118">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="6ca8b-118">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6ca8b-119">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="6ca8b-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="6ca8b-120">Excel JavaScript API を使用して大きな範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="6ca8b-120">Read or write to a large range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-large.md)
- [<span data-ttu-id="6ca8b-121">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="6ca8b-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
