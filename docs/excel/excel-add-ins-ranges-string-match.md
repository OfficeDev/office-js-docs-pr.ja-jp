---
title: Excel JavaScript API を使用して文字列を検索する
description: Excel JavaScript API を使用して範囲内の文字列を検索する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652887"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="78e3a-103">Excel JavaScript API を使用して範囲内の文字列を検索する</span><span class="sxs-lookup"><span data-stu-id="78e3a-103">Find a string within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="78e3a-104">この記事では、Excel JavaScript API を使用して範囲内の文字列を検索するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="78e3a-104">This article provides a code sample that finds a string within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="78e3a-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="78e3a-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a><span data-ttu-id="78e3a-106">範囲内の文字列と一致する</span><span class="sxs-lookup"><span data-stu-id="78e3a-106">Match a string within a range</span></span>

<span data-ttu-id="78e3a-107">`Range` オブジェクトには、範囲内で指定された文字列を検索するための `find` メソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="78e3a-107">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="78e3a-108">このメソッドは、一致するテキストがある最初のセルの範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="78e3a-108">It returns the range of the first cell with matching text.</span></span>

<span data-ttu-id="78e3a-109">次のコード サンプルは、文字列 **Food** と等しい値を持つ最初のセルを検索して、そのアドレスをコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="78e3a-109">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="78e3a-110">指定した文字列が範囲に存在しない場合、`ItemNotFound` エラーが `find` によってスローされます。</span><span class="sxs-lookup"><span data-stu-id="78e3a-110">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="78e3a-111">指定した文字列が範囲に存在しない可能性がある場合は、自分のコードで適切にシナリオを処理できるように、[findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) メソッドを使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="78e3a-111">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="78e3a-112">単一のセルを表す範囲に対して `find` メソッドが呼び出されると、ワークシート全体が検索されます。</span><span class="sxs-lookup"><span data-stu-id="78e3a-112">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="78e3a-113">検索はその単一のセルから始まり、`SearchCriteria.searchDirection` によって指定された方向へ行われ、場合によってはワークシートの最終部分で折り返されます。</span><span class="sxs-lookup"><span data-stu-id="78e3a-113">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="78e3a-114">関連項目</span><span class="sxs-lookup"><span data-stu-id="78e3a-114">See also</span></span>

- [<span data-ttu-id="78e3a-115">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="78e3a-115">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="78e3a-116">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="78e3a-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="78e3a-117">Excel JavaScript API を使用して範囲内の特別なセルを検索する</span><span class="sxs-lookup"><span data-stu-id="78e3a-117">Find special cells within a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-special-cells.md)
