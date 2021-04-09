---
title: Excel JavaScript API を使用して範囲を取得する
description: Excel JavaScript API を使用して範囲を取得する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652928"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="62499-103">Excel JavaScript API を使用して範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="62499-103">Get a range using the Excel JavaScript API</span></span>

<span data-ttu-id="62499-104">この記事では、Excel JavaScript API を使用してワークシート内の範囲を取得するさまざまな方法を示す例を示します。</span><span class="sxs-lookup"><span data-stu-id="62499-104">This article provides examples that show different ways to get a range within a worksheet using the Excel JavaScript API.</span></span> <span data-ttu-id="62499-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="62499-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a><span data-ttu-id="62499-106">アドレスによって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="62499-106">Get range by address</span></span>

<span data-ttu-id="62499-107">次のコード サンプルは **、Sample** という名前のワークシートから **アドレス B2:C5** の範囲を取得し、そのプロパティを読み込み、コンソールに `address` メッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="62499-107">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a><span data-ttu-id="62499-108">名前によって範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="62499-108">Get range by name</span></span>

<span data-ttu-id="62499-109">次のコード サンプルは、Sample という名前のワークシートからという名前の範囲を取得し、そのプロパティを読み込み、コンソール `MyRange`  `address` にメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="62499-109">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a><span data-ttu-id="62499-110">使用範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="62499-110">Get used range</span></span>

<span data-ttu-id="62499-111">次のコード サンプルは **、Sample** という名前のワークシートから使用範囲を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="62499-111">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="62499-112">使用範囲とは、値または書式設定が割り当てられているワークシート内のセルを含む、最小の範囲です。</span><span class="sxs-lookup"><span data-stu-id="62499-112">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="62499-113">ワークシート全体が空白の場合、メソッドは左上のセルだけで構成される `getUsedRange()` 範囲を返します。</span><span class="sxs-lookup"><span data-stu-id="62499-113">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a><span data-ttu-id="62499-114">範囲全体を取得する</span><span class="sxs-lookup"><span data-stu-id="62499-114">Get entire range</span></span>

<span data-ttu-id="62499-115">次のコード サンプルは **、Sample** という名前のワークシートからワークシートの範囲全体を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="62499-115">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="62499-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="62499-116">See also</span></span>

- [<span data-ttu-id="62499-117">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="62499-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="62499-118">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="62499-118">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="62499-119">Excel JavaScript API を使用して範囲を挿入する</span><span class="sxs-lookup"><span data-stu-id="62499-119">Insert a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-insert.md)
