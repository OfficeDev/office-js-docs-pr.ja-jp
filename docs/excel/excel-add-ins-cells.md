---
title: Excel JavaScript API を使用してセルを使用します。
description: セルの Excel JavaScript API 定義について説明し、セルを使用する方法について説明します。
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917101"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="fca32-103">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="fca32-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="fca32-104">Excel JavaScript API には、"Cell" オブジェクトまたはクラスがありません。</span><span class="sxs-lookup"><span data-stu-id="fca32-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="fca32-105">代わりに、すべての Excel セルはオブジェクト `Range` です。</span><span class="sxs-lookup"><span data-stu-id="fca32-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="fca32-106">Excel UI の個々のセルは、Excel JavaScript API の 1 つのセルを持つ `Range` オブジェクトに変換されます。</span><span class="sxs-lookup"><span data-stu-id="fca32-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="fca32-107">オブジェクト `Range` には、複数の連続するセルを含め、複数のセルを含めできます。</span><span class="sxs-lookup"><span data-stu-id="fca32-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="fca32-108">連続するセルは、(単一の行または列を含む) 未結合の四角形を形成します。</span><span class="sxs-lookup"><span data-stu-id="fca32-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="fca32-109">連続していないセルの操作については、「RangeAreas オブジェクトを使用して不連続セルを操作する [」を参照してください](#work-with-discontiguous-cells-using-the-rangeareas-object)。</span><span class="sxs-lookup"><span data-stu-id="fca32-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="fca32-110">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="fca32-110">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="fca32-111">RangeAreas オブジェクトを使用して不一視セルを使用する</span><span class="sxs-lookup"><span data-stu-id="fca32-111">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="fca32-112">[RangeAreas オブジェクトを](/javascript/api/excel/excel.rangeareas)使用すると、アドインは複数の範囲に対して一度に操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="fca32-112">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="fca32-113">これらの範囲は連続している可能性がありますが、必要はありません。</span><span class="sxs-lookup"><span data-stu-id="fca32-113">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="fca32-114">`RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。</span><span class="sxs-lookup"><span data-stu-id="fca32-114">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fca32-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="fca32-115">See also</span></span>

- [<span data-ttu-id="fca32-116">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="fca32-116">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="fca32-117">Excel JavaScript API を使用して範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="fca32-117">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="fca32-118">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="fca32-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
