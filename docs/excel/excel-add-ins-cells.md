---
title: Excel JavaScript API を使用してセルを使用します。
description: セルの Excel JavaScript API 定義について説明し、セルを使用する方法について説明します。
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652977"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="00d28-103">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="00d28-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="00d28-104">Excel JavaScript API には、"Cell" オブジェクトまたはクラスが含か指定されています。</span><span class="sxs-lookup"><span data-stu-id="00d28-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="00d28-105">代わりに、すべての Excel セルはオブジェクト `Range` です。</span><span class="sxs-lookup"><span data-stu-id="00d28-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="00d28-106">Excel UI の個々のセルは、Excel JavaScript API で 1 つのセルを持つオブジェクト `Range` に変換されます。</span><span class="sxs-lookup"><span data-stu-id="00d28-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="00d28-107">オブジェクト `Range` には、複数の連続するセルを含め、複数のセルを含めできます。</span><span class="sxs-lookup"><span data-stu-id="00d28-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="00d28-108">連続するセルは、(単一の行または列を含む) 未結合の四角形を形成します。</span><span class="sxs-lookup"><span data-stu-id="00d28-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="00d28-109">連続していないセルの操作については、「RangeAreas オブジェクトを使用して不連続セルを操作する [」を参照してください](#work-with-discontiguous-cells-using-the-rangeareas-object)。</span><span class="sxs-lookup"><span data-stu-id="00d28-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="00d28-110">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="00d28-110">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="excel-javascript-apis-that-mention-cells"></a><span data-ttu-id="00d28-111">セルを言及する Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="00d28-111">Excel JavaScript APIs that mention cells</span></span>

<span data-ttu-id="00d28-112">Excel JavaScript API に "Cell" オブジェクトまたはクラスがない場合でも、多数の API 名がセルを表します。</span><span class="sxs-lookup"><span data-stu-id="00d28-112">Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells.</span></span> <span data-ttu-id="00d28-113">これらの API は、色、テキストの書式設定、フォントなど、セルのプロパティを制御します。</span><span class="sxs-lookup"><span data-stu-id="00d28-113">These APIs control cell properties like color, text formatting, and font.</span></span>

<span data-ttu-id="00d28-114">Excel JavaScript API の次のリストは、セルを参照します。</span><span class="sxs-lookup"><span data-stu-id="00d28-114">The following list of Excel JavaScript APIs refer to cells.</span></span>

- [<span data-ttu-id="00d28-115">CellBorder</span><span class="sxs-lookup"><span data-stu-id="00d28-115">CellBorder</span></span>](/javascript/api/excel/excel.cellborder)
- [<span data-ttu-id="00d28-116">CellBorderCollection</span><span class="sxs-lookup"><span data-stu-id="00d28-116">CellBorderCollection</span></span>](/javascript/api/excel/excel.cellbordercollection)
- [<span data-ttu-id="00d28-117">CellProperties</span><span class="sxs-lookup"><span data-stu-id="00d28-117">CellProperties</span></span>](/javascript/api/excel/excel.cellproperties)
- [<span data-ttu-id="00d28-118">CellPropertiesFill</span><span class="sxs-lookup"><span data-stu-id="00d28-118">CellPropertiesFill</span></span>](/javascript/api/excel/excel.cellpropertiesfill)
- [<span data-ttu-id="00d28-119">CellPropertiesFont</span><span class="sxs-lookup"><span data-stu-id="00d28-119">CellPropertiesFont</span></span>](/javascript/api/excel/excel.cellpropertiesfont)
- [<span data-ttu-id="00d28-120">CellPropertiesFormat</span><span class="sxs-lookup"><span data-stu-id="00d28-120">CellPropertiesFormat</span></span>](/javascript/api/excel/excel.cellpropertiesformat)
- [<span data-ttu-id="00d28-121">CellPropertiesProtection</span><span class="sxs-lookup"><span data-stu-id="00d28-121">CellPropertiesProtection</span></span>](/javascript/api/excel/excel.cellpropertiesprotection)
- [<span data-ttu-id="00d28-122">CellValueConditionalFormat</span><span class="sxs-lookup"><span data-stu-id="00d28-122">CellValueConditionalFormat</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)
- [<span data-ttu-id="00d28-123">ConditionalCellValueRule</span><span class="sxs-lookup"><span data-stu-id="00d28-123">ConditionalCellValueRule</span></span>](/javascript/api/excel/excel.conditionalcellvaluerule)
- [<span data-ttu-id="00d28-124">SettableCellProperties</span><span class="sxs-lookup"><span data-stu-id="00d28-124">SettableCellProperties</span></span>](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="00d28-125">RangeAreas オブジェクトを使用して不一視セルを使用する</span><span class="sxs-lookup"><span data-stu-id="00d28-125">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="00d28-126">[RangeAreas オブジェクトを](/javascript/api/excel/excel.rangeareas)使用すると、アドインは複数の範囲に対して一度に操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="00d28-126">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="00d28-127">これらの範囲は連続している可能性がありますが、必要はありません。</span><span class="sxs-lookup"><span data-stu-id="00d28-127">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="00d28-128">`RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。</span><span class="sxs-lookup"><span data-stu-id="00d28-128">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="00d28-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="00d28-129">See also</span></span>

- [<span data-ttu-id="00d28-130">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="00d28-130">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="00d28-131">Excel JavaScript API を使用して範囲を取得する</span><span class="sxs-lookup"><span data-stu-id="00d28-131">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="00d28-132">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="00d28-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
