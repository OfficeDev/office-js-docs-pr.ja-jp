---
title: Excel JavaScript API を使用して大きな範囲に対する読み取りまたは書き込み
description: Excel JavaScript API を使用して、大きな範囲を読み取りまたは書き込む方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652917"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a><span data-ttu-id="018c6-103">Excel JavaScript API を使用して大きな範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="018c6-103">Read or write to a large range using the Excel JavaScript API</span></span>

<span data-ttu-id="018c6-104">この記事では、Excel JavaScript API を使用して大きな範囲への読み取りおよび書き込みを処理する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="018c6-104">This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.</span></span>

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a><span data-ttu-id="018c6-105">大きな範囲に対して個別の読み取り操作または書き込み操作を実行する</span><span class="sxs-lookup"><span data-stu-id="018c6-105">Run separate read or write operations for large ranges</span></span>

<span data-ttu-id="018c6-106">範囲に多数のセル、値、数値形式、または数式が含まれている場合、その範囲で API 操作を実行できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="018c6-106">If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="018c6-107">API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。</span><span class="sxs-lookup"><span data-stu-id="018c6-107">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="018c6-108">このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="018c6-108">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="018c6-109">システムの制限の詳細については、「リソースの制限とパフォーマンスの最適化」の「Excel アドイン」セクションを参照Office [してください](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="018c6-109">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="018c6-110">範囲の条件付き書式</span><span class="sxs-lookup"><span data-stu-id="018c6-110">Conditional formatting of ranges</span></span>

<span data-ttu-id="018c6-111">範囲には、条件に基づいて個々のセルに適用する書式設定を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="018c6-111">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="018c6-112">この詳細については、「[Excel の範囲に条件付き書式を適用する](excel-add-ins-conditional-formatting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="018c6-112">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="018c6-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="018c6-113">See also</span></span>

- [<span data-ttu-id="018c6-114">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="018c6-114">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="018c6-115">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="018c6-115">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="018c6-116">Excel JavaScript API を使用して、非バウンド範囲に対する読み取りまたは書き込み</span><span class="sxs-lookup"><span data-stu-id="018c6-116">Read or write to an unbounded range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-unbounded.md)
- [<span data-ttu-id="018c6-117">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="018c6-117">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
