---
ms.date: 04/30/2019
description: 揮発性およびオフラインのストリーミングカスタム関数を実装する方法について説明します。
title: 関数内の揮発性値 (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 63618adecff57398e1630e6b5ab43c0dbc753b36
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527319"
---
## <a name="volatile-values-in-functions"></a><span data-ttu-id="0e5d5-103">関数内の揮発性値</span><span class="sxs-lookup"><span data-stu-id="0e5d5-103">Volatile values in functions</span></span>

<span data-ttu-id="0e5d5-104">Volatile 関数は、セルが計算されるたびに値が変更される関数です。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="0e5d5-105">この値は、関数の引数が変更されていない場合でも変更できます。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="0e5d5-106">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="0e5d5-107">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="0e5d5-108">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="0e5d5-109">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="0e5d5-110">Excel のすべての揮発性関数の一覧は、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="0e5d5-111">カスタム関数を使用すると、独自の揮発性関数を作成することができます。これは、日付、時刻、乱数、およびモデリングを処理するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="0e5d5-112">たとえば、[モンテカルロモンテカルロシミュレーション](https://en.wikipedia.org/wiki/Monte_Carlo_method
)では、最適なソリューションを決定するためにランダムな入力を生成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="0e5d5-113">JSON ファイルの自動生成を選択する場合は、JSDOC comment タグ`@volatile`を使用して揮発性関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="0e5d5-114">Autogeneration の詳細については、「[カスタム関数の JSON メタデータの作成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e5d5-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0e5d5-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="0e5d5-115">See also</span></span>

* [<span data-ttu-id="0e5d5-116">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0e5d5-116">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0e5d5-117">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0e5d5-117">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0e5d5-118">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0e5d5-118">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0e5d5-119">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="0e5d5-119">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="0e5d5-120">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0e5d5-120">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
