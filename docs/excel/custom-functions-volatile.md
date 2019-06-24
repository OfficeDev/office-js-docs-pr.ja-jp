---
ms.date: 06/17/2019
description: 揮発性およびオフラインのストリーミングカスタム関数を実装する方法について説明します。
title: 関数の揮発性の値
localization_priority: Normal
ms.openlocfilehash: bcaef092ec386a7d80760c1e2a567b9de1fdad21
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127820"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="2d892-103">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="2d892-103">Volatile values in functions</span></span>

<span data-ttu-id="2d892-104">Volatile 関数は、セルが計算されるたびに値が変更される関数です。</span><span class="sxs-lookup"><span data-stu-id="2d892-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="2d892-105">この値は、関数の引数が変更されていない場合でも変更できます。</span><span class="sxs-lookup"><span data-stu-id="2d892-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="2d892-106">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="2d892-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="2d892-107">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="2d892-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="2d892-108">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="2d892-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="2d892-109">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="2d892-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="2d892-110">Excel のすべての揮発性関数の一覧は、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="2d892-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="2d892-111">カスタム関数を使用すると、独自の揮発性関数を作成することができます。これは、日付、時刻、乱数、およびモデリングを処理するときに便利です。</span><span class="sxs-lookup"><span data-stu-id="2d892-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="2d892-112">たとえば、[モンテカルロモンテカルロシミュレーション](https://en.wikipedia.org/wiki/Monte_Carlo_method)では、最適なソリューションを決定するためにランダムな入力を生成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d892-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="2d892-113">JSON ファイルの自動生成を選択する場合は、JSDOC comment タグ`@volatile`を使用して揮発性関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="2d892-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="2d892-114">Autogeneration の詳細については、「[カスタム関数の JSON メタデータの作成](custom-functions-json-autogeneration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d892-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="2d892-115">揮発性のカスタム関数の例を次に示します。これは6つのサイドダイスの重ね合わせをシミュレートします。</span><span class="sxs-lookup"><span data-stu-id="2d892-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a><span data-ttu-id="2d892-116">次のステップ</span><span class="sxs-lookup"><span data-stu-id="2d892-116">Next steps</span></span>
<span data-ttu-id="2d892-117">[カスタム関数に状態を保存](custom-functions-save-state.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="2d892-117">Learn how to [save state in your custom functions](custom-functions-save-state.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2d892-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="2d892-118">See also</span></span>

* [<span data-ttu-id="2d892-119">カスタム関数のパラメータオプション</span><span class="sxs-lookup"><span data-stu-id="2d892-119">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="2d892-120">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="2d892-120">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2d892-121">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="2d892-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
