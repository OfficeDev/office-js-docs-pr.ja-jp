---
ms.date: 01/14/2020
description: 揮発性およびオフラインのストリーミング カスタム関数を実装する方法について説明します。
title: 関数の揮発性の値
localization_priority: Normal
ms.openlocfilehash: f441ef4fb7f90add5318546e3ccf4cc8bc60a8cf
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075888"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="85b9a-103">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="85b9a-103">Volatile values in functions</span></span>

<span data-ttu-id="85b9a-104">揮発性関数は、セルが計算されるごとに値が変化する関数です。</span><span class="sxs-lookup"><span data-stu-id="85b9a-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="85b9a-105">関数の引数が変更された場合でも、値は変更できます。</span><span class="sxs-lookup"><span data-stu-id="85b9a-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="85b9a-106">これらの関数は、Excel が再計算するたびに再計算を行います。</span><span class="sxs-lookup"><span data-stu-id="85b9a-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="85b9a-107">たとえば、`NOW` 関数を呼び出すセルがあるとします。</span><span class="sxs-lookup"><span data-stu-id="85b9a-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="85b9a-108">`NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。</span><span class="sxs-lookup"><span data-stu-id="85b9a-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="85b9a-109">Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="85b9a-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="85b9a-110">Excel の揮発性関数の完全なリストは、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="85b9a-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="85b9a-111">カスタム関数を使用すると、独自の揮発性関数を作成できます。これは、日付、時刻、乱数、およびモデリングを処理するときに役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="85b9a-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="85b9a-112">たとえば、 [モンテカルロシミュレーションでは、](https://en.wikipedia.org/wiki/Monte_Carlo_method) 最適なソリューションを決定するためにランダムな入力を生成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="85b9a-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="85b9a-113">JSON ファイルの自動生成を選択する場合は、JSDoc コメント タグを使用して揮発性関数を宣言します `@volatile` 。</span><span class="sxs-lookup"><span data-stu-id="85b9a-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="85b9a-114">自動生成の詳細については、「カスタム関数の [JSON メタデータの自動生成」を参照してください](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="85b9a-114">From more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="85b9a-115">揮発性のカスタム関数の例を次に示します。これは、6 辺のサイコロの回転をシミュレートします。</span><span class="sxs-lookup"><span data-stu-id="85b9a-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![ランダムな値を返すカスタム関数を示す GIF を使用して、6 辺のサイコロのローリングをシミュレートします。](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a><span data-ttu-id="85b9a-117">次の手順</span><span class="sxs-lookup"><span data-stu-id="85b9a-117">Next steps</span></span>
* <span data-ttu-id="85b9a-118">カスタム関数 [パラメーター オプションについて説明します](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="85b9a-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="85b9a-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="85b9a-119">See also</span></span>

* [<span data-ttu-id="85b9a-120">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="85b9a-120">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="85b9a-121">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="85b9a-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
