---
ms.date: 12/18/2019
description: Office Excel アドインで、カスタム関数から複数の結果を返します。
title: カスタム関数から複数の結果を返す
localization_priority: Normal
ms.openlocfilehash: 687ffcd66cff16d92fec372a778fe94bad7b38d5
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2020
ms.locfileid: "40970379"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="63b45-103">カスタム関数から複数の結果を返す</span><span class="sxs-lookup"><span data-stu-id="63b45-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="63b45-104">隣接するセルに返される、カスタム関数から複数の結果を返すことができます。</span><span class="sxs-lookup"><span data-stu-id="63b45-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="63b45-105">この動作は spilling と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="63b45-105">This behavior is called spilling.</span></span> <span data-ttu-id="63b45-106">カスタム関数が結果の配列を返す場合は、動的配列数式と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="63b45-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="63b45-107">Excel の動的配列数式の詳細については、「動的配列」[および「こぼれた配列の動作](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63b45-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="63b45-108">次の図は、**並べ替え**関数が隣接するセルにどのように表示されるかを示しています。</span><span class="sxs-lookup"><span data-stu-id="63b45-108">The following image shows how the **SORT** function spills down into neighboring cells.</span></span> <span data-ttu-id="63b45-109">カスタム関数は、次のような複数の結果を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="63b45-109">Your custom function can also return multiple results like this.</span></span>

![複数のセルに複数の結果を表示する SORT 関数のスクリーンショット。](../images/dynamic-array-spill.png)

<span data-ttu-id="63b45-111">動的配列数式であるカスタム関数を作成するには、値の2次元配列を返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="63b45-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="63b45-112">結果が、既に値を持つ隣接するセルにスピルされる場合、数式は #SPILL を表示します **。**</span><span class="sxs-lookup"><span data-stu-id="63b45-112">If the results spill into neighboring cells that already have values, the formula will display a **#SPILL!**</span></span> <span data-ttu-id="63b45-113">エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="63b45-113">error.</span></span> 

<span data-ttu-id="63b45-114">次の例は、分解した動的配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63b45-114">The following example shows how to return a dynamic array that spills down.</span></span>

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

<span data-ttu-id="63b45-115">次の例は、右に液体をこぼれた動的配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63b45-115">The following example shows how to return a dynamic array that spills right.</span></span> 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

<span data-ttu-id="63b45-116">次の例は、右下の配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="63b45-116">The following example shows how to return a dynamic array that spills both down and right.</span></span>

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a><span data-ttu-id="63b45-117">関連項目</span><span class="sxs-lookup"><span data-stu-id="63b45-117">See also</span></span>

- [<span data-ttu-id="63b45-118">動的配列とこぼれた配列の動作</span><span class="sxs-lookup"><span data-stu-id="63b45-118">Dynamic arrays and spilled array behavior</span></span>](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="63b45-119">Excel カスタム関数のオプション</span><span class="sxs-lookup"><span data-stu-id="63b45-119">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)