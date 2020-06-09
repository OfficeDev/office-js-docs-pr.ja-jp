---
ms.date: 05/11/2020
description: Office Excel アドインで、カスタム関数から複数の結果を返します。
title: カスタム関数から複数の結果を返す
localization_priority: Normal
ms.openlocfilehash: e25965277fbbe1c39007f79f401bf62b25760488
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609651"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="e4772-103">カスタム関数から複数の結果を返す</span><span class="sxs-lookup"><span data-stu-id="e4772-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="e4772-104">隣接するセルに返される、カスタム関数から複数の結果を返すことができます。</span><span class="sxs-lookup"><span data-stu-id="e4772-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="e4772-105">この動作は spilling と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="e4772-105">This behavior is called spilling.</span></span> <span data-ttu-id="e4772-106">カスタム関数が結果の配列を返す場合は、動的配列数式と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="e4772-106">When your custom function returns an array of results, it's known as a dynamic array formula.</span></span> <span data-ttu-id="e4772-107">Excel の動的配列数式の詳細については、「動的配列」[および「こぼれた配列の動作](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4772-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="e4772-108">次の図は、関数が隣接するセルにどのように分解されるかを示して `SORT` います。</span><span class="sxs-lookup"><span data-stu-id="e4772-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="e4772-109">カスタム関数は、次のような複数の結果を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="e4772-109">Your custom function can also return multiple results like this.</span></span>

![複数のセルに複数の結果を表示する ' SORT ' 関数のスクリーンショット。](../images/dynamic-array-spill.png)

<span data-ttu-id="e4772-111">動的配列数式であるカスタム関数を作成するには、値の2次元配列を返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e4772-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="e4772-112">結果が、既に値を持つ隣接するセルにスピルされる場合、数式はエラーを表示し `#SPILL!` ます。</span><span class="sxs-lookup"><span data-stu-id="e4772-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="e4772-113">次の例は、分解した動的配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e4772-113">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="e4772-114">次の例は、右に液体をこぼれた動的配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e4772-114">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="e4772-115">次の例は、右下の配列を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e4772-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e4772-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="e4772-116">See also</span></span>

- [<span data-ttu-id="e4772-117">動的配列とこぼれた配列の動作</span><span class="sxs-lookup"><span data-stu-id="e4772-117">Dynamic arrays and spilled array behavior</span></span>](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="e4772-118">Excel カスタム関数のオプション</span><span class="sxs-lookup"><span data-stu-id="e4772-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)