---
title: Excel アドインの空白値と null 値
description: Excel オブジェクトモデルのメソッドとプロパティで空白の null 値を操作する方法について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409408"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a><span data-ttu-id="c7c1d-103">Excel アドインの空白値と null 値</span><span class="sxs-lookup"><span data-stu-id="c7c1d-103">Blank and null values in Excel add-ins</span></span>

<span data-ttu-id="c7c1d-104">`null` と空の文字列は、Excel JavaScript API では特別な意味を持ちます。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-104">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="c7c1d-105">これらは、空のセル、書式設定なし、既定値を表すために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-105">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="c7c1d-106">このセクションでは、プロパティの取得や設定を行うときに `null` や空の文字列を使用する方法について詳しく説明します。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-106">This section details the use of `null` and empty string when getting and setting properties.</span></span>

## <a name="null-input-in-2-d-array"></a><span data-ttu-id="c7c1d-107">2 次元配列での null の入力</span><span class="sxs-lookup"><span data-stu-id="c7c1d-107">null input in 2-D Array</span></span>

<span data-ttu-id="c7c1d-p102">Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。 範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p102">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="c7c1d-p103">たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。 次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p103">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a><span data-ttu-id="c7c1d-112">プロパティに対する null の入力</span><span class="sxs-lookup"><span data-stu-id="c7c1d-112">null input for a property</span></span>

<span data-ttu-id="c7c1d-p104">`null` は単一プロパティに有効な入力ではありません。たとえば、次のコード スニペットは、範囲の `values` プロパティを `null` に設定できないため無効です。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p104">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null; // This is not a valid snippet. 
```

<span data-ttu-id="c7c1d-115">同様に、次のコード スニペットは、`null` が `color` プロパティで有効な値ではないため無効です。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-115">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a><span data-ttu-id="c7c1d-116">応答内の null プロパティ値</span><span class="sxs-lookup"><span data-stu-id="c7c1d-116">null property values in the response</span></span>

<span data-ttu-id="c7c1d-p105">指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。 たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p105">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="c7c1d-119">範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-119">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="c7c1d-120">範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-120">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

## <a name="blank-input-for-a-property"></a><span data-ttu-id="c7c1d-121">プロパティに対する空白の入力</span><span class="sxs-lookup"><span data-stu-id="c7c1d-121">Blank input for a property</span></span>

<span data-ttu-id="c7c1d-p106">プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p106">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="c7c1d-124">範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-124">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="c7c1d-125">`numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-125">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="c7c1d-126">`formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-126">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

## <a name="blank-property-values-in-the-response"></a><span data-ttu-id="c7c1d-127">応答内の空白のプロパティ値</span><span class="sxs-lookup"><span data-stu-id="c7c1d-127">Blank property values in the response</span></span>

<span data-ttu-id="c7c1d-p107">読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。 次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。 2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。</span><span class="sxs-lookup"><span data-stu-id="c7c1d-p107">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
