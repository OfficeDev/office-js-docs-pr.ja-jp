---
title: Excel アドインで複数の範囲を同時に操作する
description: Excel JavaScript ライブラリを使用して、複数の範囲に対して操作を実行したり、プロパティを設定したりする方法について説明します。
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: 6a508d8481d9851c7f7ae98ec959fcec9663972c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609770"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a><span data-ttu-id="dc06f-103">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="dc06f-103">Work with multiple ranges simultaneously in Excel add-ins</span></span>

<span data-ttu-id="dc06f-104">Excel JavaScript ライブラリを使用すると、同時に複数の範囲に対してアドインによる操作の実行とプロパティの設定が可能になります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-104">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="dc06f-105">範囲は連続している必要はありません。</span><span class="sxs-lookup"><span data-stu-id="dc06f-105">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="dc06f-106">コードがよりシンプルになることに加え、この方法でプロパティを設定すれば、各範囲に同じプロパティを個別に設定する方法よりも処理速度が格段に速くなります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-106">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

## <a name="rangeareas"></a><span data-ttu-id="dc06f-107">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="dc06f-107">RangeAreas</span></span>

<span data-ttu-id="dc06f-108">範囲の集合 (連続している可能性もあります) は、 [Rangeareas](/javascript/api/excel/excel.rangeareas)オブジェクトによって表されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-108">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="dc06f-109">`Range` 型と同様のプロパティとメソッドを持ちますが (多くの場合は同じまたは類似した名前)、以下に対しては調整が行われています。</span><span class="sxs-lookup"><span data-stu-id="dc06f-109">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="dc06f-110">プロパティのデータ型と、セッターとゲッターの動作。</span><span class="sxs-lookup"><span data-stu-id="dc06f-110">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="dc06f-111">メソッド パラメーターのデータ型と、メソッドの動作。</span><span class="sxs-lookup"><span data-stu-id="dc06f-111">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="dc06f-112">メソッドの戻り値のデータ型。</span><span class="sxs-lookup"><span data-stu-id="dc06f-112">The data types of method return values.</span></span>

<span data-ttu-id="dc06f-113">次にいくつか例を示します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-113">Some examples:</span></span>

- <span data-ttu-id="dc06f-114">`RangeAreas` には `address` プロパティがあり、`Range.address` プロパティのように 1 つのアドレスを返すのではなく、複数の範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-114">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="dc06f-115">`RangeAreas` には、一貫性がある場合、`RangeAreas` に指定された全範囲のデータ検証を表す `DataValidation` オブジェクトを返す `dataValidation` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-115">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="dc06f-116">`RangeAreas` に指定された全範囲に同じ `DataValidation` オブジェクトが適用されていない場合、このプロパティは `null` となります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-116">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="dc06f-117">これは、`RangeAreas` オブジェクトに関する、汎用的ではありませんが一般的な原則です: *`RangeAreas` に指定された全範囲のプロパティの値に一貫性がない場合、`null` となります*。</span><span class="sxs-lookup"><span data-stu-id="dc06f-117">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="dc06f-118">より詳しい情報といくつかの例外については、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dc06f-118">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="dc06f-119">`RangeAreas.cellCount` は、`RangeAreas` に指定された全範囲の合計セル数を取得します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-119">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="dc06f-120">`RangeAreas.calculate` は、`RangeAreas` に指定された全範囲のセルを再計算します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-120">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="dc06f-121">`RangeAreas.getEntireColumn` と `RangeAreas.getEntireRow` は、`RangeAreas` に指定された全範囲のセルの列 (または行) すべてを表す、別の `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-121">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="dc06f-122">たとえば、`RangeAreas` が "A1:C4" と "F14:L15" を表す場合、`RangeAreas.getEntireColumn` は "A:C" と "F:L" を表す `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-122">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="dc06f-123">`RangeAreas.copyFrom` は、コピー操作のコピー元範囲を表す `Range` または `RangeAreas` パラメーターのいずれかを取得できます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-123">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="dc06f-124">RangeAreas でも利用可能な Range メンバーの全リスト</span><span class="sxs-lookup"><span data-stu-id="dc06f-124">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="dc06f-125">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dc06f-125">Properties</span></span>

<span data-ttu-id="dc06f-126">リストにあるプロパティを読み取るコードを書く前に、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」の内容を理解しておいてください。</span><span class="sxs-lookup"><span data-stu-id="dc06f-126">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="dc06f-127">繰り返される内容について細かい注意点があります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-127">There are subtleties to what gets returned.</span></span>

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a><span data-ttu-id="dc06f-128">Methods</span><span class="sxs-lookup"><span data-stu-id="dc06f-128">Methods</span></span>

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- <span data-ttu-id="dc06f-129">`getOffsetRange()`( `getOffsetRangeAreas` オブジェクトでの名前 `RangeAreas` )</span><span class="sxs-lookup"><span data-stu-id="dc06f-129">`getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)</span></span>
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- <span data-ttu-id="dc06f-130">`getUsedRange()`( `getUsedRangeAreas` オブジェクトでの名前 `RangeAreas` )</span><span class="sxs-lookup"><span data-stu-id="dc06f-130">`getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)</span></span>
- <span data-ttu-id="dc06f-131">`getUsedRangeOrNullObject()`( `getUsedRangeAreasOrNullObject` オブジェクトでの名前 `RangeAreas` )</span><span class="sxs-lookup"><span data-stu-id="dc06f-131">`getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)</span></span>
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="dc06f-132">RangeArea 固有のプロパティとメソッド</span><span class="sxs-lookup"><span data-stu-id="dc06f-132">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="dc06f-133">`RangeAreas` 型には、`Range` オブジェクトには存在しないプロパティとメソッドがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-133">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="dc06f-134">次にいくつか選択したものを示します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-134">The following is a selection of them:</span></span>

- <span data-ttu-id="dc06f-135">`areas`: `RangeAreas` オブジェクトが表す全範囲を含む `RangeCollection` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="dc06f-135">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="dc06f-136">`RangeCollection` オブジェクトも新しいオブジェクトであり、他の Excel コレクション オブジェクトと類似しています。</span><span class="sxs-lookup"><span data-stu-id="dc06f-136">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="dc06f-137">これには、範囲を表す `Range` オブジェクトの配列である `items` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-137">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="dc06f-138">`areaCount`: `RangeAreas` で指定された範囲の合計数。</span><span class="sxs-lookup"><span data-stu-id="dc06f-138">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="dc06f-139">`getOffsetRangeAreas`: [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) と同じように動作します。ただし、`RangeAreas` を返し、元の `RangeAreas` で指定された範囲の 1 つからの各オフセットである範囲を含みます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-139">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="dc06f-140">RangeAreas の作成</span><span class="sxs-lookup"><span data-stu-id="dc06f-140">Create RangeAreas</span></span>

<span data-ttu-id="dc06f-141">`RangeAreas` オブジェクトの作成には、2 つの基本的な方法があります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-141">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="dc06f-142">`Worksheet.getRanges()` を呼び出して、範囲のアドレスがコンマで区切られた文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-142">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="dc06f-143">含める対象の範囲が既に [NamedItem](/javascript/api/excel/excel.nameditem) に指定されている場合、文字列にはアドレスではなくその名前を指定することができます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-143">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="dc06f-144">`Workbook.getSelectedRanges()` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-144">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="dc06f-145">このメソッドは、現在アクティブなワークシート上で選択されている全範囲を表す `RangeAreas` を返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-145">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="dc06f-146">一度 `RangeAreas` オブジェクトを作成すると、`getOffsetRangeAreas` や `getIntersection` など、`RangeAreas` を返すオブジェクト上のメソッドを使用して別のオブジェクトを作成できます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-146">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="dc06f-147">`RangeAreas` オブジェクトに新たな範囲を直接追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="dc06f-147">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="dc06f-148">たとえば、`RangeAreas.areas` 内のコレクションには `add` メソッドが存在しません。</span><span class="sxs-lookup"><span data-stu-id="dc06f-148">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="dc06f-149">`RangeAreas.areas.items` 配列のメンバーの追加または削除を直接試行してはいけません。</span><span class="sxs-lookup"><span data-stu-id="dc06f-149">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="dc06f-150">これにより、後でコード内で望ましくない動作が発生します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-150">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="dc06f-151">たとえば、追加の `Range` オブジェクトを配列にプッシュすることは可能ですが、エラーが発生します。`RangeAreas` のプロパティとメソッドは、その新しいアイテムがその場所に存在していないかのように動作するためです。</span><span class="sxs-lookup"><span data-stu-id="dc06f-151">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="dc06f-152">たとえば、`areaCount` プロパティにはこの方法でプッシュされた範囲は含まれません。また、`RangeAreas.getItemAt(index)` は、`index` が `areasCount-1`より大きい場合、エラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="dc06f-152">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="dc06f-153">同様に、`RangeAreas.areas.items` 配列内の `Range` オブジェクトを、参照を取得してその `Range.delete` メソッドを呼び出すという方法で削除すると、バグとなります。`Range` オブジェクトは*削除されます*が、親 `RangeAreas` オブジェクトのプロパティとメソッドは、そのオブジェクトがまだ存在するものとして動作するためです。</span><span class="sxs-lookup"><span data-stu-id="dc06f-153">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="dc06f-154">たとえば、コードで `RangeAreas.calculate` を呼び出すと、Office は範囲を計算しようとしますが、範囲オブジェクトが既に存在しないためにエラーとなります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-154">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="dc06f-155">複数の範囲でのプロパティの設定</span><span class="sxs-lookup"><span data-stu-id="dc06f-155">Set properties on multiple ranges</span></span>

<span data-ttu-id="dc06f-156">`RangeAreas` オブジェクトでプロパティを設定すると、`RangeAreas.areas` コレクション内の全範囲の対応するプロパティが設定されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-156">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="dc06f-157">次に、複数の範囲にプロパティを設定する例を示します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-157">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="dc06f-158">この関数は、**F3:F5** と **H3:H5** の範囲を強調表示します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-158">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="dc06f-159">この例は、`getRanges` に渡す範囲のアドレスをハード コーディングできる場合や実行時に簡単に計算できる場合に適用されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-159">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="dc06f-160">たとえば、これが適切なのは次のような場合です。</span><span class="sxs-lookup"><span data-stu-id="dc06f-160">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="dc06f-161">コードが、既知のテンプレートのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="dc06f-161">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="dc06f-162">コードが、データのスキーマが既知であるインポート済みデータのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="dc06f-162">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="dc06f-163">複数の範囲からの特定のセルの取得</span><span class="sxs-lookup"><span data-stu-id="dc06f-163">Get special cells from multiple ranges</span></span>

<span data-ttu-id="dc06f-164">`RangeAreas` オブジェクトの `getSpecialCells` メソッドと `getSpecialCellsOrNullObject` メソッドは、`Range` オブジェクトの同じ名前のメソッドと同じように機能します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-164">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="dc06f-165">これらのメソッドでは、`RangeAreas.areas` コレクション内のすべての範囲から、指定された特性を持つセルが返されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-165">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="dc06f-166">特殊なセルの詳細については、「[範囲内の特殊なセルの検索](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="dc06f-166">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) section for more details on special cells.</span></span>

<span data-ttu-id="dc06f-167">`RangeAreas` オブジェクトで `getSpecialCells` メソッドまたは `getSpecialCellsOrNullObject` メソッドを呼び出す場合:</span><span class="sxs-lookup"><span data-stu-id="dc06f-167">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="dc06f-168">最初のパラメーターとして `Excel.SpecialCellType.sameConditionalFormat` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じ条件付き書式を持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-168">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="dc06f-169">最初のパラメーターとして `Excel.SpecialCellType.sameDataValidation` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-169">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="dc06f-170">RangeAreas のプロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="dc06f-170">Read properties of RangeAreas</span></span>

<span data-ttu-id="dc06f-171">`RangeAreas` のプロパティ値の読み取りには、注意が必要です。`RangeAreas`内の範囲それぞれで、プロパティの値が異なる可能性があるためです。</span><span class="sxs-lookup"><span data-stu-id="dc06f-171">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="dc06f-172">一貫性のある値を返すことが*できる*場合には返す、というのが一般的なルールです。</span><span class="sxs-lookup"><span data-stu-id="dc06f-172">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="dc06f-173">たとえば、次のコードでは、ピンクの RGB コード (`#FFC0CB`) と `true` がコンソールに記録されます。`RangeAreas`オブジェクト内の範囲のどちらも、塗りつぶし色がピンクであり、列全体であるためです。</span><span class="sxs-lookup"><span data-stu-id="dc06f-173">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

<span data-ttu-id="dc06f-174">一貫性を期待できない場合、事態は複雑となります。</span><span class="sxs-lookup"><span data-stu-id="dc06f-174">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="dc06f-175">`RangeAreas` プロパティの動作は、次の 3 つの原則に従います。</span><span class="sxs-lookup"><span data-stu-id="dc06f-175">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="dc06f-176">`RangeAreas` オブジェクトのブール値プロパティは、すべてのメンバー範囲でプロパティが true でない限り、`false` を返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-176">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="dc06f-177">ブール値以外のプロパティ (`address` プロパティを除く) は、すべてのメンバー範囲で対応するプロパティが同じ値ではない限り、`null` を返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-177">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="dc06f-178">`address` プロパティは、メンバー範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-178">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="dc06f-179">たとえば、次のコードでは、1 つの範囲のみが列全体であり、1 つの範囲のみがピンクで塗りつぶされている `RangeAreas` を作成します。</span><span class="sxs-lookup"><span data-stu-id="dc06f-179">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="dc06f-180">コンソールには、塗りつぶし色の場合は `null`、`isEntireRow` プロパティの場合は `false`、`address` プロパティの場合は "Sheet1!F3:F5, Sheet1!H:H" ("Sheet1" はシート名) が表示されます。</span><span class="sxs-lookup"><span data-stu-id="dc06f-180">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a><span data-ttu-id="dc06f-181">関連項目</span><span class="sxs-lookup"><span data-stu-id="dc06f-181">See also</span></span>

- [<span data-ttu-id="dc06f-182">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="dc06f-182">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="dc06f-183">Excel JavaScript API を使用して範囲を操作する (基本)</span><span class="sxs-lookup"><span data-stu-id="dc06f-183">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="dc06f-184">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="dc06f-184">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
