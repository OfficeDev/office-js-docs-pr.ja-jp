---
title: Excel のアドインで同時に複数の範囲の操作をします。
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506295"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="5442f-102">Excel のアドイン (プレビュー) で同時に複数のセル範囲を操作します。</span><span class="sxs-lookup"><span data-stu-id="5442f-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="5442f-p101">Excel の JavaScript ライブラリは、操作を実行し、同時に複数の範囲のプロパティを設定するように追加できます。範囲が隣接している必要はありません。コードを簡単にするだけでなく、このプロパティを設定する方法は、範囲ごとに個別に同じプロパティを設定するよりもはるかに高速実行されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p101">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="5442f-p102">この資料に記載されている Apiは、 **Office 2016 クイック実行バージョン ビルト 1809 10820.20000 の** 以降を日梅雨とします。(適切なビルドを取得する [Office 内部からのプログラム](https://products.office.com/office-insider) に参加する必要があります)。また、 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)からベータ版の Office の JavaScript ライブラリを読み込む必要があります。最後に、これらの Api の参照ページまだありません。次の種類の定義ファイルにはそれらについての説明: [ベータ版の office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)です。</span><span class="sxs-lookup"><span data-stu-id="5442f-p102">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="5442f-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="5442f-110">RangeAreas</span></span>

<span data-ttu-id="5442f-p103">(場合によっては連続していない) の範囲のセットは、 `Excel.RangeAreas` オブジェクトで表されます。`Range`に似ているプロパティとメソッドを持ちますが、 (多くの場合、同じまたは同様の名前を持つ)、以下のような修正が加えられました。</span><span class="sxs-lookup"><span data-stu-id="5442f-p103">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="5442f-113">プロパティのデータ型とセッターとゲッターの動作です。</span><span class="sxs-lookup"><span data-stu-id="5442f-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="5442f-114">メソッドパラメーターのデータ型と、メソッドの動作です。</span><span class="sxs-lookup"><span data-stu-id="5442f-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="5442f-115">メソッドのデータ型は、値を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-115">The data types of method return values.</span></span>

<span data-ttu-id="5442f-116">例:</span><span class="sxs-lookup"><span data-stu-id="5442f-116">Some examples:</span></span>

- <span data-ttu-id="5442f-117">`RangeAreas` `Range.address` プロパティとともに1 つのアドレスとしてではなく、範囲のアドレスのコンマ区切りの文字列を返す`address` プロパティを持ちます。</span><span class="sxs-lookup"><span data-stu-id="5442f-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="5442f-p104">`RangeAreas` 一貫性のある場合に`RangeAreas`内のすべての範囲のデータの有効性を表す`DataValidation`オブジェクトを返す `dataValidation` プロパティを持ちます。`RangeAreas`内のすべての範囲に適用されない  `DataValidation`と同一の場合、  プロパティは、オブジェクトは`null` です。これは、全般的でかつ汎用的でない `RangeAreas` オブジェクト: *プロパティが、`RangeAreas`すべての範囲の値に一貫性を持っていない場合、  `null`です* 。いくつかの例外の詳細については、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-p104">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="5442f-122">`RangeAreas.cellCount` `RangeAreas` 内のすべての範囲内のセルの合計数を取得します。</span><span class="sxs-lookup"><span data-stu-id="5442f-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="5442f-123">`RangeAreas.calculate` `RangeAreas` 内のすべての範囲のセルを再計算します。</span><span class="sxs-lookup"><span data-stu-id="5442f-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="5442f-p105">`RangeAreas.getEntireColumn` また `RangeAreas.getEntireRow` は、 `RangeAreas` 内のすべての範囲のすべての列 (または行) を表す `RangeAreas`オブジェクトを返します。例えば、 `RangeAreas` が、「a1: c4」および「F14:L15」を表す場合、 `RangeAreas.getEntireColumn` は、 "A:C"と"F:L"を表すオブジェクト`RangeAreas` を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p105">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="5442f-126">`RangeAreas.copyFrom` コピー操作のソース範囲を表すパラメーターである`Range` または `RangeAreas` のいずれをとることができます。</span><span class="sxs-lookup"><span data-stu-id="5442f-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="5442f-127">RangeAreas でも利用できる範囲のメンバーの完全なリスト</span><span class="sxs-lookup"><span data-stu-id="5442f-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="5442f-128">プロパティ</span><span class="sxs-lookup"><span data-stu-id="5442f-128">Properties</span></span>

<span data-ttu-id="5442f-p106">任意のプロパティを読み取るコードの一覧を記述する前に、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) に精通します。何が返されるかは微妙です。</span><span class="sxs-lookup"><span data-stu-id="5442f-p106">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="5442f-131">アドレス</span><span class="sxs-lookup"><span data-stu-id="5442f-131">address</span></span>
- <span data-ttu-id="5442f-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="5442f-132">addressLocal</span></span>
- <span data-ttu-id="5442f-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="5442f-133">cellCount</span></span>
- <span data-ttu-id="5442f-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="5442f-134">conditionalFormats</span></span>
- <span data-ttu-id="5442f-135">コンテキスト</span><span class="sxs-lookup"><span data-stu-id="5442f-135">context</span></span>
- <span data-ttu-id="5442f-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="5442f-136">dataValidation</span></span>
- <span data-ttu-id="5442f-137">format</span><span class="sxs-lookup"><span data-stu-id="5442f-137">format</span></span>
- <span data-ttu-id="5442f-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="5442f-138">isEntireColumn</span></span>
- <span data-ttu-id="5442f-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="5442f-139">isEntireRow</span></span>
- <span data-ttu-id="5442f-140">style</span><span class="sxs-lookup"><span data-stu-id="5442f-140">style</span></span>
- <span data-ttu-id="5442f-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="5442f-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="5442f-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="5442f-142">Methods</span></span>

<span data-ttu-id="5442f-143">プレビューの範囲メソッドを示します。</span><span class="sxs-lookup"><span data-stu-id="5442f-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="5442f-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="5442f-144">calculate()</span></span>
- <span data-ttu-id="5442f-145">clear()</span><span class="sxs-lookup"><span data-stu-id="5442f-145">clear()</span></span>
- <span data-ttu-id="5442f-146">convertDataTypeToText() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="5442f-147">convertToLinkedDataType() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="5442f-148">copyFrom() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="5442f-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="5442f-149">getEntireColumn()</span></span>
- <span data-ttu-id="5442f-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="5442f-150">getEntireRow()</span></span>
- <span data-ttu-id="5442f-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="5442f-151">getIntersection()</span></span>
- <span data-ttu-id="5442f-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="5442f-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="5442f-153">getOffsetRange() (RangeAreas オブジェクトの named getOffsetRangeAreas を名前付け)</span><span class="sxs-lookup"><span data-stu-id="5442f-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="5442f-154">getSpecialCells() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="5442f-155">getSpecialCellsOrNullObject() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="5442f-156">getTables() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-156">getTables() (preview)</span></span>
- <span data-ttu-id="5442f-157">getUsedRange() (RangeAreas オブジェクトの getUsedRangeAreas を名前付け)</span><span class="sxs-lookup"><span data-stu-id="5442f-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="5442f-158">getUsedRangeOrNullObject() (RangeAreas オブジェクトでは getUsedRangeAreasOrNullObject という名前)</span><span class="sxs-lookup"><span data-stu-id="5442f-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="5442f-159">load()</span><span class="sxs-lookup"><span data-stu-id="5442f-159">load()</span></span>
- <span data-ttu-id="5442f-160">set()</span><span class="sxs-lookup"><span data-stu-id="5442f-160">set\*</span></span>
- <span data-ttu-id="5442f-161">setDirty() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5442f-161">setDirty() (preview)</span></span>
- <span data-ttu-id="5442f-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="5442f-162">toJSON()</span></span>
- <span data-ttu-id="5442f-163">track()</span><span class="sxs-lookup"><span data-stu-id="5442f-163">track</span></span>
- <span data-ttu-id="5442f-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="5442f-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="5442f-165">RangeArea に固有のプロパティおよびメソッド</span><span class="sxs-lookup"><span data-stu-id="5442f-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="5442f-p107">`RangeAreas` 型には、 `Range` オブジェクト上にないいくつかのプロパティとメソッドがあります。それらのいくつかを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p107">The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them:</span></span>

- <span data-ttu-id="5442f-p108">`areas`  `RangeAreas` で表される範囲のすべてを含むオブジェクトの  `RangeCollection` オブジェクトです。 `RangeCollection` オブジェクトは新しく、他の Excel のコレクション オブジェクトに似ています。 プロパティの範囲を表す `Range` オブジェクトの配列である`items` プロパティを持ちます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p108">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="5442f-171">`areaCount`: `RangeAreas` 内の合計数。</span><span class="sxs-lookup"><span data-stu-id="5442f-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="5442f-172">`getOffsetRangeAreas`:[   が返され、元の](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) `RangeAreas`   の範囲の一つからそれぞれからのオフセットの範囲を含む点を除き、 Range.getOffsetRange`RangeAreas` と同じように動作します。</span><span class="sxs-lookup"><span data-stu-id="5442f-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="5442f-173">RangeAreas の作成と、プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="5442f-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="5442f-174">`RangeAreas`オブジェクトを 2 つの基本的な方法で作成することができます。</span><span class="sxs-lookup"><span data-stu-id="5442f-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="5442f-p109">`Worksheet.getRanges()` を呼び出し、コンマで区切られた範囲アドレスを含む文字列を渡します。含めたい任意の範囲を [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)  とした場合、文字列にアドレスではなく、名前を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p109">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="5442f-p110">`Workbook.getSelectedRanges()`を呼出します。このメソッドは、現在アクティブなワークシートで選択されているすべての範囲を表す`RangeAreas`を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p110">Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="5442f-179">`RangeAreas` オブジェクトを作成したら、`getOffsetRangeAreas` および `getIntersection` のような `RangeAreas` を返すオブジェクト上のメソッドを使用して他のユーザーを作成することができます。</span><span class="sxs-lookup"><span data-stu-id="5442f-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="5442f-p111">`RangeAreas`オブジェクトに、追加の範囲を直接追加することはできません。例えば、`RangeAreas.areas`内のコレクションは、`add`メソッドを持ちません。</span><span class="sxs-lookup"><span data-stu-id="5442f-p111">You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="5442f-p112">`RangeAreas.areas.items`  の配列のメンバーを直接追加または削除しないようにしてください。コード内で望ましくない動作をしてしまいます。例えば、配列に追加の `Range`  オブジェクトをプッシュすることは可能ですが、これにより `RangeAreas`  プロパティやメソッドは、新しいアイテムがないかのように動作するためエラーが発生します。例えば、`areaCount` プロパティには、このような方法でプッシュされた範囲を含んでおらず、`RangeAreas.getItemAt(index)` が `index`より大きい場合に、`areasCount-1` がエラーをスローします。同様に、参照を取得し、その メソッドを呼び出して、`Range` 内の <オブ`RangeAreas.areas.items`  `Range.delete`ジェクトを削除すると、バグが発生します。 オブジェクト`Range`は\* 、\* 削除されますが、親の`RangeAreas`オブジェクトのプロパティとメソッドは、それがまだ存在するかのように動作するか、動作しようとします。例えば、`RangeAreas.calculate` を呼出した場合、Office は範囲を計算しようとしますが、範囲オブジェクトがないためエラーが生じます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p112">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="5442f-188">`RangeAreas`上のプロパティの設定は、`RangeAreas.areas`コレクション上のすべての範囲に対応するプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="5442f-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="5442f-p113">次は、複数の範囲のプロパティの設定の例です。関数には、**F3:F5** と **H3:H5** の範囲が強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p113">The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="5442f-p114">この例では、`getRanges`に渡す範囲のアドレスを渡すハード コードできるか 、簡単に実行時に自動的に計算できるシナリオを適用します。これが正しいであろうシナリオの一部は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="5442f-p114">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="5442f-193">コードは、既知のテンプレートのコンテキストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="5442f-194">コードは、データのスキーマがわかっているインポートされたデータのコンテキストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="5442f-p115">コーディング時にどのような範囲で実行すれはよいかわからない場合には、ランタイムで検出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5442f-p115">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="5442f-197">範囲の領域をプログラムで検出します。</span><span class="sxs-lookup"><span data-stu-id="5442f-197">Discover range areas programmatically</span></span>

<span data-ttu-id="5442f-p116">`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、実行時に、セルの特性とセルの値の型を基に操作し範囲を検索できます。TypeScript データタイプのファイルからのメソッドのシグネチャを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p116">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="5442f-p117">次は、最初のシグネチャを使用する場合の例です。このコードに関して以下に注意してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-p117">The following is an example of using the first one. About this code, note:</span></span>

- <span data-ttu-id="5442f-202">その範囲のみで最初に`Worksheet.getUsedRange`を呼び出し、`getSpecialCells`を呼出して検索する必要があるシートの一部を制限します。</span><span class="sxs-lookup"><span data-stu-id="5442f-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="5442f-p118">`Excel.SpecialCellType`列挙型からの値の文字列バージョンをパラメータとして`getSpecialCells`に渡します。代わりに渡される他の値のいくつかは、空白のセルには「空白」、数式のかわりにリテラル値を持つセルには「定数」、`usedRange`の最初のセルと同じ条件付き書式が設定されるセルには「SameConditionalFormat」です。最初のセルは、上の左端のセルです。列挙型の値の一覧は、[ベータ版の「office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-p118">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="5442f-207">`getSpecialCells`メソッドは、`RangeAreas`オブジェクトを返します。数式を入力した全てのセルは、すべて連続していない場合も、ピンク色に色分け表示されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="5442f-p119">対象となる特性を持つ*任意*のセル範囲はありません。`getSpecialCells`が何も検出しない場合は、**ItemNotFound** エラーがスローされます。これがある場合は、`catch`ブロック/メソッドへのコントロールのフローを逸します。ない場合は、エラーは、機能を停止します。エラーをスローすることが、対象となる特性を持つセルが存在しない場合に、正確に必要されることであるシナリオがある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="5442f-p119">Sometimes the range doesn't have *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="5442f-p120">通常の動作では、シナリオでは、一致するセルがない場合もありますが、これはおそらく一般的ではありません。コードは、この可能性を確認し、エラーをスローすることがなく適切に処理すること必要があります。これらのシナリオに関しては、`getSpecialCellsOrNullObject`メソッドを使用して`RangeAreas.isNullObject`プロパティをテストします。次に、例を示します。このコードに関して以下に注意してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-p120">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property. The following is an example. Note about this code:</span></span>

- <span data-ttu-id="5442f-p121">`getSpecialCellsOrNullObject`メソッドは、常にプロキシ オブジェクトを返します。したがって、通常 JavaScript という意味では、決して`null`にはなりません。一致するセルが見つからない場合は、オブジェクトの`isNullObject`プロパティが、`true`に設定されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p121">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="5442f-p122">`isNullObject`プロパティをテストする*前に* 、`context.sync`を呼び出す。読み込むために常にプロパティを読み込み同期させる必要があるため、これはすべての`*OrNullObject`メソッドとプロパティの必要条件です。ただし、`isNullObject`プロパティを*明示的に*ロードする必要はありません。`load`がオブジェクトで呼び出されない場合でも、`context.sync`が自動的にロードします。詳細情報に関しては、[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) を参照してください 。</span><span class="sxs-lookup"><span data-stu-id="5442f-p122">It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="5442f-p123">最初に数式のセルを含まない範囲を選択して実行することで、このコードをテストできます。次に、少なくとも 1 つのセルに数式がある範囲を選択し、それをもう一度実行します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p123">You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="5442f-226">わかりやすくするために、この記事の他のすべての例では、`getSpecialCellsOrNullObject` ではなく`getSpecialCells` メソッドを使用しています。</span><span class="sxs-lookup"><span data-stu-id="5442f-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="5442f-227">セルの値の型で対象セルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="5442f-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="5442f-p124">列挙型`Excel.SpecialCellValueType`の省略可能な 2 番目のパラメータがあり、対象に対してさらにセルを狭めるます。「数式」または「定数」のいずれかを`getSpecialCells`または`getSpecialCellsOrNullObject`に渡す場合にのみ使用できます。パラメータは、特定の種類の値のあるセルが必要であることを指定します。「エラー」、「論理」(ブール値を指す)、「番号」、および「テキスト」の 4 つの基本的な種類があります。(列挙型は、これら以外の他の値について次に説明する 4 つです。)次に、例を示します。このコードに関して次に注意してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-p124">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target. You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`. The parameter specifies that you only want cells with certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:</span></span>

- <span data-ttu-id="5442f-p125">リテラルの数値のあるセルのみが強調表示されます。数式(結果は数値の場合でも)またはブール値、文字列、またはエラーの状態のセルが強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p125">It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="5442f-236">コードをテストするには、リテラルの数値、他の種類のリテラル値、一部の数式のセルがワークシートにあることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="5442f-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="5442f-p126">場合によってすべてのテキスト値およびすべてのブール値 (「論理」) のセルのように 1 つ以上のセル値の種類を操作する必要があります。`Excel.SpecialCellValueType`列挙型はタイプを組み合わせることができる値を持っています。たとえば、「LogicalText」はブール値、テキスト値を持つすべてのセルをターゲットにします。4 つの基本的なタイプの 2 つまたは 3 つを組み合わせることができます。基本的な種類を組み合わせるこれらの列挙値の名前は、常にアルファベット順にします。したがって、エラー値、テキスト値、およびブール値を持つセルを組み合わせるには、「LogicalErrorText」または「TextErrorLogical」ではなく「ErrorLogicalText」を使用します。「すべて」の既定のパラメータが全 4 種類を組み合わせます。次の使用例では、数値またはブール値を生成する式を持つすべてのセルが強調表示されています。</span><span class="sxs-lookup"><span data-stu-id="5442f-p126">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". The default parameter of "All" combines all four types. The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="5442f-245">`Excel.SpecialCellValueType` パラメータは、`Excel.SpecialCellType` パラメータが 「式」または「定数」である場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="5442f-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="5442f-246">RangeAreas 内の RangeAreas を取得します。</span><span class="sxs-lookup"><span data-stu-id="5442f-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="5442f-p127">`RangeAreas`タイプ自体も `getSpecialCells` と `getSpecialCellsOrNullObject` を持ち、これらは 2 つの同じパラメータを受け取るメソッドです。これらのメソッドは、`RangeAreas.areas` コレクションのすべての範囲から対象となるセルを返します。`Range` オブジェクトではなく`RangeAreas` オブジェクトを呼び出した場合のメソッドの動作には 1 つの小さな違いがあります。「SameConditionalFormat」を最初のパラメータとして渡すと、メソッドは、*`RangeAreas.areas` コレクション内の最初の範囲の*上の左端のセルと同じ条件付き書式を持つすべてのセルを返します。同じポイントが、「SameDataValidation」にも当てはまります。`Range.getSpecialCells`に渡される場合、 *範囲内の*上の左端のセルと同じデータの入力規則をもつすべてのセルを返します。しかし、`RangeAreas.getSpecialCells` に渡される場合、*`RangeAreas.areas` コレクション内の最初の範囲の*上の左端のセルと同じデータの入力規則を持つすべてのセルを返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p127">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*. But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="5442f-252">RangeAreas のプロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="5442f-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="5442f-p128">`RangeAreas`のプロパティ値の読み取りには、`RangeAreas`内の指定したプロパティの別の範囲内の値が異なる可能性があるため、注意が必要です。一般的な規則では、一貫性のある値を返すことが*できる*場合は、返されます。例えば、次のコードでは、ピンク色の (`#FFC0CB`) と`true`用のRGBコードは、`RangeAreas`オブジェクトの両方の範囲がピンク色で塗りつぶされ、両方が全列であるため、コンソールに格納されます。</span><span class="sxs-lookup"><span data-stu-id="5442f-p128">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

<span data-ttu-id="5442f-p129">整合性が可能でない場合、より複雑になります。`RangeAreas`プロパティのビヘイビアーは、次の 3 つの原則に従います。</span><span class="sxs-lookup"><span data-stu-id="5442f-p129">Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="5442f-258">`RangeAreas`オブジェクトのブール値プロパティは、すべてのメンバーの範囲が真でない限り、`false`を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="5442f-259">非ブール値プロパティは、 `address` プロパティ例外を除いて、全てのメンバーの範囲に対応するプロパティが同じ値を持っていない限り、 `null` を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="5442f-260">`address`プロパティは、メンバーの範囲のアドレスのコンマ区切りの文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="5442f-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="5442f-p130">たとえば、次のコードは、1 つだけ列全体であり、 1 つだけがピンク色で塗りつぶされる`RangeAreas`を生成します。コンソールが、塗りつぶしの色に`null`を、`isEntireRow`プロパティに`false`を、および「Sheet1!F3:F5, Sheet1!H:H」(シート名は「Sheet1」と仮定) を`address`プロパティに表示します。</span><span class="sxs-lookup"><span data-stu-id="5442f-p130">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

## <a name="see-also"></a><span data-ttu-id="5442f-263">関連項目</span><span class="sxs-lookup"><span data-stu-id="5442f-263">See also</span></span>

- [<span data-ttu-id="5442f-264">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="5442f-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="5442f-265">Range Object (Excel 向け JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="5442f-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="5442f-p131">[RangeAreas Object (Excel 向け JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー中の場合、このリンクが動作しない 可能性があります。代わりに[ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="5442f-p131">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview. As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>