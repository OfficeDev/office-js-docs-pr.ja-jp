---
title: Excel アドインで複数の範囲を同時に操作する
description: ''
ms.date: 12/26/2018
ms.openlocfilehash: ab7cd9757adaedf2b6cc43fdcc604b98a60b6ecd
ms.sourcegitcommit: 8d248cd890dae1e9e8ef1bd47e09db4c1cf69593
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/27/2018
ms.locfileid: "27447233"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="edc05-102">Excel アドインで複数の範囲を同時に操作する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="edc05-103">Excel JavaScript ライブラリを使用すると、同時に複数の範囲に対してアドインによる操作の実行とプロパティの設定が可能になります。</span><span class="sxs-lookup"><span data-stu-id="edc05-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="edc05-104">範囲は連続している必要はありません。</span><span class="sxs-lookup"><span data-stu-id="edc05-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="edc05-105">コードがよりシンプルになることに加え、この方法でプロパティを設定すれば、各範囲に同じプロパティを個別に設定する方法よりも処理速度が格段に速くなります。</span><span class="sxs-lookup"><span data-stu-id="edc05-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="edc05-106">この記事で説明する API には、**Office 2016 クイック実行バージョン 1809 Build 10820.20000** 以降が必要です </span><span class="sxs-lookup"><span data-stu-id="edc05-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="edc05-107">([Office Insider プログラム](https://products.office.com/office-insider)に参加して、適切なビルドを取得することが必要な場合があります)。また、Office JavaScript ライブラリのベータ版を [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) からロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="edc05-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="edc05-108">最後に、これらの API セットに関する参照ページはまだありません。</span><span class="sxs-lookup"><span data-stu-id="edc05-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="edc05-109">ただし、定義の種類ファイル [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) に説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="edc05-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="edc05-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="edc05-110">RangeAreas</span></span>

<span data-ttu-id="edc05-111">範囲のセット (連続している必要はなし) は、`Excel.RangeAreas` オブジェクトで表されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="edc05-112">`Range` 型と同様のプロパティとメソッドを持ちますが (多くの場合は同じまたは類似した名前)、以下に対しては調整が行われています。</span><span class="sxs-lookup"><span data-stu-id="edc05-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="edc05-113">プロパティのデータ型と、セッターとゲッターの動作。</span><span class="sxs-lookup"><span data-stu-id="edc05-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="edc05-114">メソッド パラメーターのデータ型と、メソッドの動作。</span><span class="sxs-lookup"><span data-stu-id="edc05-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="edc05-115">メソッドの戻り値のデータ型。</span><span class="sxs-lookup"><span data-stu-id="edc05-115">The data types of method return values.</span></span>

<span data-ttu-id="edc05-116">次にいくつか例を示します。</span><span class="sxs-lookup"><span data-stu-id="edc05-116">Some examples:</span></span>

- <span data-ttu-id="edc05-117">`RangeAreas` には `address` プロパティがあり、`Range.address` プロパティのように 1 つのアドレスを返すのではなく、複数の範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="edc05-118">`RangeAreas` には、一貫性がある場合、`RangeAreas` に指定された全範囲のデータ検証を表す `DataValidation` オブジェクトを返す `dataValidation` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="edc05-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="edc05-119">`RangeAreas` に指定された全範囲に同じ `DataValidation` オブジェクトが適用されていない場合、このプロパティは `null` となります。</span><span class="sxs-lookup"><span data-stu-id="edc05-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="edc05-120">これは、`RangeAreas` オブジェクトに関する、汎用的ではありませんが一般的な原則です: *`RangeAreas` に指定された全範囲のプロパティの値に一貫性がない場合、`null` となります*。</span><span class="sxs-lookup"><span data-stu-id="edc05-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="edc05-121">より詳しい情報といくつかの例外については、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="edc05-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="edc05-122">`RangeAreas.cellCount` は、`RangeAreas` に指定された全範囲の合計セル数を取得します。</span><span class="sxs-lookup"><span data-stu-id="edc05-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="edc05-123">`RangeAreas.calculate` は、`RangeAreas` に指定された全範囲のセルを再計算します。</span><span class="sxs-lookup"><span data-stu-id="edc05-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="edc05-124">`RangeAreas.getEntireColumn` と `RangeAreas.getEntireRow` は、`RangeAreas` に指定された全範囲のセルの列 (または行) すべてを表す、別の `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="edc05-125">たとえば、`RangeAreas` が "A1:C4" と "F14:L15" を表す場合、`RangeAreas.getEntireColumn` は "A:C" と "F:L" を表す `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="edc05-126">`RangeAreas.copyFrom` は、コピー操作のコピー元範囲を表す `Range` または `RangeAreas` パラメーターのいずれかを取得できます。</span><span class="sxs-lookup"><span data-stu-id="edc05-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="edc05-127">RangeAreas でも利用可能な Range メンバーの全リスト</span><span class="sxs-lookup"><span data-stu-id="edc05-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="edc05-128">プロパティ</span><span class="sxs-lookup"><span data-stu-id="edc05-128">Properties</span></span>

<span data-ttu-id="edc05-129">リストにあるプロパティを読み取るコードを書く前に、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」の内容を理解しておいてください。</span><span class="sxs-lookup"><span data-stu-id="edc05-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="edc05-130">繰り返される内容について細かい注意点があります。</span><span class="sxs-lookup"><span data-stu-id="edc05-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="edc05-131">address</span><span class="sxs-lookup"><span data-stu-id="edc05-131">address</span></span>
- <span data-ttu-id="edc05-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="edc05-132">addressLocal</span></span>
- <span data-ttu-id="edc05-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="edc05-133">cellCount</span></span>
- <span data-ttu-id="edc05-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="edc05-134">conditionalFormats</span></span>
- <span data-ttu-id="edc05-135">context</span><span class="sxs-lookup"><span data-stu-id="edc05-135">context</span></span>
- <span data-ttu-id="edc05-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="edc05-136">dataValidation</span></span>
- <span data-ttu-id="edc05-137">format</span><span class="sxs-lookup"><span data-stu-id="edc05-137">format</span></span>
- <span data-ttu-id="edc05-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="edc05-138">isEntireColumn</span></span>
- <span data-ttu-id="edc05-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="edc05-139">isEntireRow</span></span>
- <span data-ttu-id="edc05-140">style</span><span class="sxs-lookup"><span data-stu-id="edc05-140">style</span></span>
- <span data-ttu-id="edc05-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="edc05-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="edc05-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="edc05-142">Methods</span></span>

<span data-ttu-id="edc05-143">プレビュー段階の Range メソッドについてはマークが付いています。</span><span class="sxs-lookup"><span data-stu-id="edc05-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="edc05-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="edc05-144">calculate()</span></span>
- <span data-ttu-id="edc05-145">clear()</span><span class="sxs-lookup"><span data-stu-id="edc05-145">clear()</span></span>
- <span data-ttu-id="edc05-146">convertDataTypeToText() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="edc05-147">convertToLinkedDataType() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="edc05-148">copyFrom() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="edc05-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="edc05-149">getEntireColumn()</span></span>
- <span data-ttu-id="edc05-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="edc05-150">getEntireRow()</span></span>
- <span data-ttu-id="edc05-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="edc05-151">getIntersection()</span></span>
- <span data-ttu-id="edc05-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="edc05-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="edc05-153">getOffsetRange() (RangeAreas オブジェクトでの名前は getOffsetRangeAreas)</span><span class="sxs-lookup"><span data-stu-id="edc05-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="edc05-154">getSpecialCells() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="edc05-155">getSpecialCellsOrNullObject() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="edc05-156">getTables() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-156">getTables() (preview)</span></span>
- <span data-ttu-id="edc05-157">getUsedRange() (RangeAreas オブジェクトでの名前は getUsedRangeAreas)</span><span class="sxs-lookup"><span data-stu-id="edc05-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="edc05-158">getUsedRangeOrNullObject() (RangeAreas オブジェクトでの名前は getUsedRangeAreasOrNullObject)</span><span class="sxs-lookup"><span data-stu-id="edc05-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="edc05-159">load()</span><span class="sxs-lookup"><span data-stu-id="edc05-159">load()</span></span>
- <span data-ttu-id="edc05-160">set()</span><span class="sxs-lookup"><span data-stu-id="edc05-160">set()</span></span>
- <span data-ttu-id="edc05-161">setDirty() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="edc05-161">setDirty() (preview)</span></span>
- <span data-ttu-id="edc05-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="edc05-162">toJSON()</span></span>
- <span data-ttu-id="edc05-163">track()</span><span class="sxs-lookup"><span data-stu-id="edc05-163">track()</span></span>
- <span data-ttu-id="edc05-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="edc05-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="edc05-165">RangeArea 固有のプロパティとメソッド</span><span class="sxs-lookup"><span data-stu-id="edc05-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="edc05-166">`RangeAreas` 型には、`Range` オブジェクトには存在しないプロパティとメソッドがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="edc05-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="edc05-167">次にいくつか選択したものを示します。</span><span class="sxs-lookup"><span data-stu-id="edc05-167">The following is a selection of them:</span></span>

- <span data-ttu-id="edc05-168">`areas`: `RangeAreas` オブジェクトが表す全範囲を含む `RangeCollection` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="edc05-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="edc05-169">`RangeCollection` オブジェクトも新しいオブジェクトであり、他の Excel コレクション オブジェクトと類似しています。</span><span class="sxs-lookup"><span data-stu-id="edc05-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="edc05-170">これには、範囲を表す `Range` オブジェクトの配列である `items` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="edc05-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="edc05-171">`areaCount`: `RangeAreas` で指定された範囲の合計数。</span><span class="sxs-lookup"><span data-stu-id="edc05-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="edc05-172">`getOffsetRangeAreas`: [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) と同じように動作します。ただし、`RangeAreas` を返し、元の `RangeAreas` で指定された範囲の 1 つからの各オフセットである範囲を含みます。</span><span class="sxs-lookup"><span data-stu-id="edc05-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="edc05-173">RangeAreas の作成</span><span class="sxs-lookup"><span data-stu-id="edc05-173">Create RangeAreas</span></span>

<span data-ttu-id="edc05-174">`RangeAreas` オブジェクトの作成には、2 つの基本的な方法があります。</span><span class="sxs-lookup"><span data-stu-id="edc05-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="edc05-175">`Worksheet.getRanges()` を呼び出して、範囲のアドレスがコンマで区切られた文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="edc05-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="edc05-176">含める対象の範囲が既に [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) に指定されている場合、文字列にはアドレスではなくその名前を指定することができます。</span><span class="sxs-lookup"><span data-stu-id="edc05-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="edc05-177">`Workbook.getSelectedRanges()` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="edc05-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="edc05-178">このメソッドは、現在アクティブなワークシート上で選択されている全範囲を表す `RangeAreas` を返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="edc05-179">一度 `RangeAreas` オブジェクトを作成すると、`getOffsetRangeAreas` や `getIntersection` など、`RangeAreas` を返すオブジェクト上のメソッドを使用して別のオブジェクトを作成できます。</span><span class="sxs-lookup"><span data-stu-id="edc05-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="edc05-180">`RangeAreas` オブジェクトに新たな範囲を直接追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="edc05-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="edc05-181">たとえば、`RangeAreas.areas` 内のコレクションには `add` メソッドが存在しません。</span><span class="sxs-lookup"><span data-stu-id="edc05-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="edc05-182">`RangeAreas.areas.items` 配列のメンバーの追加または削除を直接試行してはいけません。</span><span class="sxs-lookup"><span data-stu-id="edc05-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="edc05-183">これにより、後でコード内で望ましくない動作が発生します。</span><span class="sxs-lookup"><span data-stu-id="edc05-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="edc05-184">たとえば、追加の `Range` オブジェクトを配列にプッシュすることは可能ですが、エラーが発生します。`RangeAreas` のプロパティとメソッドは、その新しいアイテムがその場所に存在していないかのように動作するためです。</span><span class="sxs-lookup"><span data-stu-id="edc05-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="edc05-185">たとえば、`areaCount` プロパティにはこの方法でプッシュされた範囲は含まれません。また、`RangeAreas.getItemAt(index)` は、`index` が `areasCount-1`より大きい場合、エラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="edc05-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="edc05-186">同様に、`RangeAreas.areas.items` 配列内の `Range` オブジェクトを、参照を取得してその `Range.delete` メソッドを呼び出すという方法で削除すると、バグとなります。`Range` オブジェクトは*削除されます*が、親 `RangeAreas` オブジェクトのプロパティとメソッドは、そのオブジェクトがまだ存在するものとして動作するためです。</span><span class="sxs-lookup"><span data-stu-id="edc05-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="edc05-187">たとえば、コードで `RangeAreas.calculate` を呼び出すと、Office は範囲を計算しようとしますが、範囲オブジェクトが既に存在しないためにエラーとなります。</span><span class="sxs-lookup"><span data-stu-id="edc05-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="edc05-188">複数の範囲でのプロパティの設定</span><span class="sxs-lookup"><span data-stu-id="edc05-188">Set properties on multiple ranges</span></span>

<span data-ttu-id="edc05-189">`RangeAreas` オブジェクトでプロパティを設定すると、`RangeAreas.areas` コレクション内の全範囲の対応するプロパティが設定されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-189">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="edc05-190">次に、複数の範囲にプロパティを設定する例を示します。</span><span class="sxs-lookup"><span data-stu-id="edc05-190">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="edc05-191">この関数は、**F3:F5** と **H3:H5** の範囲を強調表示します。</span><span class="sxs-lookup"><span data-stu-id="edc05-191">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="edc05-192">この例は、`getRanges` に渡す範囲のアドレスをハード コーディングできる場合や実行時に簡単に計算できる場合に適用されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-192">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="edc05-193">たとえば、これが適切なのは次のような場合です。</span><span class="sxs-lookup"><span data-stu-id="edc05-193">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="edc05-194">コードが、既知のテンプレートのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="edc05-194">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="edc05-195">コードが、データのスキーマが既知であるインポート済みデータのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="edc05-195">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="edc05-196">複数の範囲からの特定のセルの取得</span><span class="sxs-lookup"><span data-stu-id="edc05-196">Get special cells from multiple ranges</span></span>

<span data-ttu-id="edc05-197">`RangeAreas` オブジェクトの `getSpecialCells` メソッドと `getSpecialCellsOrNullObject` メソッドは、`Range` オブジェクトの同じ名前のメソッドと同じように機能します。</span><span class="sxs-lookup"><span data-stu-id="edc05-197">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="edc05-198">これらのメソッドでは、`RangeAreas.areas` コレクション内のすべての範囲から、指定された特性を持つセルが返されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-198">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="edc05-199">特殊なセルの詳細については、「[範囲内の特殊なセルの検索](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview)」のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="edc05-199">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) section for more details on special cells.</span></span>

<span data-ttu-id="edc05-200">`RangeAreas` オブジェクトで `getSpecialCells` メソッドまたは `getSpecialCellsOrNullObject` メソッドを呼び出す場合:</span><span class="sxs-lookup"><span data-stu-id="edc05-200">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="edc05-201">最初のパラメーターとして `Excel.SpecialCellType.sameConditionalFormat` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じ条件付き書式を持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-201">There is one small difference in the behavior of the methods when called on a  object instead of a  object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell in the first range in the  collection.</span></span>
- <span data-ttu-id="edc05-202">最初のパラメーターとして `Excel.SpecialCellType.sameDataValidation` を渡した場合、このメソッドでは、`RangeAreas.areas` コレクション内の最初の範囲の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-202">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="edc05-203">RangeAreas のプロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="edc05-203">Read properties of RangeAreas</span></span>

<span data-ttu-id="edc05-204">`RangeAreas` のプロパティ値の読み取りには、注意が必要です。`RangeAreas`内の範囲それぞれで、プロパティの値が異なる可能性があるためです。</span><span class="sxs-lookup"><span data-stu-id="edc05-204">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="edc05-205">一貫性のある値を返すことが*できる*場合には返す、というのが一般的なルールです。</span><span class="sxs-lookup"><span data-stu-id="edc05-205">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="edc05-206">たとえば、次のコードでは、ピンクの RGB コード (`#FFC0CB`) と `true` がコンソールに記録されます。`RangeAreas`オブジェクト内の範囲のどちらも、塗りつぶし色がピンクであり、列全体であるためです。</span><span class="sxs-lookup"><span data-stu-id="edc05-206">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="edc05-207">一貫性を期待できない場合、事態は複雑となります。</span><span class="sxs-lookup"><span data-stu-id="edc05-207">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="edc05-208">`RangeAreas` プロパティの動作は、次の 3 つの原則に従います。</span><span class="sxs-lookup"><span data-stu-id="edc05-208">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="edc05-209">`RangeAreas` オブジェクトのブール値プロパティは、すべてのメンバー範囲でプロパティが true でない限り、`false` を返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-209">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="edc05-210">ブール値以外のプロパティ (`address` プロパティを除く) は、すべてのメンバー範囲で対応するプロパティが同じ値ではない限り、`null` を返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-210">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="edc05-211">`address` プロパティは、メンバー範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="edc05-211">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="edc05-212">たとえば、次のコードでは、1 つの範囲のみが列全体であり、1 つの範囲のみがピンクで塗りつぶされている `RangeAreas` を作成します。</span><span class="sxs-lookup"><span data-stu-id="edc05-212">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="edc05-213">コンソールには、塗りつぶし色の場合は `null`、`isEntireRow` プロパティの場合は `false`、`address` プロパティの場合は "Sheet1!F3:F5, Sheet1!H:H" ("Sheet1" はシート名) が表示されます。</span><span class="sxs-lookup"><span data-stu-id="edc05-213">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="edc05-214">関連項目</span><span class="sxs-lookup"><span data-stu-id="edc05-214">See also</span></span>

- [<span data-ttu-id="edc05-215">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="edc05-215">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="edc05-216">Excel JavaScript API を使用して範囲を操作する (基本)</span><span class="sxs-lookup"><span data-stu-id="edc05-216">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="edc05-217">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="edc05-217">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)