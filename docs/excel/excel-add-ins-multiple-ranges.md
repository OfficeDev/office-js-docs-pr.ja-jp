---
title: Excel アドインで複数の範囲を同時に操作する
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: 37f9c8a9f3127d78e1cc794aea9e6d1502cdeaf9
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270979"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="d1d8b-102">Excel アドインで複数の範囲を同時に操作する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="d1d8b-103">Excel JavaScript ライブラリを使用すると、同時に複数の範囲に対してアドインによる操作の実行とプロパティの設定が可能になります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="d1d8b-104">範囲は連続している必要はありません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="d1d8b-105">コードがよりシンプルになることに加え、この方法でプロパティを設定すれば、各範囲に同じプロパティを個別に設定する方法よりも処理速度が格段に速くなります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="d1d8b-106">この記事で説明する API には、**Office 2016 クイック実行バージョン 1809 Build 10820.20000** 以降が必要です </span><span class="sxs-lookup"><span data-stu-id="d1d8b-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="d1d8b-107">([Office Insider プログラム](https://products.office.com/office-insider)に参加して、適切なビルドを取得することが必要な場合があります)。また、Office JavaScript ライブラリのベータ版を [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) からロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="d1d8b-108">最後に、これらの API セットに関する参照ページはまだありません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="d1d8b-109">ただし、定義の種類ファイル [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) に説明が含まれています。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="d1d8b-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="d1d8b-110">RangeAreas</span></span>

<span data-ttu-id="d1d8b-111">範囲のセット (連続している必要はなし) は、`Excel.RangeAreas` オブジェクトで表されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="d1d8b-112">`Range` 型と同様のプロパティとメソッドを持ちますが (多くの場合は同じまたは類似した名前)、以下に対しては調整が行われています。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="d1d8b-113">プロパティのデータ型と、セッターとゲッターの動作。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="d1d8b-114">メソッド パラメーターのデータ型と、メソッドの動作。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="d1d8b-115">メソッドの戻り値のデータ型。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-115">The data types of method return values.</span></span>

<span data-ttu-id="d1d8b-116">次にいくつか例を示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-116">Some examples:</span></span>

- <span data-ttu-id="d1d8b-117">`RangeAreas` には `address` プロパティがあり、`Range.address` プロパティのように 1 つのアドレスを返すのではなく、複数の範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="d1d8b-118">`RangeAreas` には、一貫性がある場合、`RangeAreas` に指定された全範囲のデータ検証を表す `DataValidation` オブジェクトを返す `dataValidation` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="d1d8b-119">`RangeAreas` に指定された全範囲に同じ `DataValidation` オブジェクトが適用されていない場合、このプロパティは `null` となります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="d1d8b-120">これは、`RangeAreas` オブジェクトに関する、汎用的ではありませんが一般的な原則です: *`RangeAreas` に指定された全範囲のプロパティの値に一貫性がない場合、`null` となります*。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="d1d8b-121">より詳しい情報といくつかの例外については、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="d1d8b-122">`RangeAreas.cellCount` は、`RangeAreas` に指定された全範囲の合計セル数を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="d1d8b-123">`RangeAreas.calculate` は、`RangeAreas` に指定された全範囲のセルを再計算します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="d1d8b-124">`RangeAreas.getEntireColumn` と `RangeAreas.getEntireRow` は、`RangeAreas` に指定された全範囲のセルの列 (または行) すべてを表す、別の `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="d1d8b-125">たとえば、`RangeAreas` が "A1:C4" と "F14:L15" を表す場合、`RangeAreas.getEntireColumn` は "A:C" と "F:L" を表す `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="d1d8b-126">`RangeAreas.copyFrom` は、コピー操作のコピー元範囲を表す `Range` または `RangeAreas` パラメーターのいずれかを取得できます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="d1d8b-127">RangeAreas でも利用可能な Range メンバーの全リスト</span><span class="sxs-lookup"><span data-stu-id="d1d8b-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="d1d8b-128">プロパティ</span><span class="sxs-lookup"><span data-stu-id="d1d8b-128">Properties</span></span>

<span data-ttu-id="d1d8b-129">リストにあるプロパティを読み取るコードを書く前に、「[RangeAreas のプロパティの読み取り](#read-properties-of-rangeareas)」の内容を理解しておいてください。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="d1d8b-130">繰り返される内容について細かい注意点があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="d1d8b-131">address</span><span class="sxs-lookup"><span data-stu-id="d1d8b-131">address</span></span>
- <span data-ttu-id="d1d8b-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="d1d8b-132">addressLocal</span></span>
- <span data-ttu-id="d1d8b-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="d1d8b-133">cellCount</span></span>
- <span data-ttu-id="d1d8b-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="d1d8b-134">conditionalFormats</span></span>
- <span data-ttu-id="d1d8b-135">context</span><span class="sxs-lookup"><span data-stu-id="d1d8b-135">context</span></span>
- <span data-ttu-id="d1d8b-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="d1d8b-136">dataValidation</span></span>
- <span data-ttu-id="d1d8b-137">format</span><span class="sxs-lookup"><span data-stu-id="d1d8b-137">format</span></span>
- <span data-ttu-id="d1d8b-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="d1d8b-138">isEntireColumn</span></span>
- <span data-ttu-id="d1d8b-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="d1d8b-139">isEntireRow</span></span>
- <span data-ttu-id="d1d8b-140">style</span><span class="sxs-lookup"><span data-stu-id="d1d8b-140">style</span></span>
- <span data-ttu-id="d1d8b-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="d1d8b-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="d1d8b-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1d8b-142">Methods</span></span>

<span data-ttu-id="d1d8b-143">プレビュー段階の Range メソッドについてはマークが付いています。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="d1d8b-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-144">calculate()</span></span>
- <span data-ttu-id="d1d8b-145">clear()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-145">clear()</span></span>
- <span data-ttu-id="d1d8b-146">convertDataTypeToText() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="d1d8b-147">convertToLinkedDataType() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="d1d8b-148">copyFrom() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="d1d8b-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-149">getEntireColumn()</span></span>
- <span data-ttu-id="d1d8b-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-150">getEntireRow()</span></span>
- <span data-ttu-id="d1d8b-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-151">getIntersection()</span></span>
- <span data-ttu-id="d1d8b-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="d1d8b-153">getOffsetRange() (RangeAreas オブジェクトでの名前は getOffsetRangeAreas)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="d1d8b-154">getSpecialCells() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="d1d8b-155">getSpecialCellsOrNullObject() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="d1d8b-156">getTables() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-156">getTables() (preview)</span></span>
- <span data-ttu-id="d1d8b-157">getUsedRange() (RangeAreas オブジェクトでの名前は getUsedRangeAreas)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="d1d8b-158">getUsedRangeOrNullObject() (RangeAreas オブジェクトでの名前は getUsedRangeAreasOrNullObject)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="d1d8b-159">load()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-159">load()</span></span>
- <span data-ttu-id="d1d8b-160">set()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-160">Set</span></span>
- <span data-ttu-id="d1d8b-161">setDirty() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d1d8b-161">setDirty() (preview)</span></span>
- <span data-ttu-id="d1d8b-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-162">toJSON()</span></span>
- <span data-ttu-id="d1d8b-163">track()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-163">track</span></span>
- <span data-ttu-id="d1d8b-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="d1d8b-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="d1d8b-165">RangeArea 固有のプロパティとメソッド</span><span class="sxs-lookup"><span data-stu-id="d1d8b-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="d1d8b-166">`RangeAreas` 型には、`Range` オブジェクトには存在しないプロパティとメソッドがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="d1d8b-167">次にいくつか選択したものを示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-167">The following is a selection of them:</span></span>

- <span data-ttu-id="d1d8b-168">`areas`: `RangeAreas` オブジェクトが表す全範囲を含む `RangeCollection` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="d1d8b-169">`RangeCollection` オブジェクトも新しいオブジェクトであり、他の Excel コレクション オブジェクトと類似しています。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="d1d8b-170">これには、範囲を表す `Range` オブジェクトの配列である `items` プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="d1d8b-171">`areaCount`: `RangeAreas` で指定された範囲の合計数。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="d1d8b-172">`getOffsetRangeAreas`: [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) と同じように動作します。ただし、`RangeAreas` を返し、元の `RangeAreas` で指定された範囲の 1 つからの各オフセットである範囲を含みます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="d1d8b-173">RangeAreas の作成とプロパティの設定</span><span class="sxs-lookup"><span data-stu-id="d1d8b-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="d1d8b-174">`RangeAreas` オブジェクトの作成には、2 つの基本的な方法があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="d1d8b-175">`Worksheet.getRanges()` を呼び出して、範囲のアドレスがコンマで区切られた文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="d1d8b-176">含める対象の範囲が既に [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) に指定されている場合、文字列にはアドレスではなくその名前を指定することができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="d1d8b-177">`Workbook.getSelectedRanges()` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="d1d8b-178">このメソッドは、現在アクティブなワークシート上で選択されている全範囲を表す `RangeAreas` を返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="d1d8b-179">一度 `RangeAreas` オブジェクトを作成すると、`getOffsetRangeAreas` や `getIntersection` など、`RangeAreas` を返すオブジェクト上のメソッドを使用して別のオブジェクトを作成できます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="d1d8b-180">`RangeAreas` オブジェクトに新たな範囲を直接追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="d1d8b-181">たとえば、`RangeAreas.areas` 内のコレクションには `add` メソッドが存在しません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="d1d8b-182">`RangeAreas.areas.items` 配列のメンバーの追加または削除を直接試行してはいけません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="d1d8b-183">これにより、後でコード内で望ましくない動作が発生します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="d1d8b-184">たとえば、追加の `Range` オブジェクトを配列にプッシュすることは可能ですが、エラーが発生します。`RangeAreas` のプロパティとメソッドは、その新しいアイテムがその場所に存在していないかのように動作するためです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="d1d8b-185">たとえば、`areaCount` プロパティにはこの方法でプッシュされた範囲は含まれません。また、`RangeAreas.getItemAt(index)` は、`index` が `areasCount-1`より大きい場合、エラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="d1d8b-186">同様に、`RangeAreas.areas.items` 配列内の `Range` オブジェクトを、参照を取得してその `Range.delete` メソッドを呼び出すという方法で削除すると、バグとなります。`Range` オブジェクトは*削除されます*が、親 `RangeAreas` オブジェクトのプロパティとメソッドは、そのオブジェクトがまだ存在するものとして動作するためです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="d1d8b-187">たとえば、コードで `RangeAreas.calculate` を呼び出すと、Office は範囲を計算しようとしますが、範囲オブジェクトが既に存在しないためにエラーとなります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="d1d8b-188">`RangeAreas` に対してプロパティを設定すると、`RangeAreas.areas` コレクション内の全範囲の対応するプロパティが設定されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="d1d8b-189">次に、複数の範囲にプロパティを設定する例を示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="d1d8b-190">この関数は、**F3:F5** と **H3:H5** の範囲を強調表示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="d1d8b-191">この例は、`getRanges` に渡す範囲のアドレスをハード コーディングできる場合や実行時に簡単に計算できる場合に適用されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="d1d8b-192">たとえば、これが適切なのは次のような場合です。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="d1d8b-193">コードが、既知のテンプレートのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="d1d8b-194">コードが、データのスキーマが既知であるインポート済みデータのコンテキスト内で実行される。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="d1d8b-195">コーディング時に操作対象の範囲がわからない場合は、実行時に特定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="d1d8b-196">次のセクションでは、そのような場合について説明します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="d1d8b-197">プログラムを使用して範囲を特定する</span><span class="sxs-lookup"><span data-stu-id="d1d8b-197">Discover range areas programmatically</span></span>

<span data-ttu-id="d1d8b-198">`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、セルの特性とセル値の種類を基に、操作対象のセルを実行時に特定することができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="d1d8b-199">次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="d1d8b-200">このうち最初のものを使用する例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="d1d8b-201">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-201">About this code, note:</span></span>

- <span data-ttu-id="d1d8b-202">検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="d1d8b-203">`Excel.SpecialCellType` 列挙からの値の文字列バージョンをパラメーターとして `getSpecialCells` に渡します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="d1d8b-204">代わりに渡すことができる他の値には、空のセルの場合は "Blanks"、数式ではなくリテラル値を含むセルの場合は "Constants"、`usedRange` 内の最初のセルと同じ条件付き書式を持つセルの場合は "SameConditionalFormat" などがあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="d1d8b-205">最初のセルとは、左上隅のセルです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="d1d8b-206">列挙内の値の完全なリストについては、[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="d1d8b-207">`getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="d1d8b-208">範囲内に対象の特性を持つセルが*まったくない*場合もあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="d1d8b-209">`getSpecialCells` で対象のセルが見つからないと、**ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="d1d8b-210">この場合、制御のフローが `catch` ブロック/メソッドに移ります (存在する場合)。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="d1d8b-211">存在しない場合は、このエラーにより関数が停止します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="d1d8b-212">対象の特性を持つセルがない場合はエラーをスローするという動作が求められるシナリオもあるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="d1d8b-213">ただし、一般的ではありませんが、一致するセルがないということが通常であるようなシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="d1d8b-214">このようなシナリオの場合、`getSpecialCellsOrNullObject` メソッドを使用し、`RangeAreas.isNullObject` プロパティをテストします。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="d1d8b-215">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-215">The following is an example.</span></span> <span data-ttu-id="d1d8b-216">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-216">Note about this code:</span></span>

- <span data-ttu-id="d1d8b-217">`getSpecialCellsOrNullObject` メソッドは常にプロキシ オブジェクトを返します。そのため、通常の JavaScript 使用環境では `null` となることはありません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="d1d8b-218">ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="d1d8b-219">`isNullObject` プロパティをテストする*前*に、`context.sync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="d1d8b-220">これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="d1d8b-221">ただし、*明示的*に `isNullObject` プロパティをロードする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="d1d8b-222">`load` がオブジェクトに対して呼び出されていない場合であっても、プロパティは `context.sync` によって自動的にロードされます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="d1d8b-223">詳細については、「[\*OrNullObject メソッド](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-223">For more information see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)</span></span>
- <span data-ttu-id="d1d8b-224">このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="d1d8b-225">次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="d1d8b-226">わかりやすくするため、この記事内のすべての他の例では、`getSpecialCells` メソッドを `getSpecialCellsOrNullObject` の代わりに使用しています。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="d1d8b-227">セルの値の型に応じて対象のセルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="d1d8b-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="d1d8b-228">オプションの 2 つめのパラメーター (列挙型 `Excel.SpecialCellValueType`) を使用すると、対象のセルをさらに絞り込むことができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="d1d8b-229">このパラメーターは、"Formulas" または "Constants" を`getSpecialCells` または `getSpecialCellsOrNullObject` に渡す場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="d1d8b-230">このパラメーターにより、特定の型の値を持つセルのみ対象として指定することができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="d1d8b-231">4 つの基本的な型があります: "Error"、"Logical" (ブール値を意味します)、"Numbers"、"Text" です </span><span class="sxs-lookup"><span data-stu-id="d1d8b-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="d1d8b-232">(列挙の場合はこの 4 つ以外の値もあります。詳細は後述します)。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="d1d8b-233">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-233">About this code, note:</span></span>

- <span data-ttu-id="d1d8b-234">リテラル数値を持つセルのみ強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="d1d8b-235">数式 (結果が数字の場合であっても)、ブール値、テキストを持つセル、およびエラー状態にあるセルは強調表示されません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="d1d8b-236">コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="d1d8b-237">テキスト値のセルすべてとブール値 ("Logical") のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="d1d8b-238">`Excel.SpecialCellValueType` 列挙に含まれる値を使用すると、型を組み合わせることができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="d1d8b-239">たとえば、"LogicalText" を使用すると、すべてのブール値のセルとテキスト値のセルを対象とすることができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="d1d8b-240">4 つの基本的な型のうち、任意の 2 つまたは 3 つの型を組み合わせることができます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="d1d8b-241">基本的な型を組み合わせるこれらの列挙値の名前は、常にアルファベット順で指定します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="d1d8b-242">したがって、エラー値、テキスト値、ブール値のセルを組み合わせる場合は "ErrorLogicalText" を使用します。"LogicalErrorText" や "TextErrorLogical" とはしてはいけません。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="d1d8b-243">既定のパラメーターである "All" は、4 つの型すべてを組み合わせます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="d1d8b-244">次の例では、結果が数値またはブール値となる数式を含むすべてのセルが強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="d1d8b-245">`Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` パラメーターが "Formulas" または "Constants" の場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-245">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="d1d8b-246">RangeAreas 内の RangeAreas を取得する</span><span class="sxs-lookup"><span data-stu-id="d1d8b-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="d1d8b-247">`RangeAreas` 型自体には、同じ 2 つのパラメーターを使用する `getSpecialCells` および `getSpecialCellsOrNullObject` メソッドもあります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="d1d8b-248">これらのメソッドは、`RangeAreas.areas` コレクション内の全範囲から対象のセルをすべて返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="d1d8b-249">`Range`オブジェクトではなく `RangeAreas` に対して呼び出された場合のメソッドの動作には、少し異なる点が 1 つあります。最初のパラメーターとして "SameConditionalFormat" を渡した場合、*`RangeAreas.areas` コレクション内の最初の範囲*の左上隅のセルと同じ条件付き書式を持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="d1d8b-250">同じ点が "SameDataValidation" にも適用されます。`Range.getSpecialCells` にこれを渡すと、*範囲内*の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="d1d8b-251">一方、`RangeAreas.getSpecialCells` に渡した場合は、*`RangeAreas.areas` コレクション内の最初の範囲*の左上隅のセルと同じデータ検証ルールを持つセルがすべて返されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="d1d8b-252">RangeAreas のプロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="d1d8b-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="d1d8b-253">`RangeAreas` のプロパティ値の読み取りには、注意が必要です。`RangeAreas`内の範囲それぞれで、プロパティの値が異なる可能性があるためです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="d1d8b-254">一貫性のある値を返すことが*できる*場合には返す、というのが一般的なルールです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="d1d8b-255">たとえば、次のコードでは、ピンクの RGB コード (`#FFC0CB`) と `true` がコンソールに記録されます。`RangeAreas`オブジェクト内の範囲のどちらも、塗りつぶし色がピンクであり、列全体であるためです。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="d1d8b-256">一貫性を期待できない場合、事態は複雑となります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="d1d8b-257">`RangeAreas` プロパティの動作は、次の 3 つの原則に従います。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="d1d8b-258">`RangeAreas` オブジェクトのブール値プロパティは、すべてのメンバー範囲でプロパティが true でない限り、`false` を返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="d1d8b-259">ブール値以外のプロパティ (`address` プロパティを除く) は、すべてのメンバー範囲で対応するプロパティが同じ値ではない限り、`null` を返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="d1d8b-260">`address` プロパティは、メンバー範囲のアドレスをコンマで区切った文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="d1d8b-261">たとえば、次のコードでは、1 つの範囲のみが列全体であり、1 つの範囲のみがピンクで塗りつぶされている `RangeAreas` を作成します。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="d1d8b-262">コンソールには、塗りつぶし色の場合は `null`、`isEntireRow` プロパティの場合は `false`、`address` プロパティの場合は "Sheet1!F3:F5, Sheet1!H:H" ("Sheet1" はシート名) が表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="d1d8b-263">関連項目</span><span class="sxs-lookup"><span data-stu-id="d1d8b-263">See also</span></span>

- [<span data-ttu-id="d1d8b-264">Excel の JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="d1d8b-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="d1d8b-265">Excel.​Range class</span><span class="sxs-lookup"><span data-stu-id="d1d8b-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="d1d8b-266">[RangeAreas オブジェクト (Excel の JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー段階の間は、このリンクは機能しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="d1d8b-267">その場合は、代わりに [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="d1d8b-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>