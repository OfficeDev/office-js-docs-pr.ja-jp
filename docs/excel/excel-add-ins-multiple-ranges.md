---
title: Excel のアドインで同時に複数の範囲の操作をします。
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016459"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="810c4-102">Excel のアドイン (プレビュー) で同時に複数のセル範囲を操作します。</span><span class="sxs-lookup"><span data-stu-id="810c4-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="810c4-103">Excel の JavaScript ライブラリは、アドインに操作を実行し、同時に複数の範囲のプロパティを設定するようにします。</span><span class="sxs-lookup"><span data-stu-id="810c4-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="810c4-104">範囲が隣接している必要はありません。</span><span class="sxs-lookup"><span data-stu-id="810c4-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="810c4-105">コードを簡単にするには、このプロパティの方法が、範囲ごとに個別に同じプロパティを設定するよりもはるかに高速に実行されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="810c4-106">この資料に記載されている APIは、 **Office 2016 クイック実行バージョン 1809 10820.20000 の構築** 以降を必要とします。</span><span class="sxs-lookup"><span data-stu-id="810c4-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="810c4-107">(適切なビルドを取得するため、 [Office 内部プログラム](https://products.office.com/office-insider) に参加する必要があります。)　また、 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)からベータ版の Office の JavaScript ライブラリを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="810c4-108">最後に、これらの API は、参照ページはまだ必要はありません。</span><span class="sxs-lookup"><span data-stu-id="810c4-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="810c4-109">次の定義型ファイルは、それらについての説明です: [ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)です。</span><span class="sxs-lookup"><span data-stu-id="810c4-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="810c4-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="810c4-110">RangeAreas</span></span>

<span data-ttu-id="810c4-111">(場合によっては連続していない) 範囲のセットは、 `Excel.RangeAreas` オブジェクトで表示されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="810c4-112">型(多くの場合、同じまたは同様の名前を持つ)に似たプロパティとメソッドを持っていますが、以下の補正が加えられました。`Range`</span><span class="sxs-lookup"><span data-stu-id="810c4-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="810c4-113">プロパティのデータ型とセッターとゲッターの動作です。</span><span class="sxs-lookup"><span data-stu-id="810c4-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="810c4-114">メソッドパラメーターのデータ型と、メソッドの動作です。</span><span class="sxs-lookup"><span data-stu-id="810c4-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="810c4-115">メソッドのデータ型は、値を返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-115">The data types of method return values.</span></span>

<span data-ttu-id="810c4-116">例:</span><span class="sxs-lookup"><span data-stu-id="810c4-116">Some examples:</span></span>

- <span data-ttu-id="810c4-117">`RangeAreas` プロパティとともに1 つのアドレスとしてではなく、範囲のアドレスのコンマ区切りの文字列を返す`address` プロパティを持ちます。`Range.address`</span><span class="sxs-lookup"><span data-stu-id="810c4-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="810c4-118">`RangeAreas` 一貫性がある場合、`RangeAreas`  内のすべての範囲のデータの入力規則を表す`DataValidation` オブジェクトを返す`dataValidation` プロパティを持ちます。</span><span class="sxs-lookup"><span data-stu-id="810c4-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="810c4-119">同一な `DataValidation` オブジェクトが、 `RangeAreas`内のすべての範囲に適用されない場合、プロパティは `null` です。</span><span class="sxs-lookup"><span data-stu-id="810c4-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="810c4-120">全般的に、汎用的でない場合、 `RangeAreas` オブジェクトの原則は: *もしプロパティが、 `RangeAreas`内のすべての範囲の値に一貫性を持っていない場合、 `null`です*。</span><span class="sxs-lookup"><span data-stu-id="810c4-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="810c4-121">いくつかの例外の詳細については、 [RangeAreas のプロパティを読み取り中](#reading-properties-of-rangeareas) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="810c4-121">See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="810c4-122">`RangeAreas.cellCount` 内のすべての範囲内のセルの合計数を取得します。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="810c4-123">`RangeAreas.calculate` 内のすべての範囲のセルを再計算します。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="810c4-124">`RangeAreas.getEntireColumn` また、 `RangeAreas.getEntireRow` は、 `RangeAreas`内のすべての範囲のすべての列 (または行) を表す他の `RangeAreas` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="810c4-125">例えば、 `RangeAreas` が”A1: C4"と”F14:L15”を表し、`RangeAreas.getEntireColumn` が "A:C"と"F:L"を表すオブジェクト `RangeAreas`を返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="810c4-126">`RangeAreas.copyFrom` コピー操作のソース範囲を表すパラメーター`Range` または `RangeAreas` のいずれがをとることができます。</span><span class="sxs-lookup"><span data-stu-id="810c4-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="810c4-127">RangeAreas でも利用できる範囲のメンバーの完全なリスト</span><span class="sxs-lookup"><span data-stu-id="810c4-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="810c4-128">プロパティ</span><span class="sxs-lookup"><span data-stu-id="810c4-128">Properties</span></span>

<span data-ttu-id="810c4-129">リストされた任意のプロパティを読み取るコードを記述する前に、「[RangeAreas のプロパティを読み取る](#reading-properties-of-rangeareas)」をご確認ください。</span><span class="sxs-lookup"><span data-stu-id="810c4-129">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="810c4-130">返されるものは微妙に異なります。</span><span class="sxs-lookup"><span data-stu-id="810c4-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="810c4-131">address</span><span class="sxs-lookup"><span data-stu-id="810c4-131">address</span></span>
- <span data-ttu-id="810c4-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="810c4-132">addressLocal</span></span>
- <span data-ttu-id="810c4-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="810c4-133">cellCount</span></span>
- <span data-ttu-id="810c4-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="810c4-134">conditionalFormats</span></span>
- <span data-ttu-id="810c4-135">context</span><span class="sxs-lookup"><span data-stu-id="810c4-135">context</span></span>
- <span data-ttu-id="810c4-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="810c4-136">dataValidation</span></span>
- <span data-ttu-id="810c4-137">format</span><span class="sxs-lookup"><span data-stu-id="810c4-137">format</span></span>
- <span data-ttu-id="810c4-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="810c4-138">isEntireColumn</span></span>
- <span data-ttu-id="810c4-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="810c4-139">isEntireRow</span></span>
- <span data-ttu-id="810c4-140">style</span><span class="sxs-lookup"><span data-stu-id="810c4-140">style</span></span>
- <span data-ttu-id="810c4-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="810c4-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="810c4-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="810c4-142">Methods</span></span>

<span data-ttu-id="810c4-143">プレビューの範囲メソッドを示します。</span><span class="sxs-lookup"><span data-stu-id="810c4-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="810c4-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="810c4-144">calculate()</span></span>
- <span data-ttu-id="810c4-145">clear()</span><span class="sxs-lookup"><span data-stu-id="810c4-145">clear()</span></span>
- <span data-ttu-id="810c4-146">convertDataTypeToText() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="810c4-147">convertToLinkedDataType() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="810c4-148">copyFrom() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="810c4-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="810c4-149">getEntireColumn()</span></span>
- <span data-ttu-id="810c4-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="810c4-150">getEntireRow()</span></span>
- <span data-ttu-id="810c4-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="810c4-151">getIntersection()</span></span>
- <span data-ttu-id="810c4-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="810c4-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="810c4-153">getOffsetRange() (RangeAreas オブジェクトの named getOffsetRangeAreas を名前付け)</span><span class="sxs-lookup"><span data-stu-id="810c4-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="810c4-154">getSpecialCells() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="810c4-155">getSpecialCellsOrNullObject() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="810c4-156">getTables() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-156">getTables() (preview)</span></span>
- <span data-ttu-id="810c4-157">getUsedRange() (RangeAreas オブジェクトの getUsedRangeAreas を名前付け)</span><span class="sxs-lookup"><span data-stu-id="810c4-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="810c4-158">getUsedRangeOrNullObject() (RangeAreas オブジェクトの named getUsedRangeAreasOrNullObject を名前付け)</span><span class="sxs-lookup"><span data-stu-id="810c4-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="810c4-159">load()</span><span class="sxs-lookup"><span data-stu-id="810c4-159">load()</span></span>
- <span data-ttu-id="810c4-160">set()</span><span class="sxs-lookup"><span data-stu-id="810c4-160">set\*</span></span>
- <span data-ttu-id="810c4-161">setDirty() (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="810c4-161">setDirty() (preview)</span></span>
- <span data-ttu-id="810c4-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="810c4-162">toJSON()</span></span>
- <span data-ttu-id="810c4-163">track()</span><span class="sxs-lookup"><span data-stu-id="810c4-163">track</span></span>
- <span data-ttu-id="810c4-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="810c4-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="810c4-165">RangeArea に固有のプロパティおよびメソッド</span><span class="sxs-lookup"><span data-stu-id="810c4-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="810c4-166">`RangeAreas` 型には、`Range` オブジェクト上にないいくつかのプロパティとメソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="810c4-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object:</span></span> <span data-ttu-id="810c4-167">次はその一部です。</span><span class="sxs-lookup"><span data-stu-id="810c4-167">The following is a selection of them:</span></span>

- <span data-ttu-id="810c4-168">`areas``RangeAreas` オブジェクトで表される範囲のすべてを含む `RangeCollection` オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="810c4-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="810c4-169">オブジェクトは、新しく、他の Excel のコレクション オブジェクトに似ています。`RangeCollection`</span><span class="sxs-lookup"><span data-stu-id="810c4-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="810c4-170">範囲を表す `Range` オブジェクトの配列である`items` プロパティを持ちます。</span><span class="sxs-lookup"><span data-stu-id="810c4-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="810c4-171">`areaCount`: 範囲内の合計数は、 `RangeAreas`です。</span><span class="sxs-lookup"><span data-stu-id="810c4-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="810c4-172">`getOffsetRangeAreas`:[   が返され、元の](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) `RangeAreas`   の範囲の一つからそれぞれからのオフセットの範囲を含む点を除き、 Range.getOffsetRange`RangeAreas` と同じように動作します。</span><span class="sxs-lookup"><span data-stu-id="810c4-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="810c4-173">RangeAreas の作成と、プロパティの設定</span><span class="sxs-lookup"><span data-stu-id="810c4-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="810c4-174">オブジェクトを2 つの基本的な方法で作成することができます。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="810c4-175">を呼び出し、コンマで区切られた範囲のアドレスを使用して文字列を渡します。`Worksheet.getRanges()`</span><span class="sxs-lookup"><span data-stu-id="810c4-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="810c4-176">含めたい任意の範囲が、 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)された場合は、文字列内のアドレスではなく、名前を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="810c4-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="810c4-177">`Workbook.getSelectedRanges()` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="810c4-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="810c4-178">このメソッドは、現在アクティブなワークシートで選択されているすべての範囲を表す `RangeAreas` を返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="810c4-179">オブジェクトを作成したら、  `getOffsetRangeAreas` と `getIntersection`のような`RangeAreas` を返すオブジェクト上のメソッドを使用して他のユーザーを作成することができます。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="810c4-180">オブジェクトに、追加の範囲を直接追加することはできません。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="810c4-181">例えば、`RangeAreas.areas`内のコレクションは、 `add` メソッドを持ちません。</span><span class="sxs-lookup"><span data-stu-id="810c4-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="810c4-182">の配列のメンバーを、直接追加または削除しないようにしてください。`RangeAreas.areas.items`</span><span class="sxs-lookup"><span data-stu-id="810c4-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="810c4-183">コード内で望ましくない動作をしてしまいます。</span><span class="sxs-lookup"><span data-stu-id="810c4-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="810c4-184">たとえば、さらに配列上に追加の `Range` オブジェクトをプッシュすることは可能ですが、 `RangeAreas` プロパティやメソッドは、新しいアイテムがない場合と同様に動作するため、これを行うとエラーが発生します。</span><span class="sxs-lookup"><span data-stu-id="810c4-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="810c4-185">例えば、 `areaCount` プロパティには、この方法によりプッシュされた範囲を含みません。 `index` よりも大きい `areasCount-1`場合、 `RangeAreas.getItemAt(index)` は、エラーをスローします。</span><span class="sxs-lookup"><span data-stu-id="810c4-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="810c4-186">同様に、参照を取得し、その`Range.delete` メソッドを呼び出して、  `RangeAreas.areas.items`内の`Range` オブジェクトを削除すると、バグが発生します:  `Range` オブジェクト *は、* 削除されましたが、親の `RangeAreas` オブジェクトのプロパティとメソッドは、それがまだ存在するかのように動作、またはしようとします。</span><span class="sxs-lookup"><span data-stu-id="810c4-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="810c4-187">例えば、コードが `RangeAreas.calculate`を呼び出した場合、Office は、範囲を計算しようとしますが、がエラーが発生し、range オブジェクトは失われます。</span><span class="sxs-lookup"><span data-stu-id="810c4-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="810c4-188">上のプロパティの設定は、 `RangeAreas.areas` コレクション上のすべての範囲に対応するプロパティを設定します。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="810c4-189">次は、複数の範囲のプロパティの設定の例です。</span><span class="sxs-lookup"><span data-stu-id="810c4-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="810c4-190">関数には、 **F3:F5** と **H3:H5**の範囲が強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="810c4-191">この例では、 `getRanges`に渡す範囲のアドレスを渡すハード コードできるか 、簡単に実行時に自動的に計算できるシナリオを適用します。</span><span class="sxs-lookup"><span data-stu-id="810c4-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="810c4-192">これが正しいであろうシナリオの一部は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="810c4-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="810c4-193">コードは、既知のテンプレートのコンテキストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="810c4-194">コードは、データのスキーマがわかっているインポートされたデータのコンテキストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="810c4-195">コーディング時に動作する必要がある範囲を知ることはできません、実行時に検出する必要があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="810c4-196">次のセクションでは、これらのシナリオについて説明します。</span><span class="sxs-lookup"><span data-stu-id="810c4-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="810c4-197">範囲の領域をプログラムで検出します。</span><span class="sxs-lookup"><span data-stu-id="810c4-197">Discover range areas programmatically</span></span>

<span data-ttu-id="810c4-198">`Range.getSpecialCells()` と `Range.getSpecialCellsOrNullObject()` メソッドを使用すると、実行時に、セルの特性とセルの値の型を基に操作し実行したい範囲を検索できます。</span><span class="sxs-lookup"><span data-stu-id="810c4-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="810c4-199">TypeScriptデータ型のファイルからのメソッドのシグネチャを次に示します。</span><span class="sxs-lookup"><span data-stu-id="810c4-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="810c4-200">次は、最初のシグネチャを使用する場合の例です。</span><span class="sxs-lookup"><span data-stu-id="810c4-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="810c4-201">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="810c4-201">About this code, note:</span></span>

- <span data-ttu-id="810c4-202">最初の呼び出しで検索する必要があるシートの一部を制限して、その範囲のみで `Worksheet.getUsedRange`  `getSpecialCells` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="810c4-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="810c4-203">列挙型からの値の文字列バージョンをパラメーターとして`getSpecialCells`  に渡します。`Excel.SpecialCellType`</span><span class="sxs-lookup"><span data-stu-id="810c4-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="810c4-204">代わりに渡される他の値のいくつかは、空白のセルには「空白」、数式のかわりにリテラル値を持つセルには「定数」、 `usedRange`の最初のセルと同じ条件付き書式が設定されるセルには"SameConditionalFormat"です。</span><span class="sxs-lookup"><span data-stu-id="810c4-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="810c4-205">最初のセルは、上の左端のセルです。</span><span class="sxs-lookup"><span data-stu-id="810c4-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="810c4-206">列挙型の値の完全な一覧は、 [ベータ版の office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="810c4-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="810c4-207">`getSpecialCells` メソッドは、 `RangeAreas` オブジェクトを返します。数式を入力した全てのセルは、すべて連続していない場合でも、ピンク色に色分け表示されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="810c4-208">場合によっては、対象となる特性を持つ *任意* のセルが範囲にありません。</span><span class="sxs-lookup"><span data-stu-id="810c4-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="810c4-209">`getSpecialCells` が対象となるセルをどうしても見つけられない場合、 **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="810c4-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="810c4-210">これは、一つだけある場合、コントロールのフローを `catch` ブロックまたはメソッドにそらします。</span><span class="sxs-lookup"><span data-stu-id="810c4-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="810c4-211">そうでない場合、エラーは、機能を停止します。</span><span class="sxs-lookup"><span data-stu-id="810c4-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="810c4-212">エラーをスローすることが、対象となる特性を持つセルが存在しない場合に、正確に必要されることであるシナリオがある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="810c4-213">通常の動作では、シナリオでは、一致するセルがない場合もありますが、これはおそらく一般的ではありません。コードは、この可能性を確認し、エラーをスローすることがなく適切に処理すること必要があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="810c4-214">これらのシナリオでは、 `getSpecialCellsOrNullObject` メソッドを使用し、 `RangeAreas.isNullObject` プロパティをテストします。</span><span class="sxs-lookup"><span data-stu-id="810c4-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="810c4-215">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="810c4-215">The following is an example.</span></span> <span data-ttu-id="810c4-216">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="810c4-216">Note about this code:</span></span>

- <span data-ttu-id="810c4-217">メソッドは、常にプロキシ オブジェクトを返します。したがって、通常 JavaScript という意味では、決して `null` にはなりません。`getSpecialCellsOrNullObject`</span><span class="sxs-lookup"><span data-stu-id="810c4-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="810c4-218">一致するセルが見つからない場合は、 `isNullObject` オブジェクトのプロパティが、 `true`に設定されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="810c4-219">プロパティをテストする前に、 を呼び出します。`context.sync`\*\*`isNullObject`</span><span class="sxs-lookup"><span data-stu-id="810c4-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="810c4-220">読みだすために常にプロパティを読み込み同期させる必要があるため、これはすべての `*OrNullObject` メソッドとプロパティの必要条件です。</span><span class="sxs-lookup"><span data-stu-id="810c4-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="810c4-221">ただし、 `isNullObject` プロパティを *明示的に* 読み込む必要はありません。</span><span class="sxs-lookup"><span data-stu-id="810c4-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="810c4-222">がオブジェクト上に呼び出されない場合でも、 `context.sync` により自動的に読み込まれます。`load`</span><span class="sxs-lookup"><span data-stu-id="810c4-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="810c4-223">詳細については、「[\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="810c4-223">For more information see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)</span></span>
- <span data-ttu-id="810c4-224">最初に数式のセルを含まない範囲を選択して実行することで、このコードをテストできます。</span><span class="sxs-lookup"><span data-stu-id="810c4-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="810c4-225">次に、少なくとも 1 つのセルに数式がある範囲を選択し、それをもう一度実行します。</span><span class="sxs-lookup"><span data-stu-id="810c4-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="810c4-226">この資料の他のすべての例をわかりやすくするためには、 `getSpecialCells` メソッドを `getSpecialCellsOrNullObject`の代わりに使用します。</span><span class="sxs-lookup"><span data-stu-id="810c4-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="810c4-227">セルの値の型で対象セルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="810c4-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="810c4-228">列挙型の省略可能な 2 番目のパラメーターがあり、対象とするセルをさらに絞り込みます。`Excel.SpecialCellValueType`</span><span class="sxs-lookup"><span data-stu-id="810c4-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="810c4-229">「数式」または「定数」のいずれかを渡す場合にのみに `getSpecialCells` または `getSpecialCellsOrNullObject`を使用できます。</span><span class="sxs-lookup"><span data-stu-id="810c4-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="810c4-230">パラメータは、必要な特定の種類の値のセルのみを指定します。</span><span class="sxs-lookup"><span data-stu-id="810c4-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="810c4-231">4 つの基本的な種類があります:「エラー」、「論理」(つまり、ブール値)、「番号」、および「テキスト」です。</span><span class="sxs-lookup"><span data-stu-id="810c4-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="810c4-232">(列挙型は、これら以外の他の値について次に説明する 4 つの他の値です。)次に、例を示します。</span><span class="sxs-lookup"><span data-stu-id="810c4-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="810c4-233">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="810c4-233">About this code, note:</span></span>

- <span data-ttu-id="810c4-234">リテラルの数値のあるセルのみがハイライト表示されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="810c4-235">数式(結果は数値の場合でも)またはブール値、文字列、またはエラーの状態のセルが強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="810c4-236">コードをテストするには、リテラルの数値、他の種類のリテラル値、一部の数式のセルがワークシートにあることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="810c4-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="810c4-237">場合によって 、すべてのテキスト値を持ち、すべてのブール値 (「論理」) を持つセルのよう1 つ以上のセル値型を操作する必要があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="810c4-238">列挙型には値の種類を組み合わせることができます。`Excel.SpecialCellValueType`</span><span class="sxs-lookup"><span data-stu-id="810c4-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="810c4-239">たとえば、"LogicalText"は、すべてのブール値、テキスト値を持つセルを対象にします。</span><span class="sxs-lookup"><span data-stu-id="810c4-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="810c4-240">4 つの基本的なタイプの内 2 つまたは 3 つを組み合わせることができます。</span><span class="sxs-lookup"><span data-stu-id="810c4-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="810c4-241">基本的な種類を組み合わせているこれらの列挙値の名前は、常にアルファベット順にします。</span><span class="sxs-lookup"><span data-stu-id="810c4-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="810c4-242">セルのエラー値、テキスト値、およびブール値を結合するには、"LogicalErrorText"または"TextErrorLogical"ではなく"ErrorLogicalText"を使用します。</span><span class="sxs-lookup"><span data-stu-id="810c4-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="810c4-243">「すべて」の既定のパラメーターは、4 種類全てを結合します。</span><span class="sxs-lookup"><span data-stu-id="810c4-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="810c4-244">次の使用例では、数値またはブール値を生成する数式を持つすべてのセルを強調表示しています。</span><span class="sxs-lookup"><span data-stu-id="810c4-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="810c4-245">`Excel.SpecialCellValueType` パラメーターは、 `Excel.SpecialCellType` パラメーターが "数式" または "定数" である場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="810c4-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="810c4-246">RangeAreas 内の RangeAreas を取得します。</span><span class="sxs-lookup"><span data-stu-id="810c4-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="810c4-247">型自体も、同じの 2 つのパラメーターを受け取る `getSpecialCells` と `getSpecialCellsOrNullObject` メソッドを持ちます。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="810c4-248">これらのメソッドは、 `RangeAreas.areas` コレクション内のすべての範囲のセル範囲からすべての対象となるセルを返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="810c4-249">オブジェクト の代わりに オブジェクトを呼び出した場合のメソッドの動作のには1 つの小さな違いがあります :"SameConditionalFormat"を最初のパラメーターとして渡すと、メソッドは、  コレクション 内の最初の範囲の上の左端のセル と同じ条件付き書式を持つすべてのセルを返します。`RangeAreas``Range`*`RangeAreas.areas`*</span><span class="sxs-lookup"><span data-stu-id="810c4-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="810c4-250">同じポイントは、"SameDataValidation"に適用されます:  に渡される場合、 範囲内の  一番左上と同じデータの入力規則をもつすべてのセルを返します。`Range.getSpecialCells`\*\*</span><span class="sxs-lookup"><span data-stu-id="810c4-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="810c4-251">しかｈしに渡された売位、  コレクション内の 最初の範囲の左端のセルと同じデータの入力規則が持つすべてのセルが返されます。`RangeAreas.getSpecialCells`*`RangeAreas.areas`*</span><span class="sxs-lookup"><span data-stu-id="810c4-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="810c4-252">RangeAreas のプロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="810c4-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="810c4-253">プロパティ値を読み取るときは、指定したプロパティが、 `RangeAreas`内の別の範囲内の異なる値を持つ可能性があるため、注意が必要です。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="810c4-254">一般的な規則は、一貫性のある値 *が* 返される場合、それが返されることです。</span><span class="sxs-lookup"><span data-stu-id="810c4-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="810c4-255">例えば、次のコードでは、ピンク色の (`#FFC0CB`) と `true` 用のRGBコードは、 `RangeAreas` オブジェクトの両方の範囲がピンク色の塗りであり、両方が全体の列であるため、コンソールに格納されます。</span><span class="sxs-lookup"><span data-stu-id="810c4-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="810c4-256">整合性が可能でない場合、より複雑になります。</span><span class="sxs-lookup"><span data-stu-id="810c4-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="810c4-257">プロパティの動作は、次の 3 つの原則に従います。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="810c4-258">オブジェクトのブール型のプロパティは、すべてのメンバーの範囲が true でない限り、 `false` を返します。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="810c4-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="810c4-259">非ブール値プロパティは、 `address` プロパティ例外を除いて、全てのメンバーの範囲に対応するプロパティが同じ値を持っていない限り、 `null` を返します。</span><span class="sxs-lookup"><span data-stu-id="810c4-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="810c4-260">プロパティは、メンバーの範囲のアドレスのコンマ区切りの文字列を返します。`address`</span><span class="sxs-lookup"><span data-stu-id="810c4-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="810c4-261">たとえば、次のコードは、1 つだけ列全体であり、 1 つだけがピンク色で塗りつぶされます `RangeAreas` を生成します。</span><span class="sxs-lookup"><span data-stu-id="810c4-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="810c4-262">コンソールが、塗りつぶしの色に `null` を、 `false` を `isEntireRow` プロパティに、および"Sheet1!F3:F5、Sheet1!H:H"("Sheet1"は、シート名と仮定した場合) を `address` プロパティに表示します。</span><span class="sxs-lookup"><span data-stu-id="810c4-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="810c4-263">関連項目</span><span class="sxs-lookup"><span data-stu-id="810c4-263">See also</span></span>

- [<span data-ttu-id="810c4-264">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="810c4-264">Excel JavaScript API core concepts</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="810c4-265">Range オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="810c4-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="810c4-266">[RangeAreas オブジェクト (EXCELL用JavaScript API )](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (API がプレビュー中の場合、このリンクが動作しない 可能性があります。</span><span class="sxs-lookup"><span data-stu-id="810c4-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="810c4-267">代わりに、 [ベータ版 office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="810c4-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>