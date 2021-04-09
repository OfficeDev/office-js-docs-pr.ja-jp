---
title: Excel JavaScript API を使用して重複を削除する
description: Excel JavaScript API を使用して重複を削除する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0a2a076398e15d1b3b9db963a85703782056c91e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652911"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a><span data-ttu-id="105f6-103">Excel JavaScript API を使用して重複を削除する</span><span class="sxs-lookup"><span data-stu-id="105f6-103">Remove duplicates using the Excel JavaScript API</span></span>

<span data-ttu-id="105f6-104">この記事では、Excel JavaScript API を使用して範囲内の重複するエントリを削除するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="105f6-104">This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API.</span></span> <span data-ttu-id="105f6-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="105f6-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="remove-rows-with-duplicate-entries"></a><span data-ttu-id="105f6-106">重複するエントリがある行を削除する</span><span class="sxs-lookup"><span data-stu-id="105f6-106">Remove rows with duplicate entries</span></span>

<span data-ttu-id="105f6-107">[Range.removeDuplicates メソッド](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)は、指定した列に重複するエントリがある行を削除します。</span><span class="sxs-lookup"><span data-stu-id="105f6-107">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="105f6-108">メソッドは、最も低い値のインデックスから範囲の最も高い値のインデックス (上から下) の範囲の各行を通過します。</span><span class="sxs-lookup"><span data-stu-id="105f6-108">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="105f6-109">任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。</span><span class="sxs-lookup"><span data-stu-id="105f6-109">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="105f6-110">範囲にある削除された行の下の行が上に移動します。</span><span class="sxs-lookup"><span data-stu-id="105f6-110">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="105f6-111">`removeDuplicates` は、範囲外にあるセルの位置には影響しません。</span><span class="sxs-lookup"><span data-stu-id="105f6-111">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="105f6-112">`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="105f6-112">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="105f6-113">この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。</span><span class="sxs-lookup"><span data-stu-id="105f6-113">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="105f6-114">このメソッドは、最初の行がヘッダーであるかどうかを指定するブール型パラメーターも取ります。</span><span class="sxs-lookup"><span data-stu-id="105f6-114">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="105f6-115">**true** の場合、重複について考慮するとき最初の行は無視されます。</span><span class="sxs-lookup"><span data-stu-id="105f6-115">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="105f6-116">このメソッドは、削除された行の数と残りの一意の行数を指定する `removeDuplicates` `RemoveDuplicatesResult` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="105f6-116">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="105f6-117">範囲のメソッドを使用する場合 `removeDuplicates` は、次の注意が必要です。</span><span class="sxs-lookup"><span data-stu-id="105f6-117">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="105f6-118">`removeDuplicates` は、関数の結果ではなくセルの値を考慮します。</span><span class="sxs-lookup"><span data-stu-id="105f6-118">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="105f6-119">2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。</span><span class="sxs-lookup"><span data-stu-id="105f6-119">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="105f6-120">空のセルは、`removeDuplicates` に無視されることはありません。</span><span class="sxs-lookup"><span data-stu-id="105f6-120">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="105f6-121">空のセルの値は、その他の値と同様に扱われます。</span><span class="sxs-lookup"><span data-stu-id="105f6-121">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="105f6-122">つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。</span><span class="sxs-lookup"><span data-stu-id="105f6-122">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="105f6-123">次のコード サンプルは、最初の列に重複する値を持つエントリの削除を示しています。</span><span class="sxs-lookup"><span data-stu-id="105f6-123">The following code sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a><span data-ttu-id="105f6-124">重複するエントリが削除される前のデータ</span><span class="sxs-lookup"><span data-stu-id="105f6-124">Data before duplicate entries are removed</span></span>

![範囲の remove duplicates メソッドが実行される前の Excel のデータ](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a><span data-ttu-id="105f6-126">重複するエントリが削除された後のデータ</span><span class="sxs-lookup"><span data-stu-id="105f6-126">Data after duplicate entries are removed</span></span>

![範囲の削除重複メソッドが実行された後の Excel のデータ](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="105f6-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="105f6-128">See also</span></span>

- [<span data-ttu-id="105f6-129">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="105f6-129">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="105f6-130">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="105f6-130">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="105f6-131">Excel JavaScript API を使用して範囲を切り取り、コピー、貼り付ける</span><span class="sxs-lookup"><span data-stu-id="105f6-131">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-cut-copy-paste.md)
- [<span data-ttu-id="105f6-132">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="105f6-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
