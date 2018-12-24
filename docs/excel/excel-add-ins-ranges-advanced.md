---
title: Excel JavaScript API を使用して範囲を操作する (高度)
description: ''
ms.date: 12/18/2018
ms.openlocfilehash: 6d3da1e7eff4e61ae1b88213d0b432581d8f6a8a
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383240"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="d97e9-102">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="d97e9-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="d97e9-103">この記事は、「[Excel JavaScript API を使用して範囲を操作する (基本)](excel-add-ins-ranges.md)」の情報に基づいており、コード サンプルでは Excel JavaScript API を使って範囲のより高度なタスクを実行する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="d97e9-104">**Range** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="d97e9-105">Moment-MSDate プラグインを使用した日付の操作</span><span class="sxs-lookup"><span data-stu-id="d97e9-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="d97e9-106">[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="d97e9-107">[Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="d97e9-108">これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。</span><span class="sxs-lookup"><span data-stu-id="d97e9-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="d97e9-109">次のコードは、範囲 **B4** に時刻のタイムスタンプを設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d97e9-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d97e9-110">これは、次の例に示すように、セルから日付を取得して、その日付を時刻などの形式に変換するのと同様の手法です。</span><span class="sxs-lookup"><span data-stu-id="d97e9-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d97e9-111">アドインでは、わかりやすい形式で日付が表示されるように、範囲の書式を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="d97e9-112">たとえば、`"[$-409]m/d/yy h:mm AM/PM;@"` では時刻が "12/3/18 3:57 PM" のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="d97e9-113">日付と時刻の数値書式の詳細については、「[表示形式のカスタマイズに関するガイドラインを確認する](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)」の記事で「日付と時刻の表示に関するガイドライン」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously-preview"></a><span data-ttu-id="d97e9-114">複数の範囲を同時に操作する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d97e9-114">Work with multiple ranges simultaneously in Excel add-ins</span></span>

> [!NOTE]
> <span data-ttu-id="d97e9-115">現在、`RangeAreas` オブジェクトは、パブリック プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-115">The `RangeAreas` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="d97e9-116">この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="d97e9-116">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="d97e9-117">TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-117">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="d97e9-118">`RangeAreas` オブジェクトを使用すると、アドインの操作を一度に複数の範囲で実行できます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-118">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="d97e9-119">これらの範囲は、連続していても連続していなくても構いません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-119">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="d97e9-120">`RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-120">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="copy-and-paste-preview"></a><span data-ttu-id="d97e9-121">コピーと貼り付け (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d97e9-121">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="d97e9-122">現在、`Range.copyFrom` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-122">The `Range.copyFrom` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="d97e9-123">この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="d97e9-123">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="d97e9-124">TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-124">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="d97e9-125">範囲の `copyFrom` 関数では、Excel UI のコピーと貼り付けの動作をレプリケートします。</span><span class="sxs-lookup"><span data-stu-id="d97e9-125">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="d97e9-126">`copyFrom` が呼び出される範囲オブジェクトがコピー先になります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-126">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="d97e9-127">コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-127">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="d97e9-128">次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。</span><span class="sxs-lookup"><span data-stu-id="d97e9-128">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d97e9-129">`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-129">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="d97e9-130">`copyType` では、ソースからコピー先にコピーされるデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-130">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="d97e9-131">`"Formulas"` では、ソースのセルの数式が転送され、それらの数式の範囲の相対配置は保持されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-131">`"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="d97e9-132">任意の数式以外のエントリはそのままコピーされます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-132">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="d97e9-133">`"Values"` では、データ値と、数式の場合は数式の結果をコピーします。</span><span class="sxs-lookup"><span data-stu-id="d97e9-133">`"Values"` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="d97e9-134">`"Formats"` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-134">`"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="d97e9-135">`"All"` (既定のオプション) では、データと書式設定の両方がコピーされます。見つかった場合、セルの数式は保持されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-135">`"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="d97e9-136">`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-136">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="d97e9-137">true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-137">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="d97e9-138">スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-138">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="d97e9-139">既定値は false です。</span><span class="sxs-lookup"><span data-stu-id="d97e9-139">The default is false.</span></span>

<span data-ttu-id="d97e9-140">`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-140">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="d97e9-141">行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-141">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="d97e9-142">次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。</span><span class="sxs-lookup"><span data-stu-id="d97e9-142">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d97e9-143">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="d97e9-143">*Before the preceding function has been run.*</span></span>

![範囲のコピー メソッドが実行される前の Excel のデータ](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="d97e9-145">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="d97e9-145">*After the preceding function has been run.*</span></span>

![範囲のコピー メソッドが実行された後の Excel のデータ](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a><span data-ttu-id="d97e9-147">重複を削除 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="d97e9-147">Remove duplicates (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="d97e9-148">現在、Range オブジェクトの `removeDuplicates` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-148">The Range object's `removeDuplicates` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="d97e9-149">この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="d97e9-149">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="d97e9-150">TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-150">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="d97e9-151">Range オブジェクトの `removeDuplicates` 関数は、指定された列で重複するエントリを持つ行を削除します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-151">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="d97e9-152">関数は、範囲の一番小さい値のインデックスから一番大きい値のインデックスへ向かって各行を移動します (上から下へ)。</span><span class="sxs-lookup"><span data-stu-id="d97e9-152">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="d97e9-153">任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-153">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="d97e9-154">範囲にある削除された行の下の行が上に移動します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-154">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="d97e9-155">`removeDuplicates` は、範囲外にあるセルの位置には影響しません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-155">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="d97e9-156">`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-156">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="d97e9-157">この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。</span><span class="sxs-lookup"><span data-stu-id="d97e9-157">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="d97e9-158">この関数は、最初の行がヘッダーかどうかを指定するブール値のパラメーターも受け取ります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-158">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="d97e9-159">**true** の場合、重複について考慮するとき最初の行は無視されます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-159">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="d97e9-160">`removeDuplicates` 関数は、削除する行の数と、残りの一意の行の数を指定する `RemoveDuplicatesResult` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-160">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="d97e9-161">範囲の `removeDuplicates` 関数を使う場合、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="d97e9-161">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="d97e9-162">`removeDuplicates` は、関数の結果ではなくセルの値を考慮します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-162">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="d97e9-163">2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-163">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="d97e9-164">空のセルは、`removeDuplicates` に無視されることはありません。</span><span class="sxs-lookup"><span data-stu-id="d97e9-164">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="d97e9-165">空のセルの値は、その他の値と同様に扱われます。</span><span class="sxs-lookup"><span data-stu-id="d97e9-165">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="d97e9-166">つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。</span><span class="sxs-lookup"><span data-stu-id="d97e9-166">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="d97e9-167">次の例では、最初の列に重複する値があるエントリを削除する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d97e9-167">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
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

<span data-ttu-id="d97e9-168">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="d97e9-168">*Before the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行される前の Excel のデータ](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="d97e9-170">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="d97e9-170">*After the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行された後の Excel のデータ](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="d97e9-172">関連項目</span><span class="sxs-lookup"><span data-stu-id="d97e9-172">See also</span></span>

- [<span data-ttu-id="d97e9-173">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="d97e9-173">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="d97e9-174">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="d97e9-174">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)