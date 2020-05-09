---
title: Excel JavaScript API を使用して範囲を操作する (高度)
description: 特殊なセル、重複の削除、日付の操作など、高度な範囲のオブジェクトの関数とシナリオ。
ms.date: 05/06/2020
localization_priority: Normal
ms.openlocfilehash: eb25ae3f4bbe1231cfdf49f7535490b39c7a419e
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170815"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="25c1b-103">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="25c1b-103">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="25c1b-104">この記事は、「[Excel JavaScript API を使用して範囲を操作する (基本)](excel-add-ins-ranges.md)」の情報に基づいており、コード サンプルでは Excel JavaScript API を使って範囲のより高度なタスクを実行する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-104">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="25c1b-105">オブジェクトが`Range`サポートするプロパティとメソッドの完全な一覧については、「 [Range オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-105">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="25c1b-106">Moment-MSDate プラグインを使用した日付の操作</span><span class="sxs-lookup"><span data-stu-id="25c1b-106">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="25c1b-107">[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="25c1b-108">[Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="25c1b-109">これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="25c1b-110">次のコードは、範囲 **B4** に時刻のタイムスタンプを設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-110">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="25c1b-111">これは、次の例に示すように、セルから日付を取得して、その日付を時刻などの形式に変換するのと同様の手法です。</span><span class="sxs-lookup"><span data-stu-id="25c1b-111">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="25c1b-112">アドインでは、わかりやすい形式で日付が表示されるように、範囲の書式を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-112">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="25c1b-113">たとえば、`"[$-409]m/d/yy h:mm AM/PM;@"` では時刻が "12/3/18 3:57 PM" のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-113">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="25c1b-114">日付と時刻の数値書式の詳細については、「[表示形式のカスタマイズに関するガイドラインを確認する](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)」の記事で「日付と時刻の表示に関するガイドライン」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-114">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="25c1b-115">複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="25c1b-115">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="25c1b-116">[Rangeareas](/javascript/api/excel/excel.rangeareas)オブジェクトを使用すると、アドインで複数の範囲に対して一度に操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-116">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="25c1b-117">これらの範囲は、連続していても連続していなくても構いません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-117">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="25c1b-118">`RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-118">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="25c1b-119">範囲内の特殊なセルを検索する</span><span class="sxs-lookup"><span data-stu-id="25c1b-119">Find special cells within a range</span></span>

<span data-ttu-id="25c1b-120">[範囲](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)の[getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)メソッドは、セルの特性とセルの値の種類に基づいて範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-120">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="25c1b-121">これらのメソッドでは両方とも、`RangeAreas` オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-121">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="25c1b-122">次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-122">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="25c1b-123">次の例では、`getSpecialCells` メソッドを使用して、数式を含むすべてのセルを検索します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-123">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="25c1b-124">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-124">About this code, note:</span></span>

- <span data-ttu-id="25c1b-125">検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-125">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="25c1b-126">`getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-126">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="25c1b-127">対象の特性を含むセルが範囲内に存在しない場合、`getSpecialCells` によって **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-127">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="25c1b-128">この場合、制御のフローが `catch` ブロックに移ります (存在する場合)。</span><span class="sxs-lookup"><span data-stu-id="25c1b-128">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="25c1b-129">`catch`ブロックがない場合、エラーによってメソッドは停止します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-129">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="25c1b-130">対象の特性を含むセルが常に存在するはずである場合、そうしたセルが存在しないなら、コードを使ってエラーをスローする必要があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-130">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="25c1b-131">一致するセルがないということが有効なシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-131">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="25c1b-132">`getSpecialCellsOrNullObject` メソッドと、返された `isNullObject` プロパティを使用して、この動作を実現できます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-132">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="25c1b-133">次のサンプルでは、このパターンを使用しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-133">The following example uses this pattern.</span></span> <span data-ttu-id="25c1b-134">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-134">About this code, note:</span></span>

- <span data-ttu-id="25c1b-135">`getSpecialCellsOrNullObject` メソッドは常にプロキシ オブジェクトを返します。そのため、通常の JavaScript 使用環境では `null` となることはありません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-135">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="25c1b-136">ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-136">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="25c1b-137">`isNullObject` プロパティをテストする*前*に、`context.sync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-137">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="25c1b-138">これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-138">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="25c1b-139">ただし、*明示的*に `isNullObject` プロパティをロードする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-139">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="25c1b-140">`load` がオブジェクトに対して呼び出されていない場合であっても、プロパティは `context.sync` によって自動的にロードされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-140">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="25c1b-141">詳細については、「[\*OrNullObject メソッド](../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-141">For more information, see [\*OrNullObject](../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods).</span></span>
- <span data-ttu-id="25c1b-142">このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-142">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="25c1b-143">次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-143">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

<span data-ttu-id="25c1b-144">わかりやすくするため、この記事内のすべての他の例では、`getSpecialCells` メソッドを `getSpecialCellsOrNullObject` の代わりに使用しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-144">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="25c1b-145">セルの値の型に応じて対象のセルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="25c1b-145">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="25c1b-146">`Range.getSpecialCells()` メソッドと `Range.getSpecialCellsOrNullObject()` メソッドでは、対象セルをさらに絞り込むためにオプションとして使用される 2 番目のパラメーターを承諾します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-146">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="25c1b-147">この 2 番目のパラメーターは、特定の種類の値を含むセルのみを指定するために使用される `Excel.SpecialCellValueType` パラメーターです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-147">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="25c1b-148">`Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` が `Excel.SpecialCellType.formulas` または `Excel.SpecialCellType.constants` の場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-148">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="25c1b-149">単一のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="25c1b-149">Test for a single cell value type</span></span>

<span data-ttu-id="25c1b-150">`Excel.SpecialCellValueType` 列挙型には、次の 4 つの基本型があります (このセクションで後述する他の値の組み合わせに加えて)。</span><span class="sxs-lookup"><span data-stu-id="25c1b-150">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="25c1b-151">`Excel.SpecialCellValueType.logical` (ブール型)</span><span class="sxs-lookup"><span data-stu-id="25c1b-151">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="25c1b-152">次の例では、数値定数である特殊なセルを検索し、そのセルをピンク色にします。</span><span class="sxs-lookup"><span data-stu-id="25c1b-152">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="25c1b-153">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-153">About this code, note:</span></span>

- <span data-ttu-id="25c1b-154">リテラル数値を持つセルのみ強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-154">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="25c1b-155">数式 (結果が数字の場合であっても)、ブール値、テキストを持つセル、およびエラー状態にあるセルは強調表示されません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-155">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="25c1b-156">コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-156">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="25c1b-157">複数のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="25c1b-157">Test for multiple cell value types</span></span>

<span data-ttu-id="25c1b-158">テキスト値のセルすべてとブール値 (`Excel.SpecialCellValueType.logical`) のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-158">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="25c1b-159">`Excel.SpecialCellValueType` 列挙型には、結合された型の値があります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-159">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="25c1b-160">たとえば、`Excel.SpecialCellValueType.logicalText` は、すべてのブール値のセルとテキスト値のセルを対象としています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-160">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="25c1b-161">`Excel.SpecialCellValueType.all` は既定値であり、返されるセル値の型は制限されません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-161">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="25c1b-162">次の例では、結果が数値またはブール値となる数式を含むすべてのセルが色付けされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-162">The following example colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="cut-copy-and-paste"></a><span data-ttu-id="25c1b-163">切り取り、コピー、および貼り付け</span><span class="sxs-lookup"><span data-stu-id="25c1b-163">Cut, copy, and paste</span></span>

### <a name="copy-and-paste"></a><span data-ttu-id="25c1b-164">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="25c1b-164">Copy and paste</span></span>

<span data-ttu-id="25c1b-165">このメソッドは、Excel UI の**コピー**と**貼り付け**の操作をレプリケートします[。](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)</span><span class="sxs-lookup"><span data-stu-id="25c1b-165">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="25c1b-166">`copyFrom` が呼び出される範囲オブジェクトがコピー先になります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-166">The range object that `copyFrom` is called on is the destination.</span></span> <span data-ttu-id="25c1b-167">コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-167">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="25c1b-168">次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。</span><span class="sxs-lookup"><span data-stu-id="25c1b-168">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="25c1b-169">`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-169">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="25c1b-170">`copyType` では、ソースからコピー先にコピーされるデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-170">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="25c1b-171">`Excel.RangeCopyType.formulas`元のセルの数式を移動し、それらの数式の範囲の相対的な位置を保持します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-171">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="25c1b-172">任意の数式以外のエントリはそのままコピーされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-172">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="25c1b-173">`Excel.RangeCopyType.values` では、データ値と、数式の場合は数式の結果をコピーします。</span><span class="sxs-lookup"><span data-stu-id="25c1b-173">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="25c1b-174">`Excel.RangeCopyType.formats` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-174">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="25c1b-175">`Excel.RangeCopyType.all`(既定のオプション) を選択すると、データと書式設定の両方がコピーされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-175">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="25c1b-176">`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-176">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="25c1b-177">true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-177">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="25c1b-178">スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-178">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="25c1b-179">既定値は false です。</span><span class="sxs-lookup"><span data-stu-id="25c1b-179">The default is false.</span></span>

<span data-ttu-id="25c1b-180">`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-180">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="25c1b-181">行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-181">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="25c1b-182">次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-182">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="25c1b-183">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="25c1b-183">*Before the preceding function has been run.*</span></span>

![範囲のコピーメソッドが実行される前の Excel のデータ](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="25c1b-185">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="25c1b-185">*After the preceding function has been run.*</span></span>

![範囲のコピーメソッドが実行された後の Excel のデータ](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a><span data-ttu-id="25c1b-187">セルの切り取りと貼り付け (移動)</span><span class="sxs-lookup"><span data-stu-id="25c1b-187">Cut and paste (move) cells</span></span>

<span data-ttu-id="25c1b-188">[指定範囲の moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-)メソッドは、セルをブック内の新しい位置に移動します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-188">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="25c1b-189">このセルの移動動作は、セル[範囲をドラッグ](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e)してセルを移動した場合や、**切り取り**と**貼り付け**の操作を行った場合と同じです。</span><span class="sxs-lookup"><span data-stu-id="25c1b-189">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="25c1b-190">範囲の書式設定と値の両方が、 `destinationRange`パラメーターとして指定された場所に移動します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-190">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="25c1b-191">次のコードサンプルは、 `Range.moveTo`メソッドを使用して移動する範囲を示しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-191">The following code sample shows a range being moved with the `Range.moveTo` method.</span></span> <span data-ttu-id="25c1b-192">コピー先の範囲がソースよりも小さい場合は、ソースコンテンツを含むように展開されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-192">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a><span data-ttu-id="25c1b-193">重複の削除</span><span class="sxs-lookup"><span data-stu-id="25c1b-193">Remove duplicates</span></span>

<span data-ttu-id="25c1b-194">指定した列に重複するエントリがある行を削除するには、このメソッドを使用し[ます。](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)</span><span class="sxs-lookup"><span data-stu-id="25c1b-194">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="25c1b-195">このメソッドは、値が最小のインデックスから、範囲内の最大値のインデックス (上から下) までの範囲にある各行を処理します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-195">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="25c1b-196">任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-196">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="25c1b-197">範囲にある削除された行の下の行が上に移動します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-197">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="25c1b-198">`removeDuplicates` は、範囲外にあるセルの位置には影響しません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-198">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="25c1b-199">`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-199">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="25c1b-200">この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-200">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="25c1b-201">メソッドには、最初の行がヘッダーであるかどうかを指定するブール値のパラメーターもあります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-201">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="25c1b-202">**true** の場合、重複について考慮するとき最初の行は無視されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-202">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="25c1b-203">メソッド`removeDuplicates`は、削除`RemoveDuplicatesResult`された行数と、残っている一意の行の数を指定するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-203">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="25c1b-204">範囲の`removeDuplicates`メソッドを使用する場合は、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-204">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="25c1b-205">`removeDuplicates` は、関数の結果ではなくセルの値を考慮します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-205">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="25c1b-206">2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-206">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="25c1b-207">空のセルは、`removeDuplicates` に無視されることはありません。</span><span class="sxs-lookup"><span data-stu-id="25c1b-207">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="25c1b-208">空のセルの値は、その他の値と同様に扱われます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-208">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="25c1b-209">つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-209">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="25c1b-210">次の例では、最初の列に重複する値があるエントリを削除する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-210">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

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

<span data-ttu-id="25c1b-211">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="25c1b-211">*Before the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行される前の Excel のデータ](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="25c1b-213">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="25c1b-213">*After the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行された後の Excel のデータ](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a><span data-ttu-id="25c1b-215">アウトラインのデータをグループ化する</span><span class="sxs-lookup"><span data-stu-id="25c1b-215">Group data for an outline</span></span>

<span data-ttu-id="25c1b-216">行またはセル範囲の列は、[アウトライン](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)を作成するためにまとめてグループ化することができます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-216">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="25c1b-217">これらのグループを折りたたんで展開し、対応するセルを非表示にして表示することができます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-217">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="25c1b-218">これにより、トップ行のデータの簡単な分析が容易になります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-218">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="25c1b-219">これらのアウトライングループを作成するには、[範囲グループ](/javascript/api/excel/excel.range#group-groupoption-)を使用します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-219">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="25c1b-220">アウトラインには階層を設定できます。小さなグループは、より大きいグループの下にネストされています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-220">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="25c1b-221">これにより、アウトラインを異なるレベルで表示できるようになります。</span><span class="sxs-lookup"><span data-stu-id="25c1b-221">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="25c1b-222">表示されるアウトラインレベルを変更するには、 [showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)メソッドを使用してプログラムで実行できます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-222">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="25c1b-223">Excel では8レベルのアウトライングループのみがサポートされることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="25c1b-223">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="25c1b-224">次のコードサンプルでは、行と列の両方に対して2つのレベルのグループを持つアウトラインを作成する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-224">The following code sample shows how to create an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="25c1b-225">次の図は、そのアウトラインのグループを示しています。</span><span class="sxs-lookup"><span data-stu-id="25c1b-225">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="25c1b-226">コードサンプルでは、グループ化されている範囲に、アウトラインコントロールの行または列が含まれていないことに注意してください (この例の場合は "集計")。</span><span class="sxs-lookup"><span data-stu-id="25c1b-226">Note that in the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="25c1b-227">グループは、コントロールのある行または列ではなく、折りたたまれる内容を定義します。</span><span class="sxs-lookup"><span data-stu-id="25c1b-227">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![2レベルの2次元のアウトラインがある範囲](../images/excel-outline.png)

<span data-ttu-id="25c1b-229">行または列グループのグループを解除するには、グループ化を解除するメソッドを使用します[。](/javascript/api/excel/excel.range#ungroup-groupoption-)</span><span class="sxs-lookup"><span data-stu-id="25c1b-229">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="25c1b-230">これにより、アウトラインから最上位レベルが削除されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-230">This removes the outermost level from the outline.</span></span> <span data-ttu-id="25c1b-231">同じ行または列の種類の複数のグループが指定された範囲内の同じレベルにある場合、それらすべてのグループはグループ解除されます。</span><span class="sxs-lookup"><span data-stu-id="25c1b-231">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="25c1b-232">関連項目</span><span class="sxs-lookup"><span data-stu-id="25c1b-232">See also</span></span>

- [<span data-ttu-id="25c1b-233">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="25c1b-233">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="25c1b-234">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="25c1b-234">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="25c1b-235">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="25c1b-235">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
