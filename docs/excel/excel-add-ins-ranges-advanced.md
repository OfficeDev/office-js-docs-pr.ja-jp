---
title: Excel JavaScript API を使用して範囲を操作する (高度)
description: ''
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: c8fbe1dcc75080c932b4c3e2946fe62747d35c6b
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395597"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="1b98e-102">Excel JavaScript API を使用して範囲を操作する (高度)</span><span class="sxs-lookup"><span data-stu-id="1b98e-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="1b98e-103">この記事は、「[Excel JavaScript API を使用して範囲を操作する (基本)](excel-add-ins-ranges.md)」の情報に基づいており、コード サンプルでは Excel JavaScript API を使って範囲のより高度なタスクを実行する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="1b98e-104">**Range** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b98e-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="1b98e-105">Moment-MSDate プラグインを使用した日付の操作</span><span class="sxs-lookup"><span data-stu-id="1b98e-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="1b98e-106">[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="1b98e-107">[Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="1b98e-108">これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="1b98e-109">次のコードは、範囲 **B4** に時刻のタイムスタンプを設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="1b98e-110">これは、次の例に示すように、セルから日付を取得して、その日付を時刻などの形式に変換するのと同様の手法です。</span><span class="sxs-lookup"><span data-stu-id="1b98e-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="1b98e-111">アドインでは、わかりやすい形式で日付が表示されるように、範囲の書式を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="1b98e-112">たとえば、`"[$-409]m/d/yy h:mm AM/PM;@"` では時刻が "12/3/18 3:57 PM" のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="1b98e-113">日付と時刻の数値書式の詳細については、「[表示形式のカスタマイズに関するガイドラインを確認する](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)」の記事で「日付と時刻の表示に関するガイドライン」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b98e-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="1b98e-114">複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="1b98e-114">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="1b98e-115">`RangeAreas` オブジェクトを使用すると、アドインの操作を一度に複数の範囲で実行できます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-115">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="1b98e-116">これらの範囲は、連続していても連続していなくても構いません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-116">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="1b98e-117">`RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-117">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="1b98e-118">範囲内の特殊なセルを検索する</span><span class="sxs-lookup"><span data-stu-id="1b98e-118">Find special cells within a range</span></span>

<span data-ttu-id="1b98e-119">`Range.getSpecialCells()` メソッドと`Range.getSpecialCellsOrNullObject()` メソッドでは、対象セルの特性と対象セルの値の型に基づいて範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-119">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="1b98e-120">これらのメソッドでは両方とも、`RangeAreas` オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-120">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="1b98e-121">次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-121">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="1b98e-122">次の例では、`getSpecialCells` メソッドを使用して、数式を含むすべてのセルを検索します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-122">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="1b98e-123">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-123">About this code, note:</span></span>

- <span data-ttu-id="1b98e-124">検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-124">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="1b98e-125">`getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-125">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="1b98e-126">対象の特性を含むセルが範囲内に存在しない場合、`getSpecialCells` によって **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-126">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="1b98e-127">この場合、制御のフローが `catch` ブロックに移ります (存在する場合)。</span><span class="sxs-lookup"><span data-stu-id="1b98e-127">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="1b98e-128">`catch` ブロックが存在しない場合は、このエラーにより関数が停止します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-128">If there isn't a `catch` block, the error halts the function.</span></span>

<span data-ttu-id="1b98e-129">対象の特性を含むセルが常に存在するはずである場合、そうしたセルが存在しないなら、コードを使ってエラーをスローする必要があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-129">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="1b98e-130">一致するセルがないということが有効なシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-130">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="1b98e-131">`getSpecialCellsOrNullObject` メソッドと、返された `isNullObject` プロパティを使用して、この動作を実現できます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-131">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="1b98e-132">次のサンプルでは、このパターンを使用しています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-132">The following example uses this pattern.</span></span> <span data-ttu-id="1b98e-133">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-133">About this code, note:</span></span>

- <span data-ttu-id="1b98e-134">`getSpecialCellsOrNullObject` メソッドは常にプロキシ オブジェクトを返します。そのため、通常の JavaScript 使用環境では `null` となることはありません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-134">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="1b98e-135">ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-135">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="1b98e-136">`isNullObject` プロパティをテストする*前*に、`context.sync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-136">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="1b98e-137">これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-137">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="1b98e-138">ただし、*明示的*に `isNullObject` プロパティをロードする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-138">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="1b98e-139">`load` がオブジェクトに対して呼び出されていない場合であっても、プロパティは `context.sync` によって自動的にロードされます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-139">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="1b98e-140">詳細については、「[\*OrNullObject メソッド](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b98e-140">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="1b98e-141">このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-141">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="1b98e-142">次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-142">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="1b98e-143">わかりやすくするため、この記事内のすべての他の例では、`getSpecialCells` メソッドを `getSpecialCellsOrNullObject` の代わりに使用しています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-143">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="1b98e-144">セルの値の型に応じて対象のセルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="1b98e-144">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="1b98e-145">`Range.getSpecialCells()` メソッドと `Range.getSpecialCellsOrNullObject()` メソッドでは、対象セルをさらに絞り込むためにオプションとして使用される 2 番目のパラメーターを承諾します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-145">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="1b98e-146">この 2 番目のパラメーターは、特定の種類の値を含むセルのみを指定するために使用される `Excel.SpecialCellValueType` パラメーターです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-146">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="1b98e-147">`Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` が `Excel.SpecialCellType.formulas` または `Excel.SpecialCellType.constants` の場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-147">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="1b98e-148">単一のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="1b98e-148">Test for a single cell value type</span></span>

<span data-ttu-id="1b98e-149">`Excel.SpecialCellValueType` 列挙型には、次の 4 つの基本型があります (このセクションで後述する他の値の組み合わせに加えて)。</span><span class="sxs-lookup"><span data-stu-id="1b98e-149">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="1b98e-150">`Excel.SpecialCellValueType.logical` (ブール型)</span><span class="sxs-lookup"><span data-stu-id="1b98e-150">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="1b98e-151">次の例では、数値定数である特殊なセルを検索し、そのセルをピンク色にします。</span><span class="sxs-lookup"><span data-stu-id="1b98e-151">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="1b98e-152">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b98e-152">About this code, note:</span></span>

- <span data-ttu-id="1b98e-153">リテラル数値を持つセルのみ強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-153">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="1b98e-154">数式 (結果が数字の場合であっても)、ブール値、テキストを持つセル、およびエラー状態にあるセルは強調表示されません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-154">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="1b98e-155">コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="1b98e-155">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="1b98e-156">複数のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="1b98e-156">Test for multiple cell value types</span></span>

<span data-ttu-id="1b98e-157">テキスト値のセルすべてとブール値 (`Excel.SpecialCellValueType.logical`) のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-157">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="1b98e-158">`Excel.SpecialCellValueType` 列挙型には、結合された型の値があります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-158">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="1b98e-159">たとえば、`Excel.SpecialCellValueType.logicalText` は、すべてのブール値のセルとテキスト値のセルを対象としています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-159">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="1b98e-160">`Excel.SpecialCellValueType.all` は既定値であり、返されるセル値の型は制限されません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-160">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="1b98e-161">次の例では、結果が数値またはブール値となる数式を含むすべてのセルが色付けされます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-161">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="copy-and-paste"></a><span data-ttu-id="1b98e-162">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="1b98e-162">Copy and paste</span></span>

<span data-ttu-id="1b98e-163">範囲の `copyFrom` 関数では、Excel UI のコピーと貼り付けの動作をレプリケートします。</span><span class="sxs-lookup"><span data-stu-id="1b98e-163">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="1b98e-164">`copyFrom` が呼び出される範囲オブジェクトがコピー先になります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-164">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="1b98e-165">コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-165">The source to be copied is passed as a range or a string address representing a range.</span></span>
<span data-ttu-id="1b98e-166">次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。</span><span class="sxs-lookup"><span data-stu-id="1b98e-166">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b98e-167">`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-167">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="1b98e-168">`copyType` では、ソースからコピー先にコピーされるデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-168">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="1b98e-169">`Excel.RangeCopyType.formulas` では、ソースのセルの数式が転送され、それらの数式の範囲の相対配置は保持されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-169">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="1b98e-170">任意の数式以外のエントリはそのままコピーされます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-170">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="1b98e-171">`Excel.RangeCopyType.values` では、データ値と、数式の場合は数式の結果をコピーします。</span><span class="sxs-lookup"><span data-stu-id="1b98e-171">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="1b98e-172">`Excel.RangeCopyType.formats` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-172">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="1b98e-173">`Excel.RangeCopyType.all` (既定のオプション) では、データと書式設定の両方がコピーされます。見つかった場合、セルの数式は保持されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-173">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="1b98e-174">`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-174">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="1b98e-175">true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-175">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="1b98e-176">スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-176">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="1b98e-177">既定値は false です。</span><span class="sxs-lookup"><span data-stu-id="1b98e-177">The default is false.</span></span>

<span data-ttu-id="1b98e-178">`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-178">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="1b98e-179">行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-179">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="1b98e-180">次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-180">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="1b98e-181">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="1b98e-181">*Before the preceding function has been run.*</span></span>

![範囲のコピー メソッドが実行される前の Excel のデータ](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="1b98e-183">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="1b98e-183">*After the preceding function has been run.*</span></span>

![範囲のコピー メソッドが実行された後の Excel のデータ](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="1b98e-185">重複の削除</span><span class="sxs-lookup"><span data-stu-id="1b98e-185">Remove duplicates</span></span>

<span data-ttu-id="1b98e-186">Range オブジェクトの `removeDuplicates` 関数は、指定された列で重複するエントリを持つ行を削除します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-186">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="1b98e-187">関数は、範囲の一番小さい値のインデックスから一番大きい値のインデックスへ向かって各行を移動します (上から下へ)。</span><span class="sxs-lookup"><span data-stu-id="1b98e-187">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="1b98e-188">任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-188">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="1b98e-189">範囲にある削除された行の下の行が上に移動します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-189">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="1b98e-190">`removeDuplicates` は、範囲外にあるセルの位置には影響しません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-190">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="1b98e-191">`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-191">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="1b98e-192">この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。</span><span class="sxs-lookup"><span data-stu-id="1b98e-192">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="1b98e-193">この関数は、最初の行がヘッダーかどうかを指定するブール値のパラメーターも受け取ります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-193">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="1b98e-194">**true** の場合、重複について考慮するとき最初の行は無視されます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-194">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="1b98e-195">`removeDuplicates` 関数は、削除する行の数と、残りの一意の行の数を指定する `RemoveDuplicatesResult` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-195">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="1b98e-196">範囲の `removeDuplicates` 関数を使う場合、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="1b98e-196">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="1b98e-197">`removeDuplicates` は、関数の結果ではなくセルの値を考慮します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-197">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="1b98e-198">2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-198">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="1b98e-199">空のセルは、`removeDuplicates` に無視されることはありません。</span><span class="sxs-lookup"><span data-stu-id="1b98e-199">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="1b98e-200">空のセルの値は、その他の値と同様に扱われます。</span><span class="sxs-lookup"><span data-stu-id="1b98e-200">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="1b98e-201">つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。</span><span class="sxs-lookup"><span data-stu-id="1b98e-201">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="1b98e-202">次の例では、最初の列に重複する値があるエントリを削除する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1b98e-202">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

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

<span data-ttu-id="1b98e-203">*前の関数が実行される前。*</span><span class="sxs-lookup"><span data-stu-id="1b98e-203">*Before the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行される前の Excel のデータ](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="1b98e-205">*前の関数が実行された後。*</span><span class="sxs-lookup"><span data-stu-id="1b98e-205">*After the preceding function has been run.*</span></span>

![範囲の重複を削除するメソッドが実行された後の Excel のデータ](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="1b98e-207">関連項目</span><span class="sxs-lookup"><span data-stu-id="1b98e-207">See also</span></span>

- [<span data-ttu-id="1b98e-208">Excel JavaScript API を使用して範囲を操作する</span><span class="sxs-lookup"><span data-stu-id="1b98e-208">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="1b98e-209">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="1b98e-209">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1b98e-210">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="1b98e-210">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
