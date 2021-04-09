---
title: Excel JavaScript API を使用して範囲内の特別なセルを検索する
description: Excel JavaScript API を使用して、数式、エラー、数値を含むセルなどの特別なセルを検索する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6504873bcd8ab50bd4c03fe4f54b71d0bd920c5b
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652892"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="ba50d-103">Excel JavaScript API を使用して範囲内の特別なセルを検索する</span><span class="sxs-lookup"><span data-stu-id="ba50d-103">Find special cells within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="ba50d-104">この記事では、Excel JavaScript API を使用して範囲内の特殊なセルを検索するコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-104">This article provides code samples that find special cells within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="ba50d-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel.Range クラス」を参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="ba50d-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="find-ranges-with-special-cells"></a><span data-ttu-id="ba50d-106">特殊なセルを含む範囲を検索する</span><span class="sxs-lookup"><span data-stu-id="ba50d-106">Find ranges with special cells</span></span>

<span data-ttu-id="ba50d-107">[Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)メソッドと[Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)メソッドは、セルの特性とセルの値の種類に基づいて範囲を検索します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-107">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="ba50d-108">これらのメソッドでは両方とも、`RangeAreas` オブジェクトが返されます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-108">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="ba50d-109">次に示すのは、TypeScript データ型ファイルの、このメソッドのシグネチャです。</span><span class="sxs-lookup"><span data-stu-id="ba50d-109">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="ba50d-110">次のコード サンプルでは、メソッド `getSpecialCells` を使用して数式を含むすべてのセルを検索します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-110">The following code sample uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="ba50d-111">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ba50d-111">About this code, note:</span></span>

- <span data-ttu-id="ba50d-112">検索が必要なシートの部分を制限するために、まず `Worksheet.getUsedRange` を呼び出し、その範囲に関してのみ `getSpecialCells` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-112">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="ba50d-113">`getSpecialCells` メソッドは `RangeAreas` オブジェクトを返すため、数式を含むセルはすべて、連続していないセルであっても、ピンク色になります。</span><span class="sxs-lookup"><span data-stu-id="ba50d-113">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="ba50d-114">対象の特性を含むセルが範囲内に存在しない場合、`getSpecialCells` によって **ItemNotFound** エラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-114">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="ba50d-115">この場合、制御のフローが `catch` ブロックに移ります (存在する場合)。</span><span class="sxs-lookup"><span data-stu-id="ba50d-115">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="ba50d-116">ブロックが見当たらない `catch` 場合は、メソッドが停止します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-116">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="ba50d-117">対象の特性を含むセルが常に存在するはずである場合、そうしたセルが存在しないなら、コードを使ってエラーをスローする必要があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="ba50d-117">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="ba50d-118">一致するセルがないということが有効なシナリオでは、コードでこのような可能性があるかどうかを確認し、あれば、エラーをスローせずに適切に処理するようにしておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="ba50d-118">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="ba50d-119">`getSpecialCellsOrNullObject` メソッドと、返された `isNullObject` プロパティを使用して、この動作を実現できます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-119">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="ba50d-120">次のコード サンプルでは、このパターンを使用します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-120">The following code sample uses this pattern.</span></span> <span data-ttu-id="ba50d-121">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ba50d-121">About this code, note:</span></span>

- <span data-ttu-id="ba50d-122">メソッド `getSpecialCellsOrNullObject` は常にプロキシ オブジェクトを返すので、通常の `null` JavaScript の意味では返す必要がありません。</span><span class="sxs-lookup"><span data-stu-id="ba50d-122">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="ba50d-123">ただし一致するセルが見つからなかった場合、オブジェクトの `isNullObject` プロパティは `true` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-123">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="ba50d-124">`isNullObject` プロパティをテストする *前* に、`context.sync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-124">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="ba50d-125">これは、すべての `*OrNullObject` メソッドとプロパティの必要条件です。プロパティを読み取るためには常に、そのプロパティをロードして同期する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="ba50d-125">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="ba50d-126">ただし、プロパティを明示的に *読み込む* 必要 `isNullObject` はありません。</span><span class="sxs-lookup"><span data-stu-id="ba50d-126">However, it's not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="ba50d-127">オブジェクトで呼び出されない場合 `context.sync` でも `load` 、自動的に読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-127">It's automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="ba50d-128">詳細については[ \* 、「OrNullObject メソッドとプロパティ」を参照してください](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。</span><span class="sxs-lookup"><span data-stu-id="ba50d-128">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="ba50d-129">このコードをテストするには、最初に数式を含まないセルの範囲を選択してからコードを実行します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-129">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="ba50d-130">次に、少なくとも 1 つのセルが数式を含む範囲を選択してからコードを再実行します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-130">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="ba50d-131">わかりやすくするために、この記事の他のすべてのコード サンプルでは、 `getSpecialCells` の代わりにメソッドを使用します  `getSpecialCellsOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="ba50d-131">For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

## <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="ba50d-132">セルの値の型に応じて対象のセルを絞り込む</span><span class="sxs-lookup"><span data-stu-id="ba50d-132">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="ba50d-133">`Range.getSpecialCells()` メソッドと `Range.getSpecialCellsOrNullObject()` メソッドでは、対象セルをさらに絞り込むためにオプションとして使用される 2 番目のパラメーターを承諾します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-133">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="ba50d-134">この 2 番目のパラメーターは、特定の種類の値を含むセルのみを指定するために使用される `Excel.SpecialCellValueType` パラメーターです。</span><span class="sxs-lookup"><span data-stu-id="ba50d-134">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="ba50d-135">`Excel.SpecialCellValueType` パラメーターは、`Excel.SpecialCellType` が `Excel.SpecialCellType.formulas` または `Excel.SpecialCellType.constants` の場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-135">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="ba50d-136">単一のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="ba50d-136">Test for a single cell value type</span></span>

<span data-ttu-id="ba50d-137">`Excel.SpecialCellValueType` 列挙型には、次の 4 つの基本型があります (このセクションで後述する他の値の組み合わせに加えて)。</span><span class="sxs-lookup"><span data-stu-id="ba50d-137">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="ba50d-138">`Excel.SpecialCellValueType.logical` (ブール型)</span><span class="sxs-lookup"><span data-stu-id="ba50d-138">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="ba50d-139">次のコード サンプルでは、数値定数である特殊なセルを検索し、それらのセルをピンク色で色付けします。</span><span class="sxs-lookup"><span data-stu-id="ba50d-139">The following code sample finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="ba50d-140">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="ba50d-140">About this code, note:</span></span>

- <span data-ttu-id="ba50d-141">リテラル数値を持つセルのみを強調表示します。</span><span class="sxs-lookup"><span data-stu-id="ba50d-141">It only highlights cells that have a literal number value.</span></span> <span data-ttu-id="ba50d-142">数式 (結果が数値の場合でも) またはブール値、テキスト、またはエラー状態のセルを持つセルは強調表示されます。</span><span class="sxs-lookup"><span data-stu-id="ba50d-142">It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="ba50d-143">コードをテストするには、リテラル数値を持ついくつかのセル、他の型のリテラル値を持ついくつかのセル、そして数式を持ついくつかのセルをそれぞれワークシートに含めるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="ba50d-143">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="ba50d-144">複数のセル値の型のテスト</span><span class="sxs-lookup"><span data-stu-id="ba50d-144">Test for multiple cell value types</span></span>

<span data-ttu-id="ba50d-145">テキスト値のセルすべてとブール値 (`Excel.SpecialCellValueType.logical`) のセルすべてなど、セル値の型を複数操作する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="ba50d-145">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="ba50d-146">`Excel.SpecialCellValueType` 列挙型には、結合された型の値があります。</span><span class="sxs-lookup"><span data-stu-id="ba50d-146">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="ba50d-147">たとえば、`Excel.SpecialCellValueType.logicalText` は、すべてのブール値のセルとテキスト値のセルを対象としています。</span><span class="sxs-lookup"><span data-stu-id="ba50d-147">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="ba50d-148">`Excel.SpecialCellValueType.all` は既定値であり、返されるセル値の型は制限されません。</span><span class="sxs-lookup"><span data-stu-id="ba50d-148">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="ba50d-149">次のコード サンプルでは、数値またはブール値を生成する数式ですべてのセルを色付けします。</span><span class="sxs-lookup"><span data-stu-id="ba50d-149">The following code sample colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="ba50d-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="ba50d-150">See also</span></span>

- [<span data-ttu-id="ba50d-151">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="ba50d-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ba50d-152">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="ba50d-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="ba50d-153">Excel JavaScript API を使用して文字列を検索する</span><span class="sxs-lookup"><span data-stu-id="ba50d-153">Find a string using the Excel JavaScript API</span></span>](excel-add-ins-ranges-string-match.md)
- [<span data-ttu-id="ba50d-154">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="ba50d-154">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
