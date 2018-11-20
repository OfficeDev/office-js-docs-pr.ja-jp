---
title: Excel JavaScript API のパフォーマンスの最適化
description: Excel JavaScript API を使用してパフォーマンスを最適化する
ms.date: 03/28/2018
ms.openlocfilehash: ee1687fcb1a5db74e65f5e73994653df235b4823
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505378"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="9c2ed-103">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="9c2ed-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="9c2ed-104">Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="9c2ed-105">さまざまなアプローチの間でパフォーマンスは大きく異なります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="9c2ed-106">この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="9c2ed-107">sync() 呼び出しの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="9c2ed-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="9c2ed-108">Excel JavaScript API では、```sync()``` は唯一の非同期操作であり、状況によっては遅くなる可能性があり、Excel Online の場合は特にその傾向があります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="9c2ed-109">パフォーマンスを最適化するには、```sync()``` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="9c2ed-110">このプラクティスに従うコード サンプルについては 「[Core Concepts - sync()](excel-add-ins-core-concepts.md#sync)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="9c2ed-111">作成されたプロキシ オブジェクトの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="9c2ed-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="9c2ed-112">同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="9c2ed-113">代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```javascript
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

## <a name="load-necessary-properties-only"></a><span data-ttu-id="9c2ed-114">必要なプロパティのみをロードする</span><span class="sxs-lookup"><span data-stu-id="9c2ed-114">Load necessary properties only</span></span>

<span data-ttu-id="9c2ed-115">Excel JavaScript API では、プロキシ オブジェクトのプロパティを明示的にロードする必要があります。 </span><span class="sxs-lookup"><span data-stu-id="9c2ed-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="9c2ed-116">空の ```load()``` 呼び出しで、すべてのプロパティを一度にロードすることはできますが、そのアプローチは大きなパフォーマンス オーバーヘッドを持つ可能性があります。 </span><span class="sxs-lookup"><span data-stu-id="9c2ed-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="9c2ed-117">代わりに、必要なプロパティだけをロードすることをお勧めします。特に、多数のプロパティを持つオブジェクトの場合はそうして下さい。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="9c2ed-118">たとえば、範囲オブジェクトの **address** プロパティのみを読み取る場合、**load()** メソッドを呼び出すときにそのプロパティのみを指定します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="9c2ed-119">**load()** メソッドは、次のいずれかの方法で呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="9c2ed-120">_構文:_</span><span class="sxs-lookup"><span data-stu-id="9c2ed-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="9c2ed-121">_各部分の意味は次のとおりです。_</span><span class="sxs-lookup"><span data-stu-id="9c2ed-121">_Where:_</span></span>
 
* <span data-ttu-id="9c2ed-122">`properties` は、ロードするプロパティの一覧で、コンマ区切りの文字列または名前の配列として指定されます。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="9c2ed-123">詳細については、「[Excel JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)」でオブジェクトに対して定義されている **load()** メソッドを参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="9c2ed-p106">`loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="9c2ed-126">オブジェクトの下の「プロパティ」の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="9c2ed-127">例えば、`format` は範囲オブジェクトの下のプロパティですが、`format` それ自体もオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="9c2ed-128">そのため、`range.load("format")` のような呼び出しをすると、これは以前に概説したように、パフォーマンスの問題を引き起こす可能性のある空の load() 呼び出しである `range.format.load()` に相当します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="9c2ed-129">これを避けるには、オブジェクト ツリー内の "リーフノード" のみをロードするようにしてください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="9c2ed-130">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="9c2ed-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="9c2ed-131">大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の ```context.sync()``` が呼び出されまで計算を中断することをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="9c2ed-132">非常に便利な方法で計算を中断し、再起動するための ```suspendApiCalculationUntilNextSync()``` API の使用方法は 「[Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-132">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="9c2ed-133">次のコードは、計算を一時的に中断する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);
    
    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="9c2ed-134">範囲内のすべてのセルの更新</span><span class="sxs-lookup"><span data-stu-id="9c2ed-134">Update all cells in a range</span></span> 

<span data-ttu-id="9c2ed-135">範囲内のすべてのセルを同じ値またはプロパティで更新する必要がある場合は、同じ値を繰り返し指定する 2 次元配列で行うと、更新が遅くなる可能性があります。このアプローチだと、範囲内のすべてのセルを Excel が反復しなければ、それぞれ個別に設定できないからです。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="9c2ed-136">Excelには、範囲内のすべてのセルを同じ値またはプロパティで更新するより効率的な方法が備わっています。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="9c2ed-137">セルの範囲に同じ値、同じ形式または同次数式を適用する必要がある場合は、配列の値の代わりに 1 つの値を指定する方が効率的です。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="9c2ed-138">そうすることで、パフォーマンスが大幅に向上します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="9c2ed-139">このアプローチが実際に動作していることを示すコード サンプルについては、「[コアの概念 - 範囲内のすべてのセルを更新](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="9c2ed-140">このアプローチが使える一般的なシナリオは、ワークシートの異なる列に異なる数値書式を設定する場合です。 </span><span class="sxs-lookup"><span data-stu-id="9c2ed-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="9c2ed-141">この場合、列を通って反復し、各列の数値書式を単一の値で設定するだけです。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="9c2ed-142">「[範囲内のすべてのセルを更新する](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)」のコード サンプルにあるように、各列を範囲として扱います。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="9c2ed-143">TypeScript を使用している場合は、2 次元配列に 1 つの値を設定できないことを示すコンパイル エラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="9c2ed-144">その値*は*プロパティを取得しているときは 2 次元配列なので、エラーは避けられません。TypeScript では、異なるセッター対ゲッターの型は許可されません。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="9c2ed-145">しかし、簡単な回避策として、`as any` 接尾辞 (例: `range.values = "hello world" as any`) で値を設定する方法があります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="9c2ed-146">テーブルへのデータのインポート</span><span class="sxs-lookup"><span data-stu-id="9c2ed-146">Importing data into tables</span></span>

<span data-ttu-id="9c2ed-147">膨大な量のデータを直接 [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-147">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="9c2ed-148">新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="9c2ed-149">既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="9c2ed-150">このアプローチの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-150">Here is an example in JavaScript of this operation.</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> <span data-ttu-id="9c2ed-151">[Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="enable-and-disable-events"></a><span data-ttu-id="9c2ed-152">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="9c2ed-152">Enable and disable events</span></span>

<span data-ttu-id="9c2ed-153">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-153">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="9c2ed-154">イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。</span><span class="sxs-lookup"><span data-stu-id="9c2ed-154">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="9c2ed-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="9c2ed-155">See also</span></span>

- [<span data-ttu-id="9c2ed-156">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="9c2ed-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9c2ed-157">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="9c2ed-157">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="9c2ed-158">Excel JavaScript API オープン仕様</span><span class="sxs-lookup"><span data-stu-id="9c2ed-158">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="9c2ed-159">ワークシート関数のオブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="9c2ed-159">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
