---
title: Excel JavaScript API のパフォーマンスの最適化
description: Excel JavaScript API を使用してパフォーマンスを最適化する
ms.date: 07/14/2020
localization_priority: Normal
ms.openlocfilehash: 193cbe8c8cd1a432c6567401ed645990cb93e5e9
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159095"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="69ada-103">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="69ada-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="69ada-104">Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="69ada-105">さまざまなアプローチの間でパフォーマンスは大きく異なります。</span><span class="sxs-lookup"><span data-stu-id="69ada-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="69ada-106">この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。</span><span class="sxs-lookup"><span data-stu-id="69ada-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="69ada-107">sync() 呼び出しの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="69ada-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="69ada-108">Excel JavaScript API では、`sync()` は唯一の非同期操作で、状況によっては遅くなる可能性があり、Excel on the web の場合は特にその傾向があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-108">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="69ada-109">パフォーマンスを最適化するには、`sync()` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="69ada-109">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="69ada-110">このプラクティスに従うコード サンプルについては 「[Core Concepts - sync()](excel-add-ins-core-concepts.md#sync)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="69ada-111">作成されたプロキシ オブジェクトの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="69ada-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="69ada-112">同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。</span><span class="sxs-lookup"><span data-stu-id="69ada-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="69ada-113">代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。</span><span class="sxs-lookup"><span data-stu-id="69ada-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="69ada-114">必要なプロパティのみをロードする</span><span class="sxs-lookup"><span data-stu-id="69ada-114">Load necessary properties only</span></span>

<span data-ttu-id="69ada-115">Excel JavaScript API では、プロキシ オブジェクトのプロパティを明示的にロードする必要があります。 </span><span class="sxs-lookup"><span data-stu-id="69ada-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="69ada-116">空の `load()` 呼び出しで、すべてのプロパティを一度にロードすることはできますが、そのアプローチは大きなパフォーマンス オーバーヘッドを持つ可能性があります。 </span><span class="sxs-lookup"><span data-stu-id="69ada-116">Although you're able to load all the properties at once with an empty `load()` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="69ada-117">代わりに、必要なプロパティだけをロードすることをお勧めします。特に、多数のプロパティを持つオブジェクトの場合はそうして下さい。</span><span class="sxs-lookup"><span data-stu-id="69ada-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="69ada-118">たとえば、range オブジェクトのプロパティのみを読み取る場合は、 `address` メソッドを呼び出すときにそのプロパティのみを指定し `load()` ます。</span><span class="sxs-lookup"><span data-stu-id="69ada-118">For example, if you only intend to read the `address` property of a range object, specify only that property when you call the `load()` method:</span></span>

```js
range.load('address');
```

<span data-ttu-id="69ada-119">`load()`メソッドは、次のいずれかの方法で呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="69ada-119">You can call `load()` method in any of the following ways:</span></span>

<span data-ttu-id="69ada-120">_構文:_</span><span class="sxs-lookup"><span data-stu-id="69ada-120">_Syntax:_</span></span>

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

<span data-ttu-id="69ada-121">_各部分の意味は次のとおりです。_</span><span class="sxs-lookup"><span data-stu-id="69ada-121">_Where:_</span></span>

* <span data-ttu-id="69ada-122">`properties` は、ロードするプロパティの一覧で、コンマ区切りの文字列または名前の配列として指定されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="69ada-123">詳細については、「 `load()` [EXCEL JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)」のオブジェクトに対して定義されているメソッドを参照してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-123">For more information, see the `load()` methods defined for objects in [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md).</span></span>
* <span data-ttu-id="69ada-p106">`loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](/javascript/api/office/officeextension.loadoption)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="69ada-126">オブジェクトの [プロパティ] の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-126">Please be aware that some of the "properties" under an object may have the same name as another object.</span></span> <span data-ttu-id="69ada-127">例えば、`format` は範囲オブジェクトの下のプロパティですが、`format` それ自体もオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="69ada-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="69ada-128">そのため、`range.load("format")` のような呼び出しをすると、これは以前に概説したように、パフォーマンスの問題を引き起こす可能性のある空の load() 呼び出しである `range.format.load()` に相当します。</span><span class="sxs-lookup"><span data-stu-id="69ada-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="69ada-129">これを回避するには、コードでオブジェクトツリーの "葉 nodes" のみを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-129">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="69ada-130">Excel のプロセスを一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="69ada-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="69ada-131">Excel には、ユーザーとアドインの両方からの入力に対応する多くのバックグラウンド タスクがあります。</span><span class="sxs-lookup"><span data-stu-id="69ada-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="69ada-132">これらの Excel のプロセスの一部は、パフォーマンス上の利点が得られるようにコントロールすることができます。</span><span class="sxs-lookup"><span data-stu-id="69ada-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="69ada-133">これは、アドインが大きなデータ セットを処理する場合に特に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="69ada-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="69ada-134">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="69ada-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="69ada-135">大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の `context.sync()` が呼び出されまで計算を中断することをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="69ada-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="69ada-136">非常に便利な方法で計算を中断し、再起動するための `suspendApiCalculationUntilNextSync()` API の使用方法については、「[Application Object](/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-136">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="69ada-137">次のコードは、計算を一時的に中断する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="69ada-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

<span data-ttu-id="69ada-138">数式の計算のみが中断されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="69ada-138">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="69ada-139">変更された参照はまだ再構築されています。</span><span class="sxs-lookup"><span data-stu-id="69ada-139">Any altered references are still rebuilt.</span></span> <span data-ttu-id="69ada-140">たとえば、ワークシートの名前を変更しても、そのワークシートへの数式の参照は更新されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-140">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="69ada-141">画面の更新を停止する</span><span class="sxs-lookup"><span data-stu-id="69ada-141">Suspend screen updating</span></span>

<span data-ttu-id="69ada-142">Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-142">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="69ada-143">大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。</span><span class="sxs-lookup"><span data-stu-id="69ada-143">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="69ada-144">`Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。</span><span class="sxs-lookup"><span data-stu-id="69ada-144">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="69ada-145">Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-145">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="69ada-146">繰り返し呼び出しない `suspendScreenUpdatingUntilNextSync` (ループの場合など)。</span><span class="sxs-lookup"><span data-stu-id="69ada-146">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="69ada-147">呼び出しが繰り返し行われると、Excel ウィンドウがちらつくようになります。</span><span class="sxs-lookup"><span data-stu-id="69ada-147">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="69ada-148">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="69ada-148">Enable and disable events</span></span>

<span data-ttu-id="69ada-149">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-149">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="69ada-150">イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。</span><span class="sxs-lookup"><span data-stu-id="69ada-150">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="69ada-151">テーブルへのデータのインポート</span><span class="sxs-lookup"><span data-stu-id="69ada-151">Importing data into tables</span></span>

<span data-ttu-id="69ada-152">膨大な量のデータを直接 [Table](/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="69ada-152">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="69ada-153">新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="69ada-153">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="69ada-154">既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-154">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="69ada-155">このアプローチの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="69ada-155">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="69ada-156">[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。</span><span class="sxs-lookup"><span data-stu-id="69ada-156">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="69ada-157">不要になった範囲の追跡解除</span><span class="sxs-lookup"><span data-stu-id="69ada-157">Untrack unneeded ranges</span></span>

<span data-ttu-id="69ada-158">JavaScript レイヤーは、アドインが Excel のブックと基になる範囲を操作するためのプロキシ オブジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="69ada-158">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="69ada-159">こうしたオブジェクトは、`context.sync()` が呼び出されるまでメモリに維持されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-159">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="69ada-160">大規模なバッチ操作では、アドインが 1 回のみ必要とするプロキシ オブジェクトが大量に生成されることがあります。それらのオブジェクトは、バッチの実行前にメモリから解放できます。</span><span class="sxs-lookup"><span data-stu-id="69ada-160">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="69ada-161">[Range.untrack()](/javascript/api/excel/excel.range#untrack--) メソッドにより、Excel の Range オブジェクトがメモリから解放されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-161">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="69ada-162">範囲に対してアドインを実行した後に、このメソッドを呼び出すと、大量の Range オブジェクトを使用しているときのパフォーマンスが大幅に向上します。</span><span class="sxs-lookup"><span data-stu-id="69ada-162">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="69ada-163">`Range.untrack()` は、[ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-) のショートカットです。</span><span class="sxs-lookup"><span data-stu-id="69ada-163">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="69ada-164">プロキシ オブジェクトは、コンテキスト内の追跡対象オブジェクト リストから削除することで追跡解除できます。</span><span class="sxs-lookup"><span data-stu-id="69ada-164">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="69ada-165">通常、Range オブジェクトは追跡の解除が正当化されるほどの量で使用される唯一の Excel オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="69ada-165">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="69ada-166">次のコード例では、指定した範囲に 1 セルずつデータを埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="69ada-166">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="69ada-167">セルに値が追加されると、そのセルを表している範囲の追跡が解除されます。</span><span class="sxs-lookup"><span data-stu-id="69ada-167">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="69ada-168">10,000 から 20,000 個のセルの範囲を選択して、このコードを実行します。最初の実行では `cell.untrack()` の行を使用し、その後でこの行を削除して実行します。</span><span class="sxs-lookup"><span data-stu-id="69ada-168">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="69ada-169">`cell.untrack()` の行がないコードよりも、この行があるコードの方が高速になることがわかります。</span><span class="sxs-lookup"><span data-stu-id="69ada-169">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="69ada-170">また、クリーンアップの手順にかかる時間が短くなるため、その後の応答時間も速くなることがわかります。</span><span class="sxs-lookup"><span data-stu-id="69ada-170">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="69ada-171">関連項目</span><span class="sxs-lookup"><span data-stu-id="69ada-171">See also</span></span>

- [<span data-ttu-id="69ada-172">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="69ada-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="69ada-173">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="69ada-173">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="69ada-174">Office アドインのリソースの制限とパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="69ada-174">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="69ada-175">ワークシート関数のオブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="69ada-175">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
