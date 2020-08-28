---
title: Excel JavaScript API のパフォーマンスの最適化
description: JavaScript API を使用して Excel アドインのパフォーマンスを最適化します。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: fdaccdca4779aaca64420794e382330994488606
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294102"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="77abc-103">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="77abc-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="77abc-104">Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。</span><span class="sxs-lookup"><span data-stu-id="77abc-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="77abc-105">さまざまなアプローチの間でパフォーマンスは大きく異なります。</span><span class="sxs-lookup"><span data-stu-id="77abc-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="77abc-106">この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。</span><span class="sxs-lookup"><span data-stu-id="77abc-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="77abc-107">パフォーマンスに関する多くの問題は、と呼び出しの推奨される使用方法によって解決でき `load` `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="77abc-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="77abc-108">アプリケーション固有の Api を効率的に処理するためのアドバイスについては、「 [リソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) 」の「アプリケーション固有の api を使用したパフォーマンスの向上」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="77abc-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="77abc-109">Excel のプロセスを一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="77abc-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="77abc-110">Excel には、ユーザーとアドインの両方からの入力に対応する多くのバックグラウンド タスクがあります。</span><span class="sxs-lookup"><span data-stu-id="77abc-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="77abc-111">これらの Excel のプロセスの一部は、パフォーマンス上の利点が得られるようにコントロールすることができます。</span><span class="sxs-lookup"><span data-stu-id="77abc-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="77abc-112">これは、アドインが大きなデータ セットを処理する場合に特に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="77abc-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="77abc-113">計算を一時的に中断する</span><span class="sxs-lookup"><span data-stu-id="77abc-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="77abc-114">大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の `context.sync()` が呼び出されまで計算を中断することをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="77abc-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="77abc-115">非常に便利な方法で計算を中断し、再起動するための `suspendApiCalculationUntilNextSync()` API の使用方法については、「[Application Object](/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="77abc-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="77abc-116">次のコードは、計算を一時的に中断する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="77abc-116">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

<span data-ttu-id="77abc-117">数式の計算のみが中断されることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="77abc-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="77abc-118">変更された参照はまだ再構築されています。</span><span class="sxs-lookup"><span data-stu-id="77abc-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="77abc-119">たとえば、ワークシートの名前を変更しても、そのワークシートへの数式の参照は更新されます。</span><span class="sxs-lookup"><span data-stu-id="77abc-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="77abc-120">画面の更新を停止する</span><span class="sxs-lookup"><span data-stu-id="77abc-120">Suspend screen updating</span></span>

<span data-ttu-id="77abc-121">Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。</span><span class="sxs-lookup"><span data-stu-id="77abc-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="77abc-122">大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。</span><span class="sxs-lookup"><span data-stu-id="77abc-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="77abc-123">`Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。</span><span class="sxs-lookup"><span data-stu-id="77abc-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="77abc-124">Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。</span><span class="sxs-lookup"><span data-stu-id="77abc-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="77abc-125">繰り返し呼び出しない `suspendScreenUpdatingUntilNextSync` (ループの場合など)。</span><span class="sxs-lookup"><span data-stu-id="77abc-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="77abc-126">呼び出しが繰り返し行われると、Excel ウィンドウがちらつくようになります。</span><span class="sxs-lookup"><span data-stu-id="77abc-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="77abc-127">イベントの有効化と無効化</span><span class="sxs-lookup"><span data-stu-id="77abc-127">Enable and disable events</span></span>

<span data-ttu-id="77abc-128">イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="77abc-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="77abc-129">イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。</span><span class="sxs-lookup"><span data-stu-id="77abc-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="77abc-130">テーブルへのデータのインポート</span><span class="sxs-lookup"><span data-stu-id="77abc-130">Importing data into tables</span></span>

<span data-ttu-id="77abc-131">膨大な量のデータを直接 [Table](/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="77abc-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="77abc-132">新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。</span><span class="sxs-lookup"><span data-stu-id="77abc-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="77abc-133">既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。</span><span class="sxs-lookup"><span data-stu-id="77abc-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="77abc-134">このアプローチの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="77abc-134">Here is an example of this approach:</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
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
> <span data-ttu-id="77abc-135">[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。</span><span class="sxs-lookup"><span data-stu-id="77abc-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="77abc-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="77abc-136">See also</span></span>

* [<span data-ttu-id="77abc-137">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="77abc-137">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="77abc-138">Office アドインのリソースの制限とパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="77abc-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="77abc-139">ワークシート関数のオブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="77abc-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
