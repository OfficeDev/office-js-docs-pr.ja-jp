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
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Excel の JavaScript API を使用した、パフォーマンスの最適化

Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。 さまざまなアプローチの間でパフォーマンスは大きく異なります。 この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。

> [!IMPORTANT]
> パフォーマンスに関する多くの問題は、と呼び出しの推奨される使用方法によって解決でき `load` `sync` ます。 アプリケーション固有の Api を効率的に処理するためのアドバイスについては、「 [リソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) 」の「アプリケーション固有の api を使用したパフォーマンスの向上」を参照してください。

## <a name="suspend-excel-processes-temporarily"></a>Excel のプロセスを一時的に中断する

Excel には、ユーザーとアドインの両方からの入力に対応する多くのバックグラウンド タスクがあります。 これらの Excel のプロセスの一部は、パフォーマンス上の利点が得られるようにコントロールすることができます。 これは、アドインが大きなデータ セットを処理する場合に特に役立ちます。

### <a name="suspend-calculation-temporarily"></a>計算を一時的に中断する

大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の `context.sync()` が呼び出されまで計算を中断することをおすすめします。

非常に便利な方法で計算を中断し、再起動するための `suspendApiCalculationUntilNextSync()` API の使用方法については、「[Application Object](/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。 次のコードは、計算を一時的に中断する方法を示しています。

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

数式の計算のみが中断されることに注意してください。 変更された参照はまだ再構築されています。 たとえば、ワークシートの名前を変更しても、そのワークシートへの数式の参照は更新されます。

### <a name="suspend-screen-updating"></a>画面の更新を停止する

Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。 大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。 `Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。 Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。

> [!NOTE]
> 繰り返し呼び出しない `suspendScreenUpdatingUntilNextSync` (ループの場合など)。 呼び出しが繰り返し行われると、Excel ウィンドウがちらつくようになります。

### <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。 イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。

## <a name="importing-data-into-tables"></a>テーブルへのデータのインポート

膨大な量のデータを直接 [Table](/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。 新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。 既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。

このアプローチの例を次に示します。

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
> [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。

## <a name="see-also"></a>関連項目

* [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
* [ワークシート関数のオブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
