---
title: Excel JavaScript API のパフォーマンスの最適化
description: JavaScript API Excelを使用して、アドインのパフォーマンスを最適化します。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 5313bb3fe25d165e49cc0508e81d58294db48798
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349386"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Excel の JavaScript API を使用した、パフォーマンスの最適化

Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。 さまざまなアプローチの間でパフォーマンスは大きく異なります。 この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。

> [!IMPORTANT]
> 推奨される使用法と呼び出しによって、多くのパフォーマンスの問題 `load` に対処 `sync` できます。 アプリケーション固有の API を効率的に操作するためのアドバイスについては[、「Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis)アドインのリソース制限とパフォーマンスの最適化」の「アプリケーション固有 API によるパフォーマンスの向上」セクションを参照してください。

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

数式の計算だけが中断されます。 変更された参照は、まだ再作成されます。 たとえば、ワークシートの名前を変更すると、そのワークシートへの数式の参照が更新されます。

### <a name="suspend-screen-updating"></a>画面の更新を停止する

Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。 大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。 `Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。 Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。

> [!NOTE]
> 繰り返し `suspendScreenUpdatingUntilNextSync` 呼び出す (ループ内など) は使用しない。 繰り返し呼び出しを行Excelウィンドウがちらつきます。

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

* [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
* [ワークシート関数のオブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
