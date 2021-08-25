---
title: Excel JavaScript API のパフォーマンスの最適化
description: JavaScript API Excelを使用して、アドインのパフォーマンスを最適化します。
ms.date: 08/24/2021
localization_priority: Normal
ms.openlocfilehash: f65db836d6e7e640672fa5b9e6642ef8122ed5a5
ms.sourcegitcommit: 7ced26d588cca2231902bbba3f0032a0809e4a4a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2021
ms.locfileid: "58505657"
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
> [Table.convertToRange()](/javascript/api/excel/excel.table#convertToRange__) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。

## <a name="payload-size-limit-best-practices"></a>ペイロード サイズの制限のベスト プラクティス

JavaScript API Excel API 呼び出しのサイズ制限があります。 Excel on the web 5 MB の要求と応答のペイロード サイズ制限を持ち、この制限を超えると API はエラー `RichAPI.Error` を返します。 すべてのプラットフォームで、取得操作の範囲は 500 万セルに制限されます。 大きい範囲は、通常、これらの制限の両方を超える。

要求のペイロード サイズは、次の 3 つのコンポーネントの組み合わせです。

* API 呼び出しの数
* オブジェクトなどのオブジェクトの `Range` 数
* 設定または取得する値の長さ

API がエラーを返す場合は、この記事に記載されているベスト プラクティス戦略を使用して、スクリプトを最適化し、エラー `RequestPayloadSizeLimitExceeded` を回避してください。

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>戦略 1: 変更されていない値をループから移動する

パフォーマンスを向上させるために、ループ内で発生するプロセスの数を制限します。 次のコード サンプルでは、ループ内で変更しないので、ループから移動 `context.workbook.worksheets.getActiveWorksheet()` `for` できます。

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

次のコード サンプルは、前のコード サンプルと同様のロジックを示していますが、パフォーマンス戦略が改善されています。 ループが実行されるごとにこの値を取得する必要がないので、ループの前に値 `context.workbook.worksheets.getActiveWorksheet()` `for` が `for` 取得されます。 ループのコンテキスト内で変化する値のみをそのループ内で取得する必要があります。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    // Retrieve the worksheet outside the loop.
    var worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>戦略 2: 範囲オブジェクトの作成数を少なくする

パフォーマンスを向上し、ペイロード サイズを最小限に抑えるために、範囲オブジェクトを少なくします。 範囲オブジェクトを少なくするための 2 つの方法については、次の記事のセクションとコード サンプルで説明します。

#### <a name="split-each-range-array-into-multiple-arrays"></a>各範囲配列を複数の配列に分割する

範囲オブジェクトを少なくする方法の 1 つは、各範囲配列を複数の配列に分割し、ループと新しい呼び出しで各新しい配列を処理 `context.sync()` する方法です。

> [!IMPORTANT]
> ペイロード要求のサイズ制限を超えたと最初に判断した場合にのみ、この戦略を使用します。 複数のループを使用すると、5 MB の制限を超えしないように各ペイロード要求のサイズを小さくできますが、複数のループと複数の呼び出しを使用するとパフォーマンスに悪影響 `context.sync()` を及ぼす可能性があります。

次のコード サンプルでは、1 回のループで範囲の大規模な配列を処理し、次に 1 回の呼び出しを実行 `context.sync()` します。 1 回の呼び出しで範囲の値が多すぎると、ペイロード `context.sync()` 要求のサイズが 5 MB の制限を超える原因になります。

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      var range = sheet.getRangeByIndexes(i, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

次のコード サンプルは、前のコード サンプルと同様のロジックを示していますが、5 MB のペイロード要求サイズ制限を超えるのを回避する方法を示しています。 次のコード サンプルでは、範囲は 2 つの個別のループで処理され、各ループの後に呼び出しが続 `context.sync()` きます。

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      var range = worksheet.getRangeByIndexes(i, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      var range = worksheet.getRangeByIndexes(i, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>配列内の範囲の値を設定する

範囲オブジェクトを少なくするもう 1 つの方法は、配列を作成し、ループを使用してその配列内のすべてのデータを設定し、配列の値を範囲に渡す方法です。 これにより、パフォーマンスとペイロードのサイズの両方が向上します。 ループ内の各範囲 `range.values` を呼び出す代わりに、ループの外側 `range.values` で 1 回呼び出されます。

次のコード サンプルは、配列を作成し、ループ内の配列の値を設定し、その配列の値をループの外側の範囲 `for` に渡す方法を示しています。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (var i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    var range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>関連項目

* [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
* [JavaScript API のExcel処理](excel-add-ins-error-handling.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
* [ワークシート関数のオブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
