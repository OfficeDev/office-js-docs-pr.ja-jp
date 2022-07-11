---
title: Excel JavaScript API のパフォーマンスの最適化
description: JavaScript API を使用して Excel アドインのパフォーマンスを最適化します。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: bad5d35ec1cc3f99cd37b3571dee78d3432102e6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712728"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Excel の JavaScript API を使用した、パフォーマンスの最適化

Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。 さまざまなアプローチの間でパフォーマンスは大きく異なります。 この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。

> [!IMPORTANT]
> パフォーマンスの問題の多くは、推奨される使用法 `load` と呼び出し `sync` によって対処できます。 効率的な方法でのアプリケーション固有の API の操作に関するアドバイスについては、「 [Office アドインのリソース制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) 」の「アプリケーション固有の API のパフォーマンスの向上」セクションを参照してください。

## <a name="suspend-excel-processes-temporarily"></a>Excel のプロセスを一時的に中断する

Excel には、ユーザーとアドインの両方からの入力に対応する多くのバックグラウンド タスクがあります。 これらの Excel のプロセスの一部は、パフォーマンス上の利点が得られるようにコントロールすることができます。 これは、アドインが大きなデータ セットを処理する場合に特に役立ちます。

### <a name="suspend-calculation-temporarily"></a>計算を一時的に中断する

大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の `context.sync()` が呼び出されまで計算を中断することをおすすめします。

非常に便利な方法で計算を中断し、再起動するための `suspendApiCalculationUntilNextSync()` API の使用方法については、「[Application Object](/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。 次のコードは、計算を一時的に中断する方法を示しています。

```js
await Excel.run(async (context) => {
    let app = context.workbook.application;
    let sheet = context.workbook.worksheets.getItem("sheet1");
    let rangeToSet: Excel.Range;
    let rangeToGet: Excel.Range;
    app.load("calculationMode");
    await context.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await context.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await context.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await context.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
});
```

数式の計算のみが中断されることに注意してください。 変更された参照は、引き続き再構築されます。 たとえば、ワークシートの名前を変更しても、そのワークシートに対する数式内のすべての参照が更新されます。

### <a name="suspend-screen-updating"></a>画面の更新を停止する

Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。 大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。 `Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。 Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。

> [!NOTE]
> 繰り返し呼び出 `suspendScreenUpdatingUntilNextSync` さないでください (ループ内など)。 繰り返し呼び出すと、Excel ウィンドウがちらつきます。

### <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。 イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。

## <a name="importing-data-into-tables"></a>テーブルへのデータのインポート

膨大な量のデータを直接 [Table](/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。 新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。 既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。

このアプローチの例を次に示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    let range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    let table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await context.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await context.sync();
});
```

> [!NOTE]
> [Table.convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1)) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。

## <a name="payload-size-limit-best-practices"></a>ペイロード サイズ制限のベスト プラクティス

Excel JavaScript API には、API 呼び出しのサイズ制限があります。 Excel on the webには 5 MB の要求と応答に対するペイロード サイズの制限があり、この制限を超えると API はエラーを`RichAPI.Error`返します。 すべてのプラットフォームで、取得操作の範囲は 500 万セルに制限されています。 大きな範囲は、通常、これらの制限の両方を超えています。

要求のペイロード サイズは、次の 3 つのコンポーネントの組み合わせです。

* API 呼び出しの数
* オブジェクトなどの `Range` オブジェクトの数。
* 設定または取得する値の長さ

API からエラーが `RequestPayloadSizeLimitExceeded` 返される場合は、この記事に記載されているベスト プラクティスの戦略を使用して、スクリプトを最適化し、エラーを回避します。

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>戦略 1: 変更されていない値をループから移動する

パフォーマンスを向上させるために、ループ内で発生するプロセスの数を制限します。 次のコード サンプルでは、 `context.workbook.worksheets.getActiveWorksheet()` ループ内で変更されないため、ループから `for` 移動できます。

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

次のコード サンプルは、前のコード サンプルに似たロジックを示していますが、パフォーマンス戦略は改善されています。 ループが実行されるたびにこの値を`for`取得する必要がないため、ループの前に値`context.workbook.worksheets.getActiveWorksheet()`が`for`取得されます。 ループのコンテキスト内で変化する値のみを、そのループ内で取得する必要があります。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    // Retrieve the worksheet outside the loop.
    let worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>戦略 2: 範囲オブジェクトを少なくする

より少ない範囲オブジェクトを作成して、パフォーマンスを向上させ、ペイロード サイズを最小限に抑えます。 次の記事のセクションとコード サンプルでは、より少ない範囲オブジェクトを作成するための 2 つの方法について説明します。

#### <a name="split-each-range-array-into-multiple-arrays"></a>各範囲配列を複数の配列に分割する

範囲オブジェクトを少なくする方法の 1 つは、各範囲配列を複数の配列に分割し、ループと新 `context.sync()` しい呼び出しで各新しい配列を処理することです。

> [!IMPORTANT]
> この戦略は、ペイロード要求サイズの制限を超えていると最初に判断した場合にのみ使用します。 複数のループを使用すると、5 MB の制限を超えないように各ペイロード要求のサイズを小さくできますが、複数のループと複数の `context.sync()` 呼び出しを使用すると、パフォーマンスにも悪影響を及ぼします。

次のコード サンプルでは、1 つのループ内の範囲の大きな配列を処理し、次に 1 回 `context.sync()` の呼び出しを試みます。 1 回 `context.sync()` の呼び出しで処理する範囲の値が多すぎると、ペイロード要求サイズが 5 MB の制限を超えます。

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      let range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

次のコード サンプルは、上記のコード サンプルと同様のロジックを示していますが、5 MB ペイロード要求サイズの制限を超えないようにする戦略を示しています。 次のコード サンプルでは、範囲は 2 つの別々のループで処理され、各ループの後に `context.sync()` 呼び出しが続きます。

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>配列内の範囲値を設定する

範囲オブジェクトを少なくするもう 1 つの方法は、配列を作成し、ループを使用してその配列内のすべてのデータを設定してから、配列の値を範囲に渡すことです。 これにより、パフォーマンスとペイロードのサイズの両方にメリットがあります。 ループ内の各範囲を呼び出す `range.values` 代わりに、 `range.values` ループの外側で 1 回呼び出されます。

次のコード サンプルは、配列を作成し、その配列の値をループ内に `for` 設定してから、ループの外側の範囲に配列値を渡す方法を示しています。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (let i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    let range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>関連項目

* [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
* [アプリケーション固有の JavaScript API でのエラー処理](../testing/application-specific-api-error-handling.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
* [ワークシート関数のオブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
