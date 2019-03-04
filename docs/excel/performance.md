---
title: Excel JavaScript API のパフォーマンスの最適化
description: Excel JavaScript API を使用してパフォーマンスを最適化する
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: d15a4b3ad4ae44399572282889855b1cdc32bc39
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199579"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Excel の JavaScript API を使用した、パフォーマンスの最適化

Excel JavaScript API を使用して一般的なタスクを実行するには、複数の方法があります。 さまざまなアプローチの間でパフォーマンスは大きく異なります。 この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコード サンプルが記載されています。

## <a name="minimize-the-number-of-sync-calls"></a>sync() 呼び出しの数を最小限にする

Excel JavaScript API では、```sync()``` は唯一の非同期操作であり、状況によっては遅くなる可能性があり、Excel Online の場合は特にその傾向があります。 パフォーマンスを最適化するには、```sync()``` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。

このプラクティスに従うコード サンプルについては 「[Core Concepts - sync()](excel-add-ins-core-concepts.md#sync)」を参照してください。

## <a name="minimize-the-number-of-proxy-objects-created"></a>作成されたプロキシ オブジェクトの数を最小限にする

同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。 代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。

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

## <a name="load-necessary-properties-only"></a>必要なプロパティのみをロードする

Excel JavaScript API では、プロキシ オブジェクトのプロパティを明示的にロードする必要があります。  空の ```load()``` 呼び出しで、すべてのプロパティを一度にロードすることはできますが、そのアプローチは大きなパフォーマンス オーバーヘッドを持つ可能性があります。  代わりに、必要なプロパティだけをロードすることをお勧めします。特に、多数のプロパティを持つオブジェクトの場合はそうして下さい。

たとえば、範囲オブジェクトの **address** プロパティのみを読み取る場合、**load()** メソッドを呼び出すときにそのプロパティのみを指定します。
 
```js
range.load('address');
```
 
**load()** メソッドは、次のいずれかの方法で呼び出すことができます。
 
_構文:_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_各部分の意味は次のとおりです。_
 
* `properties` は、ロードするプロパティの一覧で、コンマ区切りの文字列または名前の配列として指定されます。 詳細については、「[Excel JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)」でオブジェクトに対して定義されている **load()** メソッドを参照してください。
* `loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption)を参照してください。

オブジェクトの下の「プロパティ」の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。 例えば、`format` は範囲オブジェクトの下のプロパティですが、`format` それ自体もオブジェクトです。 そのため、`range.load("format")` のような呼び出しをすると、これは以前に概説したように、パフォーマンスの問題を引き起こす可能性のある空の load() 呼び出しである `range.format.load()` に相当します。 これを避けるには、オブジェクト ツリー内の "リーフノード" のみをロードするようにしてください。 

## <a name="suspend-excel-processes-temporarily"></a>Excel のプロセスを一時的に中断する

Excel には、ユーザーとアドインの両方からの入力に対応する多くのバックグラウンド タスクがあります。 これらの Excel のプロセスの一部は、パフォーマンス上の利点が得られるようにコントロールすることができます。 これは、アドインが大きなデータ セットを処理する場合に特に役立ちます。

### <a name="suspend-calculation-temporarily"></a>計算を一時的に中断する

大量のセル (たとえば、巨大範囲オブジェクトの値を設定する) で操作を実行しようとしていて、操作が完了するまでの間に一時的に Excel で計算が中断されても構わない場合は、次の `context.sync()` が呼び出されまで計算を中断することをおすすめします。

非常に便利な方法で計算を中断し、再起動するための `suspendApiCalculationUntilNextSync()` API の使用方法については、「[Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application)」リファレンスドキュメントを参照してください。 次のコードは、計算を一時的に中断する方法を示しています。

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

### <a name="suspend-screen-updating"></a>画面の更新を停止する

> [!NOTE]
> この記事に記載されている `suspendScreenUpdatingUntilNextSync` メソッドは、現在パブリック プレビューでのみ使用できます。 [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

Excel では、コード内で発生したのとほぼ同時に、アドインによって行われた変更が表示されます。 大規模で反復的なデータ セットの場合は、進捗状況の画面上での確認をリアルタイムで行う必要はありません。 `Application.suspendScreenUpdatingUntilNextSync()` は、アドインが `context.sync()` を呼び出すまで、または `Excel.run` が終了するまで (`context.sync` を暗黙的に呼び出す)、Excel のビジュアルの更新を一時停止します。 Excel では、更新停止の通知や表示などが次回の同期まで行われません。この遅延の準備のガイダンスや、アクティビティを示すステータス バーが、アドインによって提供される必要があります。

### <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。 イベントを有効化および無効化する方法を示すコード サンプルは、「[イベントの操作](excel-add-ins-events.md#enable-and-disable-events)」の記事に記載されています。

## <a name="update-all-cells-in-a-range"></a>範囲内のすべてのセルの更新

範囲内のすべてのセルを同じ値またはプロパティで更新する必要がある場合は、同じ値を繰り返し指定する 2 次元配列で行うと、更新が遅くなる可能性があります。このアプローチだと、範囲内のすべてのセルを Excel が反復しなければ、それぞれ個別に設定できないからです。 Excel には、範囲内のすべてのセルを同じ値またはプロパティで更新するより効率的な方法が備わっています。

セルの範囲に同じ値、同じ形式または同次数式を適用する必要がある場合は、配列の値の代わりに 1 つの値を指定する方が効率的です。 そうすることで、パフォーマンスが大幅に向上します。 このアプローチが実際に動作していることを示すコード サンプルについては、「[コアの概念 - 範囲内のすべてのセルを更新](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)」を参照してください。

このアプローチが使える一般的なシナリオは、ワークシートの異なる列に異なる数値書式を設定する場合です。  この場合、列を通って反復し、各列の数値書式を単一の値で設定するだけです。 「[範囲内のすべてのセルを更新する](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)」のコード サンプルにあるように、各列を範囲として扱います。

> [!NOTE]
> TypeScript を使用している場合は、2 次元配列に 1 つの値を設定できないことを示すコンパイル エラーが表示されます。  その値*は*プロパティを取得しているときは 2 次元配列なので、エラーは避けられません。TypeScript では、異なるセッター対ゲッターの型は許可されません。  しかし、簡単な回避策として、`as any` 接尾辞 (例: `range.values = "hello world" as any`) で値を設定する方法があります。

## <a name="importing-data-into-tables"></a>テーブルへのデータのインポート

膨大な量のデータを直接 [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) オブジェクトにインポートする場合は (例えば、`TableRowCollection.add()` を使用して)、パフォーマンスが低下する可能性があります。 新しいテーブルを追加しようとする場合は、最初に `range.values` を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。 既存のテーブルにデータを書き込もうとしている場合は、`table.getDataBodyRange()` 経由で範囲オブジェクトにデータを書き込みます。テーブルが自動的に展開されます。 

このアプローチの例を次に示します。

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
> [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) メソッドを使用すると、Table オブジェクトを Range オブジェクトに簡単に変換できます。

## <a name="untrack-unneeded-ranges"></a>不要になった範囲の追跡解除

JavaScript レイヤーは、アドインが Excel のブックと基になる範囲を操作するためのプロキシ オブジェクトを作成します。 こうしたオブジェクトは、`context.sync()` が呼び出されるまでメモリに維持されます。 大規模なバッチ操作では、アドインが 1 回のみ必要とするプロキシ オブジェクトが大量に生成されることがあります。それらのオブジェクトは、バッチの実行前にメモリから解放できます。

[Range.untrack()](/javascript/api/excel/excel.range#untrack--) メソッドにより、Excel の Range オブジェクトがメモリから解放されます。 範囲に対してアドインを実行した後に、このメソッドを呼び出すと、大量の Range オブジェクトを使用しているときのパフォーマンスが大幅に向上します。

> [!NOTE]
> `Range.untrack()` は、[ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-) のショートカットです。 プロキシ オブジェクトは、コンテキスト内の追跡対象オブジェクト リストから削除することで追跡解除できます。 通常、Range オブジェクトは追跡の解除が正当化されるほどの量で使用される唯一の Excel オブジェクトです。

次のコード例では、指定した範囲に 1 セルずつデータを埋め込みます。 セルに値が追加されると、そのセルを表している範囲の追跡が解除されます。 10,000 から 20,000 個のセルの範囲を選択して、このコードを実行します。最初の実行では `cell.untrack()` の行を使用し、その後でこの行を削除して実行します。 `cell.untrack()` の行がないコードよりも、この行があるコードの方が高速になることがわかります。 また、クリーンアップの手順にかかる時間が短くなるため、その後の応答時間も速くなることがわかります。

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

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用した高度なプログラミングの概念](excel-add-ins-advanced-concepts.md)
- [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
- [Excel JavaScript API オープン仕様](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [ワークシート関数のオブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.functions)
