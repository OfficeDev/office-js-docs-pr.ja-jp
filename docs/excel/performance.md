---
title: Excel JavaScript API パフォーマンスの最適化
description: Excel JavaScript APIを使用してパフォーマンスを最適化して下さい。
ms.date: 03/28/2018
ms.openlocfilehash: 50fac999093abb3fbfe1bd5be1cd6a77dc930399
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797316"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Excel JavaScript APIを使用したパフォーマンスの最適化

Excel JavaScript APIを使用して一般的なタスクを実行するには、複数の方法があります。 さまざまなアプローチの間に大きなパフォーマンスの違いがあります。 この記事には、Excel JavaScript API を使用して一般的なタスクを効率的に実行する方法を示すガイダンスとコードサンプルが記載されています。

## <a name="minimize-the-number-of-sync-calls"></a>sync（）呼び出しの数を最小限にして下さい。

Excel JavaScript APIでは、 ```sync()``` 唯一の非同期操作であり、Excel Online の場合は特に状況によっては遅くなる可能性があります。 パフォーマンスを最適化するには、 ```sync()``` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にして下さい。

このプラクティスに従うコードサンプルについては  [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) を参照してください。

## <a name="minimize-the-number-of-proxy-objects-created"></a>作成されたプロキシオブジェクトの数を最小限にして下さい。

同じプロキシオブジェクトを繰り返し作成することは避けてください。 代わりに、複数の操作で同じプロキシオブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用して下さい。

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

## <a name="load-necessary-properties-only"></a>必要なプロパティのみをロードして下さい。

Excel JavaScript APIでは、プロキシオブジェクトのプロパティを明示的にロードする必要があります。 空の　```load()```　呼び出しで、すべてのプロパティを一度にロードすることはできますが、そのアプローチはかなりのパフォーマンスオーバーヘッドを持つ可能性があります。 代わりに、必要なプロパティだけをロードすることをお勧めします。特に、多数のプロパティを持つオブジェクトの場合はそうして下さい。

たとえば、範囲オブジェクトの **address** プロパティのみを読み取る場合 **load()** メソッドを呼び出すときにそのプロパティのみを指定します。
 
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
 
_場所：_
 
* `properties` コンマ区切り文字列または名前の並びとして指定された、ロードするプロパティのリストです。 詳細については、「[Excel JavaScript API リファレンス](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)」でオブジェクトに対して定義されている **load()** メソッドを参照してください。
* `loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption)を参照してください。

オブジェクトの下の「プロパティ」の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。 例えば、 `format` は範囲オブジェクトの下のプロパティですが、 `format` それ自体もオブジェクトです。 だから、あなたが `range.load("format")`のような呼び出しをすれば、これは以前に概説したようなパフォーマンスの問題を引き起こす可能性のある空のload（）である `range.format.load()` と等しいことになります。 これを避けるには、オブジェクトツリー内の "リーフノード"のみをロードするようにしてください。 

## <a name="suspend-calculation-temporarily"></a>一時的に計算を中断して下さい。

大量のセル（たとえば、巨大範囲オブジェクトの値を設定する）で操作を実行しようとしていて、操作が完了している間に一時的にExcelで計算を中断しても構わない場合は、次の ```context.sync()``` が呼び出されまで計算を中断することをおすすめします。

非常に便利な方法で計算を中断し、再起動するための ```suspendApiCalculationUntilNextSync()``` API の使用方法は [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) リファレンスドキュメントを参照してください。 次のコードは、計算を一時的に中断する方法を示しています。

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

## <a name="update-all-cells-in-a-range"></a>範囲内のすべてのセルの更新 

範囲内のすべてのセルを同じ値またはプロパティで更新する必要がある場合は、同じ値を繰り返し指定する2次元配列で行うと、更新が遅くなる可能性があります。このアプローチだと、範囲内のすべてのセルをExcelが反復しなければ、それぞれ個別に設定できないからです。 Excelには、範囲内のすべてのセルを同じ値またはプロパティで更新するより効率的な方法があります。

同じ値、同じ数値書式設定、同じ数式をセルの範囲に適用する必要がある場合は、値の配列ではなく単一の値を指定する方が効率的です。 そうすることで、パフォーマンスが大幅に向上します。 このアプローチが実際に動作していることを示すコードサンプルについては、 [コアの概念 - 範囲内のすべてのセルを更新する](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)を参照してください。

このアプローチが使える一般的なシナリオは、ワークシートの異なる列に異なる数値書式を設定する場合です。 この場合、列を通って反復し、各列の数値書式を単一の値で設定するだけです。 [範囲内のすべてのセルを更新する](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) コードサンプルにあるように、各列を範囲として扱ってください。

> [!NOTE]
> TypeScriptを使用している場合、1つの値を2次元配列に設定できないというコンパイルエラーが発生します。  その値 *は* プロパティを取得しているときは2次元配列なので、エラーは避けられません。TypeScriptでは、異なるセッター対ゲッターの型は許可されません。  しかし、簡単な回避策は、例えば、 `range.values = "hello world" as any` という `as any` 接尾辞で値を設定することです。

## <a name="importing-data-into-tables"></a>表へのデータのインポート

膨大な量のデータを直接 [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) オブジェクトにインポートする場合は（例えば、 `TableRowCollection.add()`を使用して）、パフォーマンスが低下する可能性があります。 新しいテーブルを追加しようとする場合は、最初に `range.values`を設定してデータを入力してください。次に `worksheet.tables.add()` を呼び出しその範囲にわたってテーブルを作成します。 既存のテーブルにデータを書き込もうとしている場合は、 `table.getDataBodyRange()`経由で範囲オブジェクトにデータを書き込んで下さい。テーブルが自動的に展開されます。 

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
> TableオブジェクトをRangeオブジェクトに変換するには、 [Table.convertToRange（）](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) 方法が便利です。

## <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。 イベントを有効化および無効化する方法を示すコード サンプルは、 [イベントでの作業](excel-add-ins-events.md#enable-and-disable-events) の記事に記載されています。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API の高度な概念](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API オープン仕様](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [ワークシート関数のオブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.functions)
