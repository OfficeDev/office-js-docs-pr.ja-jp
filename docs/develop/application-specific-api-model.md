---
title: アプリケーション固有の API モデルの使用
description: Excel、OneNote、および Word アドインの Promise ベースの API モデルについて説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 568494dc0b92f1a4f9c6556b169293e68ae0bce9
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773497"
---
# <a name="application-specific-api-model"></a>アプリケーション固有の API モデル

この記事では、Excel、Word、OneNote でアドインを構築するために API モデルを使用する方法について説明します。 この説明では、Promise ベースの API の使用に基本的な主要な概念を説明します。

> [!NOTE]
> このモデルは、Office 2013 クライアントではサポートされていません。 これらの Office バージョンを使用しながら、[共通のAPIモデル](office-javascript-api-object-model.md) を使用します。 フル プラットフォーム可用性のノートについては、「[Office アドイン用 Office クライアント アプリケーションとプラットフォームの可用性](../overview/office-add-in-availability.md)」を参照してください。

> [!TIP]
> このページの例では Excel JavaScript API を使用しますが、概念は OneNote、Visio、Word JavaScript API にも適用されます。

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Promise ベース API の非同期の性質

Office アドインは、Excel などの Office アプリケーション内のブラウザー コンテナー内に表示される Web サイトです。 コンテナーは、Office on Windows などのデスクトップ ベースのプラットフォーム上の Office アプリケーションに組み込まれ、Office on the Web の HTML iFrame 内で実行されます。 パフォーマンスの考慮事項により、Office.js API は、すべてのプラットフォームの Office アプリケーションと同期して対話することはできません。 このため、`sync()`Office.js 内の API 呼び出しは Office アプリケーションが要求された読み取りまたは書き込み操作を完了したときに解決された[ Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返します。 また、操作ごとに別個の要求として送信する代わりに、プロパティの設定やメソッドの起動など、複数の操作をキューに登録し、`sync()`への 1 回の呼び出しでコマンドのバッチとしてそれらを実行することもできます。 次のセクションでは、`run()` および `sync()` API を使用してこれを実行する方法について説明します。

## <a name="run-function"></a>*.run 関数

`Excel.run`、 `Word.run`、 `OneNote.run`は、Excel、Word、OneNote に対して実行するアクションを指定する関数を実行できます。 `*.run` は Office オブジェクトと対話するために使用できる要求コンテキストを自動的に作成します。 `*.run`が完了すると、Promose が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。

次の例は、`Excel.run`の使用方法を説明しています。 Word と OneNote でも同じパターンが使用されます。

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a>要求コンテキスト

Office アプリケーションとユーザーのアドインは、2 つの異なるプロセスで実行されます。 それらは異なるランタイム環境を使用するため、アドインは、ワークシート、範囲、グラフ、表など、Office のオブジェクトにユーザーのアドインを接続するために `RequestContext` オブジェクトが必要です。 この `RequestContext` オブジェクトは、`*.run`を呼び出す際に引数として提供されます。

## <a name="proxy-objects"></a>プロキシ オブジェクト

Promise ベースの API と共にユーザーが宣言して使用する Office JavaScript オブジェクトはプロキシ オブジェクトです。 起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。 要求コンテキスト上 (たとえば `context.sync()`) で `sync()`メソッドを呼び出すと、キューに入れられたコマンドは Office アプリケーションにディスパッチされて実行されます。 これらの API は、基本的にバッチ中心です。 要求コンテキストに必要なだけ変更内容をキューに登録し、`sync()` メソッドを呼び出して、キューに入れられたコマンドをバッチで実行することができます。

たとえば、次のコード スニペットでは、ローカル JavaScript [Excel.Range](/javascript/api/excel/excel.range) オブジェクト、`selectedRange`が Excel ワークブック内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。 `selectedRange` オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが `context.sync()` を呼び出すまで Excel ドキュメントには反映されません。

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>作業のヒント: 作成されたプロキシ オブジェクトの数を最小限にする

同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。 代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
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

### <a name="sync"></a>sync()

要求コンテキストで `sync()`メソッドを呼び出すと、プロキシ オブジェクトと Officeドキュメント内のオブジェクトの状態が同期されます。 `sync()` メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。   `sync()`メソッドは非同期で実行されて [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。これは、`sync()` メソッドが完了すると解決されます。

次の例は、ローカル JavaScript proxy オブジェクト (`selectedRange`) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して `context.sync()` を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

前の例では、`selectedRange` が設定されており、`context.sync()` が呼び出されると `address` プロパティが読み込まれます。

`sync()`が非同期操作である場合、スクリプトが引き続き実行される前に、`Promise` オブジェクトを返して、`sync()`の操作が完了するのを確認する必要があります。 TypeScript または ES6+ JavaScript を使用している場合は、Promise を返す代わりに `context.sync()` の呼び出しを`await` にできます。

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>作業のこつ: 同期呼び出しの数を最小限にする

Excel JavaScript API では、`sync()` は唯一の非同期操作で、状況によっては遅くなる可能性があり、Excel on the web の場合は特にその傾向があります。 パフォーマンスを最適化するには、`sync()` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。 パフォーマンスを`sync()`で最適化する方法の詳細については、「[ループで context.sync メソッドの使用を避ける](../concepts/correlated-objects-pattern.md)」をご参照ください。

### <a name="load"></a>load()

プロキシ オブジェクトのプロパティを読み取るには、まず Office ドキュメントからプロキシ オブジェクトとデータを入力するためにプロパティを明確に読み込み、`context.sync()`を呼び出す必要があります。 たとえば、選択範囲を操作するプロキシ オブジェクトを作成してから選択範囲の`address` プロパティを読み取る場合、読み取る前に`address` プロパティを読み込む必要があります。 読み込むプロキシ オブジェクトのプロパティを要求するには、オブジェクトに対して `load()` メソッドを呼び出し、読み込むプロパティを指定します。 次の例は、`myRange`に読み込まれているプロパティ `Range.address`を示しています 。

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、`load()` メソッドを呼び出す必要はありません。   `load()` メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。

プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 `sync()` メソッドを呼び出すときに実行されます。`load()` の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。

#### <a name="scalar-and-navigation-properties"></a>スカラー プロパティとナビゲーション プロパティ

プロパティには、**スカラー** と **ナビゲーション** という 2 つのカテゴリがあります。 スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。 ナビゲーション プロパティは、プロパティを直接割り当てるのではなく、読み取り専用のオブジェクトと、そのフィールドが割り当てられているオブジェクトのコレクションです。 たとえば、[Excel.Worksheet](/javascript/api/excel/excel.worksheet)のオブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はナビゲーション プロパティです。

アドインは、特定のスカラー プロパティを読み込むパスとしてナビゲーション プロパティを使用できます。 次のコードは、ほかの情報を読み込む必要なく、`Excel.Range` オブジェクトで使用されるフォント名の`load` コマンドをキューに入れられます。

```js
someRange.load("format/font/name")
```

パスを詳しく調べることでナビゲーション プロパティのスカラー プロパティを設定できます。 たとえば、`someRange.format.font.size = 10;`を使用して`Excel.Range` のフォント サイズを設定できます。 設定前にプロパティを読み込む必要はありません。

オブジェクトの下の「プロパティ」の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。 例えば、`format` は`Excel.Range`オブジェクトの下のプロパティですが、`format` それ自体もオブジェクトです。 そのため、`range.load("format")`などの呼び出しを行った場合、これは `range.format.load()` (望ましくない空の空白のステートメント`load()`) と同等になります。 これを避けるには、コードがオブジェクト ツリー内の "リーフノード" のみをロードするようにしてください。

#### <a name="calling-load-without-parameters-not-recommended"></a>パラメーターを使用せず (非推奨) に `load` を呼び出す

パラメーターを指定せずにオブジェクト (またはコレクション) の `load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティ (またはコレクション内のオブジェクト) が読み込まれます。 不要なデータを読み込むと、アドインの速度が低下します。 常に読み込むプロパティを明示的に指定する必要があります。

> [!IMPORTANT]
> パラメーターのない `load` ステートメントで返されるデータの量は、サービスのサイズ制限を超える場合があります。 古いアドインのリスクを軽減するために、明示的に要求しない限り `load` によって返されないプロパティがあります。 次のプロパティは、このような読み込み操作から除外されます。
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

プリミティブ型を返す、Promise ベースの API 内のメソッドは、`load`/`sync`パラダイムと同様のパターンを持っています。 たとえば、`Excel.TableCollection.getCount` はコレクション内のテーブルの数を取得します。 `getCount` は `ClientResult<number>` を返します。つまり、返される[`ClientResult`](/javascript/api/office/officeextension.clientresult)の`value`プロパティは数値になります。 `context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。

次のコードは、Excel ワークブック内のテーブルの総数を取得し、その数をコンソールに記録します。

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a>set()

入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。 前述のナビゲーション パスを使用してプロパティを個別に設定する代わりに、Promise ベースの JavaScript API のオブジェクトで使用できる、`object.set()`メソッドを使用できます。 このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。

次のコード サンプルは、`set()` メソッドを呼び出し、`Range`Range **オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲** B2:E2 にデータがあると仮定します。

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a>一部のプロパティを直接設定できません

書き込み可能であるにもかかわらず、一部のプロパティを設定できません。 これらのプロパティは、1 つのオブジェクトとして設定する必要がある親プロパティの一部です。 これは、親プロパティが特定の論理関係を持つサブプロパティに依存しているからです。 これらの親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクト全体を設定するためにオブジェクト リテラル表記を使用して設定する必要があります。 その 1 つの例は、[PageLayout](/javascript/api/excel/excel.pagelayout)にあります。 プロパティ `zoom` は、次に示すように [、1 つの PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) オブジェクトで設定する必要があります。

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

前の例では、`zoom` 値: `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることは ***できません***。 このステートメントは、`zoom` が読み込まれないので、エラーを発生させます。 `zoom` が読み込まれるような場合でも、スケール セットは有効化されません。 すべてのコンテキスト操作は `zoom`上、でアドインのプロキシオブジェクトを更新し、ローカルに設定された値を上書きする場合に発生します。

この動作は、[Range.format](/javascript/api/excel/excel.range#format)など、[ナビゲーション プロパティ](application-specific-api-model.md#scalar-and-navigation-properties) とは異なります。 プロパティは `format` 、次に示すようにオブジェクト ナビゲーションを使用して設定できます。

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

読み取り専用の修飾キーを確認することで、サブプロパティを直接設定できないプロパティを識別できます。 読み取り専用プロパティはすべて、読み取り専用以外のサブプロパティを直接設定できます。 `PageLayout.zoom` のような書き可能なプロパティは、そのレベルのオブジェクトで設定する必要があります。 まとめると、以下のようになります。

- 読み取り専用プロパティ: ナビゲーション経由でサブプロパティを設定できます。
- 書き込み可能なプロパティ: サブプロパティをナビゲーションを介して設定することはできません (最初の親オブジェクトの一部として設定する必要があります)。

## <a name="42ornullobject-methods-and-properties"></a>&#42;OrNullObject メソッドとプロパティ

一部のアクセサリ方法とプロパティでは、目的のオブジェクトが存在しない場合に例外をスローします。 たとえば、ブックに存在しないワークシート名を指定して Excel ワークシートを取得しようとすると、`getItem()` メソッドは `ItemNotFound` 例外を返します。 アプリケーション固有のライブラリを使用すると、例外処理コードを必要とせずに、コードがドキュメント エンティティの存在をテストできます。 これは、`*OrNullObject`メソッドのバリエーションとプロパティ を使用して行います。 これらのバリエーションは、 `isNullObject` プロパティが `true`に設定されているオブジェクトを返します (指定したアイテムが存在しない場合は、例外をスローしません)。

たとえば、**Worksheets** などのコレクションで `getItemOrNullObject()` メソッドを呼び出して、コレクションからのアイテムの取得を試行できます。 `getItemOrNullObject()` メソッドは、指定された項目が存在する場合はその項目を返し、それ以外の場合は `isNullObject`プロパティが `true`に設定されているオブジェクトを返します。 コードは、このプロパティを評価して、オブジェクトが存在するかどうかを判断できます。

> [!NOTE]
> `*OrNullObject` のバリエーションは、JavaScript 値`null`を返すことはありません。 通常の Office プロキシ オブジェクトを返します。 オブジェクトが表すエンティティが存在しない場合は、オブジェクトの `isNullObject` プロパティが `true`に設定されます。 返されたオブジェクトの null 値または 真偽性はテストしません。 これは、決して `null`、 `false`、`undefined`ではありません。

次のコード サンプルは `getItemOrNullObject()` メソッドを使用して、"Data" という名前のワークシートの取得を試行します。 その名前のワークシートが存在しない場合は、新しいシートが作成されます。 コードは`isNullObject`プロパティを読み込まないことにご注意ください。 Office は、`context.sync`が呼ばれると、自動的にこのプロパティを読み込みます。ですから、`datasheet.load('isNullObject')`のような名前で明示的に読み込む必要はありません。

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a>関連項目

* [共通 JavaScript API オブジェクト モデル](office-javascript-api-object-model.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
