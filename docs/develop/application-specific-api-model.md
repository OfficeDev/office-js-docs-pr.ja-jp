---
title: アプリケーション固有の API モデルの使用
description: Excel、OneNote、および Word のアドインの promise ベースの API モデルについて説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408601"
---
# <a name="using-the-application-specific-api-model"></a>アプリケーション固有の API モデルの使用

この記事では、Excel、Word、および OneNote でアドインをビルドするための API モデルの使用方法について説明します。 Promise ベースの Api を使用するための基本的な概念について説明します。

> [!NOTE]
> このモデルは、Office 2013 クライアントではサポートされていません。 [共通 API モデル](office-javascript-api-object-model.md)を使用して、これらの Office バージョンを操作します。 完全なプラットフォームの可用性に関する注意事項については、「office [クライアントアプリケーションおよび Office アドインのプラットフォームの可用性](../overview/office-add-in-availability.md)」を参照してください。

> [!TIP]
> このページの例では、Excel JavaScript Api を使用していますが、概念は OneNote、Visio、および Word JavaScript Api にも適用されます。

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>Promise ベースの Api の非同期的な性質

Office アドインは、Excel などの Office アプリケーション内のブラウザーコンテナー内に表示される web サイトです。 このコンテナーは、office アプリケーション内のデスクトップベースのプラットフォーム (Windows 上の Office など) に組み込まれており、web 上の Office の HTML iFrame 内で実行されます。 パフォーマンスに関する考慮事項のため、Office.js Api は、すべてのプラットフォームで Office アプリケーションと同期して操作することはできません。 したがって、 `sync()` Office.js の API 呼び出しは、Office アプリケーションが要求された読み取りまたは書き込みアクションを完了したときに解決される [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。 また、 `sync()` 各アクションに対して個別の要求を送信するのではなく、プロパティの設定やメソッドの呼び出しなど、複数のアクションをキューに入れて、1回の呼び出しで1つのコマンドのバッチとして実行することもできます。 次のセクションでは、api を使用してこれを実現する方法について説明し `run()` `sync()` ます。

## <a name="run-function"></a>*. run 関数

`Excel.run`、 `Word.run` 、 `OneNote.run` Excel、Word、および OneNote に対して実行するアクションを指定する関数を実行します。 `*.run` Office オブジェクトを操作するために使用できる要求コンテキストを自動的に作成します。 完了すると `*.run` 、promise が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。

次の例は、を使用する方法を示して `Excel.run` います。 Word と OneNote でも同じパターンが使用されます。

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

Office アプリケーションとアドインは、2つの異なるプロセスで実行されます。 さまざまなランタイム環境を使用しているため、アドインを `RequestContext` Office のオブジェクト (ワークシート、範囲、段落、表など) に接続するために、アドインにはオブジェクトが必要です。 この `RequestContext` オブジェクトは、を呼び出すときに引数として提供され `*.run` ます。

## <a name="proxy-objects"></a>プロキシ オブジェクト

Promise ベースの Api を宣言して使用する Office JavaScript オブジェクトは、プロキシオブジェクトです。 起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。 `sync()`(など) 要求コンテキストでメソッドを呼び出すと `context.sync()` 、キューに入れられたコマンドが Office アプリケーションにディスパッチされて実行されます。 これらの Api は、基本的にバッチ中心です。 要求コンテキストに対して必要な数だけ変更をキューに入れて、キューに `sync()` 入れられたコマンドのバッチを実行するメソッドを呼び出します。

たとえば、次のコードスニペットでは、ローカルの JavaScript [excel. range](/javascript/api/excel/excel.range) オブジェクトを宣言して、 `selectedRange` excel ブック内の選択された範囲を参照し、そのオブジェクトにいくつかのプロパティを設定します。 `selectedRange`オブジェクトはプロキシオブジェクトなので、設定されているプロパティと、そのオブジェクトに対して呼び出されたメソッドは、アドインが呼び出されるまで Excel ドキュメントには反映されません `context.sync()` 。

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>パフォーマンスのヒント: 作成されるプロキシオブジェクトの数を最小限に抑える

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

`sync()`要求コンテキストに対してメソッドを呼び出すと、Office ドキュメント内のプロキシオブジェクトとオブジェクトの間で状態が同期されます。 この `sync()` メソッドは、要求コンテキストでキューに入れられた任意のコマンドを実行し、プロキシオブジェクトに読み込む必要があるすべてのプロパティの値を取得します。 `sync()`メソッドは非同期的に実行され、 [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返します。これは、メソッドが完了したときに解決され `sync()` ます。

次の例は、ローカルの JavaScript プロキシオブジェクト () を定義し `selectedRange` 、そのオブジェクトのプロパティを読み込んでから、JavaScript の約束パターンを使用して、 `context.sync()` Excel ドキュメント内のプロキシオブジェクトとオブジェクト間の状態を同期するバッチ関数を示しています。

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

`sync()`は非同期操作なので、スクリプトを実行し `Promise` 続ける前に操作が完了したことを確認するために、常にオブジェクトを返してください `sync()` 。 TypeScript または ES6 + JavaScript を使用している場合は、promise を返す代わりに呼び出しを行うことができ `await` `context.sync()` ます。

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>パフォーマンスに関するヒント: 同期呼び出しの数を最小限に抑える

Excel JavaScript API では、`sync()` は唯一の非同期操作で、状況によっては遅くなる可能性があり、Excel on the web の場合は特にその傾向があります。 パフォーマンスを最適化するには、`sync()` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。 でパフォーマンスを最適化する方法の詳細について `sync()` は、「 [ループでのコンテキストの同期を回避する](../concepts/correlated-objects-pattern.md)」を参照してください。

### <a name="load"></a>load()

プロキシオブジェクトのプロパティを読み取るには、その前にプロパティを明示的に読み込んで、Office ドキュメントからのデータをプロキシオブジェクトに設定してから、を呼び出して `context.sync()` ください。 たとえば、選択した範囲を参照するプロキシオブジェクトを作成してから、選択した範囲のプロパティを読み取る場合は、 `address` そのプロパティを読み込む前に読み込む必要があり `address` ます。 プロキシオブジェクトのプロパティを読み込むには、 `load()` オブジェクトに対してメソッドを呼び出し、読み込むプロパティを指定します。 次の例は、 `Range.address` 読み込むプロパティを示して `myRange` います。

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
> プロキシオブジェクトのメソッドを呼び出すか、プロパティを設定するだけの場合は、メソッドを呼び出す必要はありません `load()` 。 この `load()` メソッドは、プロキシオブジェクトのプロパティを読み取る場合にのみ必要です。

プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 `sync()` メソッドを呼び出すときに実行されます。`load()` の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。

#### <a name="scalar-and-navigation-properties"></a>スカラー プロパティとナビゲーション プロパティ

プロパティには、**スカラー**と**ナビゲーション**という 2 つのカテゴリがあります。 スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。 ナビゲーションプロパティは、プロパティを直接代入するのではなく、読み取り専用のオブジェクトと、フィールドが割り当てられているオブジェクトのコレクションです。 たとえば、 `name` および `position` は、Excel の [Worksheet](/javascript/api/excel/excel.worksheet) オブジェクトのメンバーはスカラープロパティであり `protection` 、 `tables` ナビゲーションプロパティです。

アドインでは、ナビゲーションプロパティをパスとして使用して、特定のスカラープロパティを読み込むことができます。 次のコードでは、 `load` オブジェクトで使用されているフォントの名前に対するコマンドをキューに入れ `Excel.Range` ます。その他の情報は読み込まれません。

```js
someRange.load("format/font/name")
```

また、パスを通過してナビゲーションプロパティのスカラープロパティを設定することもできます。 たとえば、を使用してのフォントサイズを設定でき `Excel.Range` `someRange.format.font.size = 10;` ます。 プロパティを設定する前に、プロパティを読み込む必要はありません。

オブジェクトのプロパティの中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。 たとえば、 `format` はオブジェクトの下にあるプロパティですが、 `Excel.Range` `format` それ自体もオブジェクトです。 そのため、などの呼び出しを行うと `range.load("format")` 、これは `range.format.load()` (望ましくない empty ステートメント) と同じです `load()` 。 これを回避するには、コードでオブジェクトツリーの "葉 nodes" のみを読み込む必要があります。

#### <a name="calling-load-without-parameters-not-recommended"></a>`load`パラメーターを使用せずに呼び出す (推奨されません)

`load()`パラメーターを指定せずにオブジェクト (またはコレクション) に対してメソッドを呼び出すと、オブジェクトまたはコレクションのオブジェクトのすべてのスカラープロパティが読み込まれます。 不要なデータを読み込むと、アドインの速度が低下します。 読み込むプロパティを常に明示的に指定する必要があります。

> [!IMPORTANT]
> パラメーターのない `load` ステートメントで返されるデータの量は、サービスのサイズ制限を超える場合があります。 古いアドインのリスクを軽減するために、明示的に要求しない限り `load` によって返されないプロパティがあります。 次のプロパティは、そのような読み込み操作で除外されます。
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

プリミティブ型を返す promise ベースの api のメソッドには、パラダイムに似たパターンがあり `load` / `sync` ます。 たとえば、`Excel.TableCollection.getCount` はコレクション内のテーブルの数を取得します。 `getCount` を返し `ClientResult<number>` ます。これは、返されたプロパティが数値であることを意味 `value` [`ClientResult`](/javascript/api/office/officeextension.clientresult) します。 `context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。

次のコードは、Excel ブック内のテーブルの合計数を取得し、その番号をコンソールに記録します。

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

入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。 前述のナビゲーションパスを使用して個々のプロパティを設定する代わりに、 `object.set()` promise ベースの JavaScript api のオブジェクトで使用可能なメソッドを使用することができます。 このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。

次のコード サンプルは、`set()` メソッドを呼び出し、`Range`Range** オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲 **B2:E2 にデータがあると仮定します。

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

### <a name="some-properties-cannot-be-set-directly"></a>一部のプロパティは直接設定できません

書き込み可能であっても、一部のプロパティを設定することはできません。 これらのプロパティは、1つのオブジェクトとして設定する必要がある親プロパティの一部です。 これは、親プロパティが特定の論理的な関係を持つサブプロパティに依存しているためです。 このような親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクトリテラル表記を使用して設定し、オブジェクト全体を設定する必要があります。 この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。 このプロパティは、次 `zoom` に示すように、1つの [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) オブジェクトで設定する必要があります。

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

前の例では、値を直接割り当てることはでき***ません***。 `zoom` `sheet.pageLayout.zoom.scale = 200;` が読み込まれていないため、このステートメントはエラーをスロー `zoom` します。 `zoom`ロードされた場合でも、スケールのセットは有効になりません。 すべてのコンテキスト操作が行われ `zoom` 、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。

この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](application-specific-api-model.md#scalar-and-navigation-properties)とは異なります。 のプロパティは `format` 、次に示すように、object ナビゲーションを使用して設定できます。

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

読み取り専用修飾子をチェックすることによって、そのサブプロパティを直接設定できないプロパティを識別できます。 読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。 書き込み可能なプロパティ `PageLayout.zoom` は、そのレベルのオブジェクトで設定する必要があります。 概要:

- 読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。
- 書き込み可能なプロパティ: ナビゲーションを使用してサブプロパティを設定することはできません (最初の親オブジェクトの割り当ての一部として設定する必要があります)。



## <a name="42ornullobject-methods-and-properties"></a>&#42;OrNullObject メソッドとプロパティ

必要なオブジェクトが存在しない場合、いくつかのアクセサーメソッドとプロパティは例外をスローします。 たとえば、ブックにないワークシート名を指定して Excel ワークシートを取得しようとすると、 `getItem()` メソッドは例外をスロー `ItemNotFound` します。 アプリケーション固有のライブラリは、コードが、例外処理コードを必要とせずにドキュメントエンティティの存在をテストするための方法を提供します。 これは、 `*OrNullObject` メソッドとプロパティのバリエーションを使用して実現されます。 これらのバリエーション `isNullObject` は、指定したアイテムが `true` 存在しない場合、例外をスローするのではなく、プロパティがに設定されているオブジェクトを返します。

たとえば、 `getItemOrNullObject()` **ワークシート** などのコレクションに対してメソッドを呼び出して、コレクションからアイテムを取得することができます。 このメソッドは、指定されたアイテムが存在する場合はそれを返します。 `getItemOrNullObject()` それ以外の場合は、 `isNullObject` プロパティがに設定されているオブジェクトを返し `true` ます。 コードでは、このプロパティを評価して、オブジェクトが存在するかどうかを判断できます。

> [!NOTE]
> バリエーションは、 `*OrNullObject` JavaScript 値を返すことはありません `null` 。 通常の Office プロキシオブジェクトを返します。 オブジェクトが表すエンティティが存在しない場合、 `isNullObject` オブジェクトのプロパティはに設定され `true` ます。 全くまたは falsity に対して返されるオブジェクトはテストしません。 決して `null` 、 `false` 、、または `undefined` です。

次のコードサンプルでは、メソッドを使用して、"Data" という名前の Excel ワークシートを取得しようとして `getItemOrNullObject()` います。 その名前のワークシートが存在しない場合は、新しいシートが作成されます。 コードによってプロパティが読み込まれないことに注意して `isNullObject` ください。 Office は、が呼び出されたときに、このプロパティを自動的に読み込み `context.sync` ます。したがって、などのように明示的に読み込む必要はありません `datasheet.load('isNullObject')` 。

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

## <a name="see-also"></a>こちらもご覧ください

* [共通 JavaScript API オブジェクト モデル](office-javascript-api-object-model.md)
* [Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)
