---
title: Excel JavaScript API を使用した基本的なプログラミングの概念
description: Excel JavaScript APIを使用して、Excel 用アドインを構築します。
ms.date: 10/03/2018
ms.openlocfilehash: f93ec7b5e34f90f2d61f29d861b7e0c19f66f6e3
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505987"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API を使用した基本的なプログラミングの概念
 
この資料では、 [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) を使用してアドイン を、Excel 2016 またはそれ以降にビルドする方法について説明します。API を使用する上で基本となる、読み取りまたは書き込み、広い範囲、範囲、およびその他のすべてのセルを更新するなどの特定のタスクを実行するためのガイダンスを提供する主要な概念が導入されています。

## <a name="asynchronous-nature-of-excel-apis"></a>Excel API の非同期性

Office for Windows などのデスクトップ ベースのプラットフォーム上の Office アプリケーション内で埋め込まれ、Office オンラインの HTML iFrame 内で動作するブラウザーのコンテナー内で実行して、web を使用した Excel のアドインを使用します。Office.js の API がサポートされているすべてのプラットフォーム間で、Excel ホストと同期的に対話を有効にすることは、パフォーマンスに関する考慮事項のため不可能です。したがって、Office.jsの **sync()** API 呼び出しは、Excel アプリケーションが要求された読み取りまたは書き込みアクションを完了したときに解決される [約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。また、操作ごとに別の要求を送信するのではなく、プロパティを設定またはメソッドを呼び出すなど、複数の操作のキューして、 **sync()** への 1 回の呼び出しで指定されたコマンドのバッチとして実行できます。次のセクションでは、 **Excel.run()** と **sync()** API を使用してこれを実現する方法について説明します。
 
## <a name="excelrun"></a>Excel.run
 
**Excel.run** では、Excel オブジェクト モデルに対して実行するアクションを指定する関数を実行します。 **Excel.run** は、自動的に Excel のオブジェクトと対話するために使用できる要求のコンテキストを作成します。 **Excel.run** が完了して、約束を解決すると、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。
 
**Excel.run**を使用する例を次に示します。Catch ステートメントでは、 **Excel.run**で発生したエラーキャッチし、ログに記録します。
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="request-context"></a>要求コンテキスト
 
Excel とユーザーのアドインは、2 つの異なるプロセスで実行されます。それらは異なるランタイム環境を使用するため、Excel アドインでは、ワークシート、範囲、グラフ、表など、Excel のオブジェクトにユーザーのアドインを接続するために **RequestContext** オブジェクトが必要です。
 
## <a name="proxy-objects"></a>プロキシ オブジェクト
 
宣言して、アドインで使用する Excel の JavaScript オブジェクトは、プロキシ オブジェクトです。起動するメソッドまたはプロパティを設定するか、プロキシ オブジェクトにロードするだけで、保留中のコマンドのキューに追加されます。要求のコンテキストに **sync()** メソッドを呼び出すとき (たとえば、 `context.sync()`)、キュー内のコマンドを Excel にディスパッチし、実行します。Excel JavaScript API では、基本的にバッチを中心としました。要求のコンテキストにメソッドを呼び出して、 **sync()** をキューに登録されたコマンドのバッチを実行すると、多くの変更をキューにできます。
 
たとえば、次のコード スニペットでは、ローカル JavaScript オブジェクト **selectedRange** が Excel ドキュメント内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。**selectedRange** オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが**context.sync()** を呼び出すまで Excel ドキュメントには反映されません。
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>sync()
 
要求コンテキストで **sync()** メソッドを呼び出すと、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態が同期されます。** sync()** メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。 **Sync()** メソッドでは、非同期的に実行し、**sync()** メソッドの完了時に解決する[約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。
 
次の例は、ローカル JavaScript proxy オブジェクト (**selectedRange**) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して **context.sync()** を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
前の例では、**selectedRange** が設定され、**context.sync()** が呼び出されるとその **address** プロパティが読み込まれます。
 
**sync()** は 約束 を返す非同期の操作であるため、常に 約束を (JavaScript で) **返す**必要があります。これにより、スクリプトは実行を継続する前に **sync()** 操作が完了しています。** sync()** を用いたパフォーマンスの最適化の詳細については、「[   Excel JavaScript API のパフォーマンス最適化](https://docs.microsoft.com/office/dev/add-ins/excel/performance)」を参照してください。
 
### <a name="load"></a>load()
 
プロキシ オブジェクトのプロパティを読み取るには、まず Excel ドキュメントからプロキシ オブジェクトとデータを入力するプロパティを明示的に読み込み、それから **context.sync()** を呼び出す必要があります。たとえば、選択範囲を参照するプロキシ オブジェクトを作成した後、選択範囲の**  address** プロパティを読み取る必要がある場合、読み取る前に** address** プロパティを読み込む必要があります。読み込むプロキシ オブジェクトのプロパティを要求するには、オブジェクトの **load()** メソッドを呼び出し、ロードするプロパティを指定します。 

> [!NOTE]
> プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、**load()** メソッドを呼び出す必要はありません。**load()** メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。
 
プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 **sync()** メソッドを呼び出すときに実行されます。**load()** の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。
 
次の例では、範囲の特定のプロパティのみが読み込まれます。
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
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
 
前の例では、`format/font` が **myRange.load()** の呼び出しで指定されていないため、`format.font.color` プロパティは読み取れませんでした。

パフォーマンスを最適化するにはプロパティと [Excel JavaScript API のパフォーマンスの最適化](performance.md) で説明したように、オブジェクトの **load()** メソッドを使用する場合の読み込みに関係を明示的に指定する必要があります。 **Load()** メソッドの詳細については、 [Excel JavaScript API を使用して高度なプログラミングの概念](excel-add-ins-advanced-concepts.md)を参照してください。

## <a name="null-or-blank-property-values"></a>null または空白のプロパティ値
 
### <a name="null-input-in-2-d-array"></a>2 次元配列での null の入力
 
Excel では、範囲は、最初の次元が行と2 番目の次元が列である、2 次元配列で表されます。値、数値の書式、または範囲内で特定のセルの数式を設定するに、2 次元配列の値、数値形式、またはそれらのセルの数式を指定して、 `null` 2 次元配列内のすべてのセルにします。
 
たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>プロパティに対する null の入力
 
`null` を、単独プロパティに対する有効な入力として指定することはできません。たとえば、次のコード スニペットは、範囲の **values** プロパティを `null` に設定できないため無効です。
 
```js
range.values = null;
```
 
同様に、次のコード スニペットは、`null` が **color** プロパティで有効ではないため無効です。
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>応答内の null プロパティ値
 
指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。たとえば、範囲を取得してその`format.font.color`   プロパティを読み込む場合:
 
* 範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。
* 範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。
 
### <a name="blank-input-for-a-property"></a>プロパティに対する空白の入力
 
プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:
 
* 範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。
 
* `numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。
 
* `formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。
 
### <a name="blank-property-values-in-the-response"></a>応答内の空白のプロパティ値
 
応答での空白のプロパティ値の読み取り操作は、(つまり、スペースを入れない `''`の間の2 つの引用符) そのセルが含まれていないデータまたは値を示します。最初次の例で、最初と最後のセル範囲のデータは含まれません。2 番目の例では、範囲内の最初の 2 つのセルには、数式が入力されません。
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>無制限の範囲への読み取りまたは書き込み
 
### <a name="read-an-unbounded-range"></a>無制限の範囲の読み取り
 
無制限の範囲のアドレスとは、列全体または行全体を指定する範囲のアドレスです。例:
 
* 範囲のアドレスは、列全体で構成されます。<ul><li>`C:C`</li><li>`A:F`</li></ul>
* 範囲のアドレスは、行全体で構成されます。<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')`)、返される応答には、`values`、`text`、`numberFormat`、または `formula` などのセル レベルのプロパティに `null` が含まれます。`address`、または `cellCount` などのその他の範囲プロパティは、無制限の範囲を反映します。
 
### <a name="write-to-an-unbounded-range"></a>無制限の範囲への書き込み
 
セル レベルのプロパティを非制限の範囲で次のように設定することはできません `values`、 `numberFormat`、および `formula` 。これは入力の要求が大きすぎるためです。たとえば、次のコード スニペットは、無限の範囲の `values` を指定しようとするので、無効です、無限の範囲のセル レベルのプロパティを設定しようとした場合、API はエラーを返します。
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>広い範囲に対する読み取りまたは書き込み
 
範囲に多数のセル、値、数値の書式、数式が含まれている場合は、その範囲に API の操作を実行することはできません。範囲上で要求された操作を実行するため、API の最適な試行を必ず確認 (つまり、指定したデータの取得または書き込み)しますが、広い範囲での 読み取りまたは書き込み操作の実行を試みると、過剰なリソース使用率による API エラーが発生します。このようなエラーを避けるためには、別の読み取りを実行するか、1 つを実行しようとしてではなく、広い範囲の小さなサブセットを操作の読み取りまたは書き込み操作に大きな範囲での書き込みをお勧めします。
 
## <a name="update-all-cells-in-a-range"></a>範囲内のすべてのセルの更新
 
範囲内のすべてのセルに同じ更新 (すべてのセルに同じ値を入力する、同じ数値書式を設定する、同じ数式ですべてのセルにデータを入力するなど) を適用するには、**range** オブジェクトの該当するプロパティを必要な 1 つの値に設定します。
 
次の例では、20 個のセルを含む範囲を取得し、数値書式を設定してその範囲のすべてのセルに **3/11/2015** という値を設定します。
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a>エラー メッセージ
 
API エラーが発生すると、API ではコードとメッセージを含む **error** オブジェクトが返されます。次の表は、API から返されるエラー一覧の定義を示します。
 
|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引数が無効であるか、存在しません。または形式が正しくありません。|
|InvalidRequest  |要求を処理できません。|
|InvalidReference|この参照は、現在の操作に対して無効です。|
|InvalidBinding  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|InvalidSelection|現在の選択内容は、この操作では無効です。|
|認証されていません |必要な認証情報が見つからないか、無効です。|
|AccessDenied |要求された操作を実行できません。|
|ItemNotFound |要求されたリソースは存在しません。|
|ActivityLimitReached|アクティビティの制限に達しました。|
|GeneralException|リクエストの処理中に内部エラーが発生しました。|
|NotImplemented  |リクエストされた機能は実装されていません。|
|ServiceNotAvailable|サービスを利用できません。|
|一致しません              |競合のため、要求を処理できませんでした。|
|ItemAlreadyExists|作成中のリソースはすでに存在しています。|
|UnsupportedOperation|試行中の操作はサポートされていません。|
|RequestAborted|実行時に要求が中止されました。|
|ApiNotAvailable|要求された API は使用できません。|
|InsertDeleteConflict|試行された挿入操作または削除操作で競合が発生しました。|
|InvalidOperation|試行された操作は、このオブジェクトでは無効です。|
 
## <a name="see-also"></a>関連項目
 
* [Excel アドインを使う](excel-add-ins-get-started-overview.md)
* [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [Excel JavaScript API を使用した高度なプログラミングの概念](excel-add-ins-advanced-concepts.md)
* [Excel JavaScript API パフォーマンスの最適化](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [Excel JavaScript API リファレンス](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
