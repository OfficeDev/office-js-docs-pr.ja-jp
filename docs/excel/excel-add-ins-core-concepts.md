# <a name="excel-javascript-api-core-concepts"></a>Excel JavaScript API の中心概念
 
この記事では、[Excel JavaScript API を使用して Excel 2016](../../reference/excel/excel-add-ins-reference-overview.md) のアドインをビルドする方法について説明します。 ここでは API の使用の基本となる中心概念について説明し、広い範囲に対する読み取り、書き込み、一定範囲内すべてのセルの更新など、特定のタスクを実行するためのガイダンスを提供します。

## <a name="asynchronous-nature-of-excel-apis"></a>Excel API の非同期性

Web ベースの Excel アドインは、Office for Windows などのデスクトップ ベースのプラットフォーム上にある Office アプリケーションに組み込まれ、Office Online の HTML iFrame 内で実行されるブラウザー コンテナー内で実行されます。 サポートされているすべてのプラットフォームで Office.js API が Excel ホストと同期的に対話することは、パフォーマンスの観点からうまくいきません。 このため、Office.js 内の **sync()** API の呼び出しにより [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) が返され、それは Excel アプリケーションが要求された読み取りまたは書き込み操作を完了したときに解決されます。 また、操作ごとに別個の要求として送信する代わりに、プロパティの設定やメソッドの起動など、複数の操作をキューに登録し、**sync()** の 1 回の呼び出しでコマンドのバッチとしてそれらを実行することもできます。 次のセクションでは、**Excel.run()** と **sync()** API を使用してこれを実行する方法について説明します。
 
## <a name="excelrun"></a>Excel.run
 
**Excel.run** は Excel オブジェクト モデルに対して実行する操作を指定した関数を実行します。 **Excel.run** は Excel オブジェクトと対話するために使用できる要求コンテキストを自動的に作成します。 **Excel.run**が完了すると、Promose が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。
 
次の例は、**Excel.run** の使用方法を示しています。 Catch ステートメントは **Excel.run** 内で発生するエラーをキャッチし、ログに記録します。
 
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
 
Excel アドインとユーザーのアドインは、2 つの異なるプロセスで実行されます。 これらはさまざまなランタイム環境を使用するため、Excel アドインでは、シート、範囲、グラフ、表など、Excel のオブジェクトにユーザーのアドインを接続するために **RequestContext** オブジェクトが必要です。
 
## <a name="proxy-objects"></a>プロキシ オブジェクト
 
アドインで宣言し、使用する Excel JavaScript オブジェクトはプロキシ オブジェクトです。 起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。 `context.sync()` など、要求コンテキスト上で **sync()** メソッドを呼び出すと、キューに入れられたコマンドは Excel にディスパッチされて実行されます。 Excel の JavaScript API では、基本的にバッチを中心としています。 要求コンテキストに必要なだけ変更内容をキューに登録し、**sync()** メソッドを呼び出して、キューに入れられたコマンドをバッチで実行することができます。
 
たとえば、次のコード スニペットでは、ローカル JavaScript オブジェクト **selectedRange** が Excel ドキュメント内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。 **selectedRange** オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが **context.sync()** を呼び出すまで Excel ドキュメントには反映されません。
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a>sync()
 
要求コンテキストで **sync()** メソッドを呼び出すと、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態が同期されます。 **sync()** メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。 **sync()** メソッドは非同期で実行されて [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。これは、**sync()** メソッドが完了すると解決されます。
 
>**注**:Excel JavaScript API では、**sync()** が唯一の非同期操作です。 パフォーマンスを最適化するには、**sync()** を呼び出す前にできる限り変更内容をキューに登録し、**sync()** を呼び出す回数を最小限に抑えます。
 
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
 
**sync()** は Promise を返す非同期の操作であるため、常に Promise を (JavaScript で) **返す**必要があります。 このようにして、スクリプトの実行を継続する前に、**sync()** 操作を完了します。
 
### <a name="load"></a>load()
 
プロキシ オブジェクトのプロパティを読み取るには、まず Excel ドキュメントからプロキシ オブジェクトとデータを入力するプロパティを明示的に読み込み、それから **context.sync()** を呼び出す必要があります。 たとえば、選択範囲を参照するプロキシ オブジェクトを作成した後、選択範囲の **address** プロパティを読み取る必要がある場合、読み取る前に **address** プロパティを読み込む必要があります。 プロキシ オブジェクトのプロパティを読み込むよう要求するには、オブジェクトに対して **load()** メソッドを呼び出し、読み込むプロパティを指定します。 

>**注**:プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、**load()** メソッドを呼び出す必要はありません。 **load()** メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。
 
プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中コマンドのキューに追加され、次回 **sync()** メソッドを呼び出すときに実行されます。 **load()** の呼び出しは、必要なだけ要求コンテキストのキューに登録できます。
 
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
 
パフォーマンスを最適化するため、オブジェクトに対して **load()** メソッドを使用するときに読み込むプロパティとリレーションシップを明示的に指定する必要があります。 たとえば、範囲オブジェクトの **address** プロパティのみを読み取る場合、**load()** メソッドを呼び出すときにそのプロパティのみを指定します。
 
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
 
* `properties` は、読み込まれるプロパティ名やリレーションシップ名の一覧で、名前のコンマ区切りの文字列または配列として指定されます。 詳細については、「[Excel JavaScript API リファレンス](../../reference/excel/excel-add-ins-reference-overview.md)」でオブジェクトに対して定義されている **load()** メソッドを参照してください。
* `loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](../../reference/excel/loadoption.md)を参照してください。

**load()** メソッドの詳細は、「[Excel JavaScript API の高度な概念](excel-add-ins-advanced-concepts.md)」を参照してください。

## <a name="null-or-blank-property-values"></a>null または空白のプロパティ値
 
### <a name="null-input-in-2-d-array"></a>2 次元配列での null の入力
 
Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。 範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。
 
たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。 次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a>プロパティに対する null の入力
 
`null` は単一プロパティに有効な入力ではありません。 たとえば、次のコード スニペットは、範囲の **values** プロパティを `null` に設定できないため無効です。
 
```js
range.values = null;
```
 
同様に、次のコード スニペットは、`null` が **color** プロパティで有効ではないため無効です。
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a>応答内の null プロパティ値
 
指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。 たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:
 
* 範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。
* 範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。
 
### <a name="blank-input-for-a-property"></a>プロパティに対する空白の入力
 
プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。 例:
 
* 範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。
 
* `numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。
 
* `formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。
 
### <a name="blank-property-values-in-the-response"></a>応答内の空白のプロパティ値
 
読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。 次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。 2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a>無制限の範囲への読み取りまたは書き込み
 
### <a name="read-an-unbounded-range"></a>無制限の範囲の読み取り
 
無制限の範囲のアドレスとは、列全体または行全体を指定する範囲のアドレスです。 例:
 
* 範囲のアドレスは、列全体で構成されます。<ul><li>`C:C`</li><li>`A:F`</li></ul>
* 範囲のアドレスは、行全体で構成されます。<ul><li>`2:2`</li><li>`1:4`</li></ul>
 
API が無制限の範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`values`、`text`、`numberFormat`、または `formula` などのセル レベルのプロパティに `null` 値が含まれます。 `address` または `cellCount` など、範囲のその他のプロパティには、無制限の範囲に有効な値が含まれます。
 
### <a name="write-to-an-unbounded-range"></a>無制限の範囲への書き込み
 
無制限の範囲では、入力要求が大きすぎるため、`values`、`numberFormat`、`formula` などのセル レベルのプロパティは設定できません。 たとえば、次のコード スニペットは、無制限の範囲に対して `values` を指定しようとしているため無効です。 無制限の範囲にセル レベルのプロパティを設定しようとすると、API からエラーが返されます。
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a>広い範囲に対する読み取りまたは書き込み
 
範囲に多数のセル、値、数値書式、数式などが含まれる場合、その範囲では API 操作を実行できない場合があります。 API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。 このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。
 
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
 
API エラーが発生すると、API ではコードとメッセージを含む **error** オブジェクトが返されます。 次の表は、API から返されるエラー一覧の定義を示します。
 
|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引数が無効であるか、存在しません。または形式が正しくありません。|
|InvalidRequest  |要求を処理できません。|
|InvalidReference|この参照は、現在の操作に対して無効です。|
|InvalidBinding  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|InvalidSelection|現在の選択内容は、この操作では無効です。|
|Unauthenticated |必要な認証情報が見つからないか、無効です。|
|AccessDenied |要求された操作を実行できません。|
|ItemNotFound |要求されたリソースは存在しません。|
|ActivityLimitReached|アクティビティの制限に達しました。|
|GeneralException|要求の処理中に内部エラーが発生しました。|
|NotImplemented  |要求された機能は実装されていません。|
|ServiceNotAvailable|サービスを利用できません。|
|Conflict              |競合のため、要求を処理できませんでした。|
|ItemAlreadyExists|作成中のリソースはすでに存在しています。|
|UnsupportedOperation|試行中の操作はサポートされていません。|
|RequestAborted|実行時に要求が中止されました。|
|ApiNotAvailable|要求された API は使用できません。|
|InsertDeleteConflict|試行された挿入操作または削除操作で競合が発生しました。|
|InvalidOperation|試行された操作はこのオブジェクトでは無効です。|
 
## <a name="additional-resources"></a>その他のリソース
 
* [Excel アドインを使う](excel-add-ins-get-started-overview.md)
* [Excel アドインのコード サンプル](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API リファレンス](../../reference/excel/excel-add-ins-reference-overview.md)
