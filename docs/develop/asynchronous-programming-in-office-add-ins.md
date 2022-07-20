---
title: Office アドインにおける非同期プログラミング
description: Office JavaScript ライブラリが Office アドインで非同期プログラミングを使用する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 64965d06544126584d7b17d078f4db9d464b39f0
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889500"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office アドインにおける非同期プログラミング

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office アドイン API で非同期プログラミングが使用されるのはなぜですか? JavaScript はシングルスレッドの言語であるため、スクリプトで実行時間の長い同期プロセスが呼び出されると、そのプロセスが完了するまで後続のすべてのスクリプト実行がブロックされます。 Office Web クライアント (ただしリッチ クライアント) に対する特定の操作は、同期的に実行される場合に実行をブロックする可能性があるため、Office JavaScript API のほとんどは非同期的に実行するように設計されています。 これにより、Office アドインの応答性と速度が向上します。 このような非同期メソッドを利用するときは、多くの場合、コールバック関数の記述も必要です。

API 内のすべての非同期メソッドの名前は、、メソッドなど`Document.getSelectedDataAsync``Binding.getDataAsync``Item.loadCustomPropertiesAsync`、"Async" で終わる。 "Async" メソッドは呼び出されるとすぐに実行され、後続のスクリプトも続けて実行することができます。 "Async" メソッドに渡す任意のコールバック関数は、データまたは要求された操作の準備が整い次第、すぐに実行されます。 コールバック関数の実行は通常、直ちに行われますが、戻るまでに若干の遅延が生じることがあります。

次の図は、サーバーベースの Word または Excel で開いているドキュメントでユーザーが選択したデータを読み取る "Async" メソッドの呼び出しの実行フローを示しています。 "Async" 呼び出しが行われる時点で、JavaScript 実行スレッドは追加のクライアント側処理を自由に実行できます (図には何も示されていません)。 "Async" メソッドが戻ると、コールバックはスレッドの実行を再開し、アドインはデータにアクセスし、何かを行い、結果を表示できます。 Word 2013 や Excel 2013 などの Office リッチ クライアント アプリケーションを操作する場合も、同じ非同期実行パターンが保持されます。

*図 1. 非同期プログラミング実行フロー*

![ユーザー、アドイン ページ、アドインをホストしている Web アプリ サーバーとの時間の経過に伴うコマンド実行操作を示す図。](../images/office-addins-asynchronous-programming-flow.png)

リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。

## <a name="write-the-callback-function-for-an-async-method"></a>"Async" メソッドのコールバック関数を記述する

コールバック引数として "Async" メソッドに渡す *コールバック* 関数は、コールバック関数の実行時に、アドイン ランタイムが [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトへのアクセスを提供するために使用する単一のパラメーターを宣言する必要があります。 次のように記述することができます。

- "Async" メソッドの *コールバック* パラメーターとして "Async" メソッドの呼び出しに沿って直接書き込み、渡す必要がある匿名関数。

- "Async" メソッドの *コールバック* パラメーターとしてその関数の名前を渡す名前付き関数。

匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。

### <a name="write-an-anonymous-callback-function"></a>匿名コールバック関数を記述する

次の匿名コールバック関数は、コールバックが戻ったときに [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) プロパティからデータを取得する名前の `result` 1 つのパラメーターを宣言します。

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

次の例は、この匿名コールバック関数を、メソッドへの完全な "Async" メソッド呼び出しのコンテキストで行に渡す方法を `Document.getSelectedDataAsync` 示しています。

- 最初の *coercionType* 引数は、 `Office.CoercionType.Text`選択したデータを文字列として返すように指定します。

- 2 番目の *コールバック* 引数は、メソッドにインラインで渡される匿名関数です。 関数が実行されると、*結果* パラメーターを使用してオブジェクトの`AsyncResult`プロパティにアクセス`value`し、ドキュメント内のユーザーが選択したデータを表示します。

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

コールバック関数のパラメーターを使用して、オブジェクトの他のプロパティに `AsyncResult` アクセスすることもできます。 呼び出しの成功または失敗を判断する場合は [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) プロパティを使用します。 呼び出しが失敗した場合は [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、エラーの詳細を確認できます。

メソッドの `getSelectedDataAsync` 使用の詳細については、「 [ドキュメントまたはスプレッドシートのアクティブな選択範囲に対するデータの読み取りと書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。

### <a name="write-a-named-callback-function"></a>名前付きコールバック関数を記述する

または、名前付き関数を記述し、その名前を "Async" メソッドの *コールバック* パラメーターに渡すこともできます。 たとえば、前の例は次のように `writeDataCallback` という名前の関数を *callback* パラメーターとして渡すように書き換えることができます。

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>AsyncResult.value プロパティに返される内容の違い

オブジェクトの `AsyncResult` 、`status`および`error`プロパティは`asyncContext`、すべての "Async" メソッドに渡されるコールバック関数に同じ種類の情報を返します。 ただし、プロパティに `AsyncResult.value` 返される内容は、"Async" メソッドの機能によって異なります。

たとえば、 `addHandlerAsync` ( [Binding](/javascript/api/office/office.binding)、 [CustomXmlPart](/javascript/api/office/office.customxmlpart)、 [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、 [Settings](/javascript/api/office/office.settings) オブジェクトの) メソッドを使用して、これらのオブジェクトで表される項目にイベント ハンドラー関数を追加します。 任意の`AsyncResult.value``addHandlerAsync`メソッドに渡すコールバック関数からプロパティにアクセスできますが、イベント ハンドラーを追加してもデータまたはオブジェクトにアクセスできないため、アクセスしようとすると、`value`プロパティは常に **未定義** を返します。

一方、メソッドを呼び出 `Document.getSelectedDataAsync` すと、ドキュメントでユーザーが選択したデータがコールバック内の `AsyncResult.value` プロパティに返されます。 または、 [Bindings.getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) メソッドを呼び出すと、ドキュメント内のすべてのオブジェクトの配列が `Binding` 返されます。 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) メソッドを呼び出すと、1 つの`Binding`オブジェクトが返されます。

メソッドのプロパティ`Async`に`AsyncResult.value`返される内容の説明については、そのメソッドの参照トピックの「コールバック値」セクションを参照してください。 メソッドを提供 `Async` するすべてのオブジェクトの概要については、 [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトトピックの下部にある表を参照してください。

## <a name="asynchronous-programming-patterns"></a>非同期プログラミング パターン

Office JavaScript API では、2 種類の非同期プログラミング パターンがサポートされています。

- 入れ子のコールバックの使用
- promise パターンの使用

コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。

入れ子のコールバックを使用することは、ほとんどの JavaScript 開発者にとってなじみのあるプログラミング パターンですが、コールバックが深い入れ子になっているコードは読みにくく、理解しにくいものです。 入れ子になったコールバックの代わりに、Office JavaScript API では promises パターンの実装もサポートされています。

> [!NOTE]
> 現在のバージョンの Office JavaScript API では、Promises パターン *の組み込みの* サポートは [、Excel スプレッドシートと Word ドキュメントのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)のコードでのみ機能します。 ただし、独自のカスタム Promise を返す関数内にコールバックを持つ他の関数をラップできます。 詳細については、「 [Promise を返す関数で共通 API をラップする」を](#wrap-common-apis-in-promise-returning-functions)参照してください。

### <a name="asynchronous-programming-using-nested-callback-functions"></a>入れ子のコールバック関数を使用する非同期プログラミング

多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。

次のコード例では、2 つの非同期呼び出しを入れ子にしています。

- 最初に、[Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) メソッドが呼び出され、"MyBinding" という名前のドキュメントのバインドにアクセスします。 そのコールバックのパラメーターに`result`返されるオブジェクトは`AsyncResult`、プロパティから`AsyncResult.value`指定されたバインディング オブジェクトにアクセスできます。
- 次に、最初 `result` のパラメーターからアクセスされたバインド オブジェクトを使用して [Binding.getDataAsync メソッドを](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) 呼び出します。
- 最後に、メソッドに `result2` 渡されたコールバックのパラメーターを `Binding.getDataAsync` 使用して、バインド内のデータを表示します。

```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この基本的な入れ子になったコールバック パターンは、Office JavaScript API のすべての非同期メソッドに使用できます。

次のセクションでは、非同期メソッドの入れ子のコールバックで匿名関数または名前付き関数を使用する方法を示します。

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>入れ子になったコールバックに匿名関数を使用する

次の例では、2 つの匿名関数がインラインで宣言され、入れ子になったコールバックとしてメソッドに`getDataAsync`渡`getByIdAsync`されます。 関数は単純でインラインのため、実装の意図は明白です。

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

#### <a name="use-named-functions-for-nested-callbacks"></a>入れ子になったコールバックに名前付き関数を使用する

複雑な実装の場合、名前付き関数を使用すると、読みやすく、保守管理がしやすく、再利用しやすくなります。 次の例では、前のセクションの例の 2 つの匿名関数が名前付`showResult`きの`deleteAllData`関数として書き換えられました。 これらの名前付き関数は、名前によってコールバックとしてメソッドに`deleteAllDataValuesAsync`渡されます`getByIdAsync`。

```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>promise パターンを使用してバインドのデータにアクセスする非同期プログラミング

コールバック関数を渡し、その関数が戻るのを待ってから実行を続行する代わりに、promise プログラミング パターンを使用すれば、その意図した結果を表す promise オブジェクトがすぐに返されます。ただし、本物の同期プログラミングとは異なり、実際には Office アドインのランタイム環境が要求を完了できるまでは、約束された結果の履行は実際には延期されます。要求が履行されない状況に対処するために *onError* ハンドラーが用意されています。

Office JavaScript API には、既存のバインド オブジェクトを操作するための Promises パターンをサポートする [Office.select](/javascript/api/office#Office_select_expression__callback_) メソッドが用意されています。 メソッドに `Office.select` 返される promise オブジェクトは、 [Binding](/javascript/api/office/office.binding) オブジェクトから直接アクセスできる 4 つのメソッド ( [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1))、 [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1))、 [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1))、 [removeHandlerAsync](/javascript/api/office/office.binding#office-office-binding-removehandlerasync-member(1))) のみをサポートします。

バインドを操作するための Promises パターンは、この形式になります。

**Office.select(**_selectorExpression_, _onError_**)。**_BindingObjectAsyncMethod_

*selectorExpression* パラメーターはフォーム`"bindings#bindingId"`を受け取ります。*bindingId* は、ドキュメントまたはスプレッドシートで以前に作成したバインディングの名前 ( ) `id`です (コレクションの "addFrom" メソッド`Bindings`のいずれかを使用します。 `addFromNamedItemAsync``addFromPromptAsync``addFromSelectionAsync` たとえば、セレクター式 `bindings#cities` では、"cities" **の ID** を持つバインドにアクセスすることを指定します。

*onError* パラメーターは、指定したバインディングへのアクセスに失敗した場合`select`に、オブジェクトへのアクセス`Error`に使用できる型`AsyncResult`の 1 つのパラメーターを受け取るエラー処理関数です。 次の例は、*onError* パラメーターに渡すことができる基本的なエラー処理関数を示しています。

```js
function onError(result){
    const err = result.error;
    write(err.name + ": " + err.message);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

*BindingObjectAsyncMethod* プレースホルダーを promise オブジェクトでサポートされている 4 つの`Binding`オブジェクト メソッドの呼び出しに置き換えます。 `addHandlerAsync``getDataAsync``setDataAsync``removeHandlerAsync` これらのメソッドの呼び出しでは追加の promise がサポートされません。 これらは[入れ子のコールバック関数パターン](#asynchronous-programming-using-nested-callback-functions)を使用して呼び出す必要があります。

`Binding`オブジェクトの Promise が満たされた後は、連結されたメソッド呼び出しで、バインドであるかのように再利用できます (アドイン ランタイムは、Promise の実行を非同期的に再試行しません)。 オブジェクト Promise を `Binding` 満たすことができない場合、アドイン ランタイムは、次に非同期メソッドの 1 つが呼び出されるときに、バインド オブジェクトへのアクセスを再試行します。

次のコード例では、メソッドを`select`使用してコレクションから "`cities`" を持つバインドを`id``Bindings`取得し、[addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) メソッドを呼び出して、バインディングの [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントのイベント ハンドラーを追加します。

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> `Binding`メソッドによって`Office.select`返されるオブジェクト promise は、オブジェクトの 4 つのメソッド`Binding`にのみアクセスできます。 オブジェクトの他のメンバー`Binding`にアクセスする必要がある場合は、代わりにプロパティまたは`Bindings.getByIdAsync``Bindings.getAllAsync`メソッドを`Document.bindings`使用してオブジェクトを`Binding`取得する必要があります。 たとえば、オブジェクトの`Binding`プロパティ (`document`、`id`、またはプロパティ) にアクセスする必要がある場合、または [MatrixBinding オブジェクトまたは](/javascript/api/office/office.matrixbinding) `type` [TableBinding](/javascript/api/office/office.tablebinding) オブジェクトのプロパティにアクセスする必要がある場合は、オブジェクトを取得`Binding`するために、または`getAllAsync`メソッドを`getByIdAsync`使用する必要があります。

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>省略可能なパラメーターを非同期メソッドに渡す

すべての "Async" メソッドの一般的な構文は、このパターンに従います。

 *AsyncMethod* `(`*RequiredParameters*`, [`*OptionalParameters*`],`*CallbackFunction*`);`

すべての非同期メソッドは省略可能なパラメーターをサポートします。これは、1 つ以上の省略可能なパラメーターを含む JavaScript オブジェクトとして渡されます。 省略可能なパラメーターを含むオブジェクトは、キーと値を区切る ":" 文字を持つキーと値のペアの順序なしコレクションです。 オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体が中かっこで囲まれています。 キーはパラメーター名で、値はそのパラメーターに渡す値です。

省略可能なパラメーターを含むオブジェクトをインラインで作成することも、オブジェクトを `options` 作成して *options* パラメーターとして渡すことで作成することもできます。

### <a name="pass-optional-parameters-inline"></a>省略可能なパラメーターをインラインで渡す

たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) メソッドを呼び出す場合の構文は、次のようになります。

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

この形式の呼び出し構文では、2 つの省略可能なパラメーター *(coercionType* と *asyncContext*) は、中かっこで囲まれた匿名 JavaScript オブジェクトとして定義されます。

次の例は、省略可能なパラメーターをインラインで `Document.setSelectedDataAsync` 指定してメソッドを呼び出す方法を示しています。

```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> 省略可能なパラメーターは、パラメーター オブジェクトの名前が正しく指定されていれば、任意の順序で指定できます。

### <a name="pass-optional-parameters-in-an-options-object"></a>オプション オブジェクトに省略可能なパラメーターを渡す

または、メソッド呼び出しとは別に省略可能なパラメーターを指定する名前`options`のオブジェクトを作成し、*その* オブジェクトを `options` options 引数として渡すこともできます。

次の例では、オブジェクトを作成する方法の 1 つを`options`示します。ここでは`parameter1``value1`、実際のパラメーター名と値のプレースホルダーを示します。

```js
const options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};
```

[ValueFormat](/javascript/api/office/office.valueformat) パラメーターおよび [FilterType](/javascript/api/office/office.filtertype) パラメーターを指定する場合は次のようになります。

```js
const options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

オブジェクトを作成する別の方法を次に示 `options` します。

```js
const options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

パラメーターの`FilterType`指定に使用すると、次の`ValueFormat`例のようになります。

```js
const options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> オブジェクトを作成 `options` するどちらのメソッドを使用する場合でも、名前が正しく指定されていれば、任意の順序で省略可能なパラメーターを指定できます。

次の例では、オブジェクトで省略可能なパラメーターを `Document.setSelectedDataAsync` 指定してメソッドを呼び出す方法を `options` 示します。

```js
const options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

どちらのオプションのパラメーター例でも、*callback* パラメーターが最後のパラメーターとして (インラインのオプションのパラメーターまたは *options* 引数オブジェクトに続けて) 指定されています。 または、インライン JavaScript オブジェクト内またはオブジェクト内で *コールバック* パラメーターを `options` 指定することもできます。 ただし、 *コールバック* パラメーターは、オブジェクト内 `options` (インラインまたは外部で作成)、または最後のパラメーターとして渡すことができますが、両方とも渡すことができません。

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Promise を返す関数で共通 API をラップする

Common API (および Outlook API) メソッドは [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返しません。 そのため、非同期操作が完了するまで [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) を使用して実行を一時停止することはできません。 動作が必要な `await` 場合は、明示的に作成された Promise でメソッド呼び出しをラップできます。

基本的なパターンは、Promise オブジェクトをすぐに返し、内部メソッドが完了したときに Promise オブジェクトを *解決* する非同期メソッドを作成することです。メソッドが失敗した場合はオブジェクトを *拒否* します。 次に簡単な例を示します。

```javascript
function getDocumentFilePath() {
    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                resolve(asyncResult.value.url);
            });
        }
        catch (error) {
            reject(WordMarkdownConversion.errorHandler(error));
        }
    })
}
```

このメソッドを待機する必要がある場合は、キーワードを使用するか、関数に `await` 渡された `then` 関数として呼び出すことができます。

> [!NOTE]
> この手法は、アプリケーション固有のオブジェクト モデルの 1 つのメソッドの呼び出し内で Common API の 1 つを `run` 呼び出す必要がある場合に特に便利です。 この方法で使用されている上記の関数の例については、 [ サンプル Word-Add-in-JavaScript-MDConversion のファイルHome.js](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)を参照してください。

TypeScript を使用する例を次に示します。

```typescript
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
