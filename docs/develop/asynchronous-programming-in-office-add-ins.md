---
title: Office アドインにおける非同期プログラミング
description: JavaScript ライブラリがOfficeアドインで非同期プログラミングを使用する方法Office説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1663f15d1b9f4191fc1f0c21f0532b5e23fdade6
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671388"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office アドインにおける非同期プログラミング

[!include[information about the common API](../includes/alert-common-api-info.md)]

アドイン API Office非同期プログラミングを使用する理由 JavaScript はシングルスレッドの言語であるため、スクリプトで実行時間の長い同期プロセスが呼び出されると、そのプロセスが完了するまで後続のすべてのスクリプト実行がブロックされます。 Office Web クライアント (ただしリッチ クライアントも含む) に対する特定の操作は、同期的に実行される場合に実行をブロックする可能性があるため、Office JavaScript API の大部分は非同期的に実行するように設計されています。 これにより、アドインOffice応答性と速度が向上します。 このような非同期メソッドを利用するときは、多くの場合、コールバック関数の記述も必要です。

API 内のすべての非同期メソッドの名前は、"Async" (、 、メソッドなど) `Document.getSelectedDataAsync` `Binding.getDataAsync` で `Item.loadCustomPropertiesAsync` 終わる。 "Async" メソッドは呼び出されるとすぐに実行され、後続のスクリプトも続けて実行することができます。 "Async" メソッドに渡す任意のコールバック関数は、データまたは要求された操作の準備が整い次第、すぐに実行されます。 コールバック関数の実行は通常、直ちに行われますが、戻るまでに若干の遅延が生じることがあります。

次の図は、サーバー ベースの Word または Excel で開いているドキュメントでユーザーが選択したデータを読み取る "Async" メソッドへの呼び出しの実行フローを示しています。 "Async" 呼び出しが行われた時点で、JavaScript 実行スレッドは追加のクライアント側処理を自由に実行できます (ただし、図には何も表示されません)。 "Async" メソッドが返されると、コールバックはスレッドでの実行を再開し、アドインはデータにアクセスし、データにアクセスして何かを行い、結果を表示できます。 同じ非同期実行パターンは、Word 2013 や 2013 Officeなど、リッチ クライアント アプリケーションを操作するときに保持Excelします。

*図 1. 非同期プログラミング実行フロー*

![ユーザー、アドイン ページ、およびアドインをホストする Web アプリ サーバーとの時間の間のコマンド実行の相互作用を示す図。](../images/office-addins-asynchronous-programming-flow.png)

リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。

## <a name="writing-the-callback-function-for-an-async-method"></a>"Async" メソッドのコールバック関数を記述する

コールバック引数として "Async" メソッドに渡すコールバック関数は、コールバック関数の実行時にアドイン ランタイムが[AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトへのアクセスを提供するために使用する 1 つのパラメーターを宣言する必要があります。  次のように記述することができます。

- "Async" メソッドのコールバック パラメーターとして"Async" メソッドへの呼び出しに沿って直接書き込み、渡す必要がある匿名関数。

- "Async" メソッドのコールバック パラメーターとしてその関数の名前を渡す名前付き関数。

匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。

### <a name="writing-an-anonymous-callback-function"></a>匿名コールバック関数を記述する

次の匿名コールバック関数は、コールバックが返されたときに `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティからデータを取得するという名前の単一のパラメーターを宣言します。

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

次の例は、完全な "Async" メソッド呼び出しのコンテキストで、この匿名コールバック関数を行に渡す方法を示 `Document.getSelectedDataAsync` しています。

- 最初の _coercionType_ 引数 , は、選択したデータを文字列として `Office.CoercionType.Text` 返す値を指定します。

- 2 _番目のコールバック_ 引数は、メソッドに行で渡される匿名関数です。 関数を実行すると _、result_ パラメーターを使用してオブジェクトのプロパティにアクセスし、ユーザーが選択したデータをドキュメント `value` `AsyncResult` に表示します。

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

コールバック関数のパラメーターを使用して、オブジェクトの他のプロパティにアクセス `AsyncResult` することもできます。 呼び出しの成功または失敗を判断する場合は [AsyncResult.status](/javascript/api/office/office.asyncresult#status) プロパティを使用します。 呼び出しが失敗した場合は [AsyncResult.error](/javascript/api/office/office.asyncresult#error) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、エラーの詳細を確認できます。

メソッドの使用の詳細については、「ドキュメントまたはスプレッドシート内のアクティブな選択範囲に対するデータの読み取りおよび書き `getSelectedDataAsync` [込み」を参照してください](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。 

### <a name="writing-a-named-callback-function"></a>名前付き関数を記述する

または、名前付き関数を記述し、その名前を "Async" メソッドの _コールバック_ パラメーターに渡します。 たとえば、前の例は次のように `writeDataCallback` という名前の関数を _callback_ パラメーターとして渡すように書き換えることができます。

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

オブジェクトの 、 、 プロパティは、すべての "Async" メソッドに渡されるコールバック関数に同じ種類の情報 `asyncContext` `status` `error` `AsyncResult` を返します。 ただし、プロパティに返される内容は `AsyncResult.value` 、"Async" メソッドの機能によって異なります。

たとえば、これらのオブジェクトで表されるアイテムにイベント ハンドラー関数を追加するには、メソッド `addHandlerAsync` (Binding、CustomXmlPart、Document、RoamingSettings、および[設定](/javascript/api/office/office.settings)オブジェクト) を使用[](/javascript/api/outlook/office.roamingsettings)[](/javascript/api/office/office.document)[](/javascript/api/office/office.binding)[](/javascript/api/office/office.customxmlpart)します。 任意のメソッドに渡すコールバック関数からプロパティにアクセスできますが、イベント ハンドラーを追加するときにデータやオブジェクトがアクセスされていないので、アクセスしようとすると、プロパティは常に未定義を返します。 `AsyncResult.value` `addHandlerAsync` `value` 

一方、メソッドを呼び出した場合は、ドキュメントで選択したユーザーのデータをコールバック `Document.getSelectedDataAsync` `AsyncResult.value` 内のプロパティに返します。 または [、Bindings.getAllAsync](/javascript/api/office/office.bindings#getAllAsync_options__callback_) メソッドを呼び出した場合は、ドキュメント内のすべてのオブジェクトの配列 `Binding` を返します。 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_)メソッドを呼び出した場合は、1 つのオブジェクトを返 `Binding` します。

メソッドのプロパティに返される内容の説明については、そのメソッドのリファレンス トピックの `AsyncResult.value` `Async` 「Callback value」セクションを参照してください。 メソッドを提供するオブジェクトの概要については `Async` [、AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトのトピックの下部にある表を参照してください。

## <a name="asynchronous-programming-patterns"></a>非同期プログラミング パターン

JavaScript API Officeは、次の 2 種類の非同期プログラミング パターンをサポートしています。

- 入れ子のコールバックの使用
- promise パターンの使用

コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。

入れ子のコールバックを使用することは、ほとんどの JavaScript 開発者にとってなじみのあるプログラミング パターンですが、コールバックが深い入れ子になっているコードは読みにくく、理解しにくいものです。 入れ子になったコールバックの代わりに、javaScript API Office約束パターンの実装もサポートしています。

> [!NOTE]
> Office JavaScript API の現在のバージョンでは、promises パターンの組み込みサポートは、Excel スプレッドシートと Word ドキュメントのバインド[のコードでのみ機能します](bind-to-regions-in-a-document-or-spreadsheet.md)。  ただし、独自のカスタム Promise-returning 関数内にコールバックを持つ他の関数をラップできます。 詳細については [、「Promise-returning 関数で一般的な API をラップする」を参照してください](#wrap-common-apis-in-promise-returning-functions)。

### <a name="asynchronous-programming-using-nested-callback-functions"></a>入れ子のコールバック関数を使用する非同期プログラミング

多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。

次のコード例では、2 つの非同期呼び出しを入れ子にしています。

- 最初に、[Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) メソッドが呼び出され、"MyBinding" という名前のドキュメントのバインドにアクセスします。 その `AsyncResult` コールバックのパラメーターに返 `result` されるオブジェクトは、プロパティから指定されたバインド オブジェクトへのアクセスを提供 `AsyncResult.value` します。
- 次に、最初のパラメーターからアクセスするバインド オブジェクトを使用して `result` [Binding.getDataAsync メソッドを呼び出](/javascript/api/office/office.binding#getDataAsync_options__callback_) します。
- 最後に、メソッド `result2` に渡されるコールバックのパラメーターを使用して、バインド `Binding.getDataAsync` 内のデータを表示します。

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

この基本的な入れ子になったコールバック パターンは、JavaScript API 内のすべての非同期メソッドOffice使用できます。

次のセクションでは、非同期メソッドの入れ子のコールバックで匿名関数または名前付き関数を使用する方法を示します。

#### <a name="using-anonymous-functions-for-nested-callbacks"></a>入れ子のコールバックとして匿名関数を使用する

次の例では、2 つの匿名関数がインラインで宣言され、and メソッドに入れ子になった `getByIdAsync` `getDataAsync` コールバックとして渡されます。 関数は単純でインラインのため、実装の意図は明白です。

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

#### <a name="using-named-functions-for-nested-callbacks"></a>入れ子のコールバックとして名前付き関数を使用する

複雑な実装の場合、名前付き関数を使用すると、読みやすく、保守管理がしやすく、再利用しやすくなります。 次の例では、前のセクションの例の 2 つの匿名関数が、という名前の関数として書き換 `deleteAllData` えされています `showResult` 。 これらの名前付き関数は、名前によって `getByIdAsync` コールバックとして and `deleteAllDataValuesAsync` メソッドに渡されます。

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

コールバック関数を渡し、その関数が戻るのを待ってから実行を続行する代わりに、promise プログラミング パターンを使用すれば、その意図した結果を表す promise オブジェクトがすぐに返されます。ただし、本物の同期プログラミングとは異なり、実際には Office アドインのランタイム環境が要求を完了できるまでは、約束された結果の履行は実際には延期されます。要求が履行されない状況に対処するために _onError_ ハンドラーが用意されています。

JavaScript API Officeは[、Office.select](/javascript/api/office#Office_select_expression__callback_)メソッドを使用して、既存のバインド オブジェクトを操作する約束パターンをサポートします。 メソッドに返される promise オブジェクトは、Binding オブジェクトから直接アクセスできる 4 つのメソッド `Office.select` [(getDataAsync、setDataAsync、addHandlerAsync、removeHandlerAsync)](/javascript/api/office/office.binding#getDataAsync_options__callback_)のみを[サポート](/javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_)[](/javascript/api/office/office.binding)[](/javascript/api/office/office.binding#setDataAsync_data__options__callback_)[](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_)します。

バインドと連携する promise パターンは次のような形式になります。

**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_

_selectorExpression_ パラメーターはフォームを受け取ります。bindingId は、ドキュメントまたはスプレッドシートで以前に作成したバインドの名前 ( ) です (コレクションの `"bindings#bindingId"`  `id` "addFrom" メソッドの 1 つを使用します `Bindings` 。 `addFromNamedItemAsync` `addFromPromptAsync` `addFromSelectionAsync` たとえば、セレクター式は、'cities' の ID を持つバインドにアクセス `bindings#cities` する場合に指定します。 

_onError パラメーターは_、メソッドが指定されたバインドにアクセスできない場合に、オブジェクトにアクセスするために使用できる型の 1 つのパラメーターを受け取るエラー処理 `AsyncResult` `Error` `select` 関数です。 次の例は、_onError_ パラメーターに渡すことができる基本的なエラー処理関数を示しています。

```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

_BindingObjectAsyncMethod_ プレースホルダーを promise オブジェクトでサポートされている 4 つのオブジェクト メソッドの呼び出しに置き換 `Binding` える: `getDataAsync` `setDataAsync` 、、、 `addHandlerAsync` または `removeHandlerAsync` 。 これらのメソッドの呼び出しでは追加の promise がサポートされません。 これらは[入れ子のコールバック関数パターン](#asynchronous-programming-using-nested-callback-functions)を使用して呼び出す必要があります。

オブジェクトの約束が満たされた後、チェーンメソッド呼び出しで、バインドである場合と同様に再利用できます (アドイン ランタイムは、約束を満たすのを非同期的に再試行する必要はありません `Binding` )。 オブジェクトの約束を満たすことができない場合、次に非同期メソッドの 1 つが呼び出されると、アドイン ランタイムはバインド オブジェクトへのアクセスを再試行 `Binding` します。

次のコード例では、メソッドを使用してコレクションから " とバインドを取得し `select` `id` `cities` `Bindings` [、addHandlerAsync](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_) メソッドを呼び出して、バインドの [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントのイベント ハンドラーを追加します。

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> メソッド `Binding` によって返されるオブジェクトの約束 `Office.select` は、オブジェクトの 4 つのメソッドにのみアクセス `Binding` できます。 オブジェクトの他のメンバーにアクセスする必要がある場合は、代わりにプロパティまたはメソッドを使用してオブジェクト `Binding` `Document.bindings` `Bindings.getByIdAsync` `Bindings.getAllAsync` を取得する必要 `Binding` があります。 たとえば、オブジェクトのプロパティ (、 、またはプロパティ) にアクセスする必要がある場合、 `Binding` `document` または `id` MatrixBinding オブジェクトまたは `type` [TableBinding](/javascript/api/office/office.matrixbinding)[](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` `Binding` オブジェクトのプロパティにアクセスする必要がある場合は、or メソッドを使用してオブジェクトを取得する必要があります。

## <a name="passing-optional-parameters-to-asynchronous-methods"></a>オプションのパラメーターを非同期メソッドに渡す

すべての "Async" メソッドの一般的な構文は、次のパターンに従います。

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

すべての非同期メソッドは、オプションのパラメーターをサポートします。これらは、1 つまたは複数のオプションのパラメーターが格納された JavaScript Object Notation (JSON) オブジェクトとして渡されます。オプションのパラメーターを格納している JSON オブジェクトは、キーと値のペアの順不同のコレクションで、キーと値は ":" 文字で区切られています。オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体はかっこで囲まれます。キーはパラメーター名で、値はそのパラメーターに渡す値です。

オプションのパラメーターをインラインで含む JSON オブジェクトを作成するか、オブジェクトを作成して options パラメーター `options` として渡します。 

### <a name="passing-optional-parameters-inline"></a>オプションのパラメーターをインラインで渡す

たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) メソッドを呼び出す場合の構文は、次のようになります。

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

呼び出し元の構文のこの形式では _、coercionType_ と _asyncContext_ という 2 つの省略可能なパラメーターは、中かっこでインラインで囲まれた JSON オブジェクトとして定義されます。

次の例は、オプションのパラメーターをインラインで `Document.setSelectedDataAsync` 指定してメソッドを呼び出す方法を示しています。

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
> オプションのパラメーターは、名前さえ正しければ、任意の順序で JSON オブジェクトに指定できます。

### <a name="passing-optional-parameters-in-an-options-object"></a>オプションのパラメーターを options オブジェクトで渡す

または、省略可能なパラメーターをメソッド呼び出しとは別に指定するオブジェクトを作成し、options 引数としてオブジェクト `options` `options` を _渡_ します。

次の例は、オブジェクトを作成する方法の 1 つを示しています。ここで、実際のパラメーター名と値のプレースホルダー `options` `parameter1` `value1` です。

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

[ValueFormat](/javascript/api/office/office.valueformat) パラメーターおよび [FilterType](/javascript/api/office/office.filtertype) パラメーターを指定する場合は次のようになります。

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

オブジェクトを作成する別の方法を次に示 `options` します。

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

and パラメーターの指定に使用する場合、次の例のようになります `ValueFormat` `FilterType` 。

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> オブジェクトを作成する方法を使用する場合は、名前が正しく指定されている限り、任意の順序でオプションのパラメーター `options` を指定できます。

次の例は、オブジェクトでオプションのパラメーター `Document.setSelectedDataAsync` を指定してメソッドを呼び出す方法を示 `options` しています。

```js
var options = {
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

どちらのオプションのパラメーター例でも、_callback_ パラメーターが最後のパラメーターとして (インラインのオプションのパラメーターまたは _options_ 引数オブジェクトに続けて) 指定されています。 _callback_ パラメーターは、インライン JSON オブジェクトの中で、または `options` オブジェクト内で指定することもできます。 ただし、 _callback_ パラメーターを渡せるのは _options_ オブジェクト (インラインまたは外部で作成) または最後のパラメーターのどちらか一方であり、両方に渡すことはできません。

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Promise-returning 関数で一般的な API をラップする

Common API (および Outlook API) メソッドは Promises を[返します](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。 したがって、非同期操作が完了するまで [、await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) を使用して実行を一時停止することはできません。 動作が必要 `await` な場合は、明示的に作成された Promise でメソッド呼び出しをラップできます。 

基本的なパターンは、Promise オブジェクトをすぐに返す非同期メソッドを作成し、内部メソッドの完了時に Promise オブジェクトを解決するか、メソッドが失敗した場合にオブジェクトを拒否します。 次に簡単な例を示します。

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

このメソッドを待つ必要がある場合は、キーワードを使用するか、関数に渡 `await` す関数として呼び出 `then` すことができます。

> [!NOTE]
> この手法は、アプリケーション固有のオブジェクト モデルの 1 つでメソッドの呼び出し内で共通 API の 1 つを呼び出す必要がある場合に特に `run` 便利です。 この方法で使用されている上記の関数の例については、 [ サンプル Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)のファイルHome.jsを参照してください。

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
