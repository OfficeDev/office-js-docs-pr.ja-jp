---
title: Office アドインにおける非同期プログラミング
description: Office JavaScript ライブラリが Office アドインで非同期プログラミングを使用する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: affe493cdf1633b3a8749b694da479a732271195
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292945"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office アドインにおける非同期プログラミング

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office アドイン API が非同期プログラミングを使用するのはなぜですか? JavaScript はシングルスレッドの言語であるため、スクリプトで実行時間の長い同期プロセスが呼び出されると、そのプロセスが完了するまで後続のすべてのスクリプト実行がブロックされます。 Office web クライアントに対する特定の操作 (ただし、リッチクライアントでも) が同期的に実行されると、実行がブロックされる可能性があるため、ほとんどの Office JavaScript Api は非同期で実行するように設計されています。 これにより、Office アドインが迅速に応答できるようになります。 このような非同期メソッドを利用するときは、多くの場合、コールバック関数の記述も必要です。

API の最後にあるすべての非同期メソッドの名前。 "Async" (、、 `Document.getSelectedDataAsync` 、 `Binding.getDataAsync` または `Item.loadCustomPropertiesAsync` メソッドなど)。 "Async" メソッドは呼び出されるとすぐに実行され、後続のスクリプトも続けて実行することができます。 "Async" メソッドに渡す任意のコールバック関数は、データまたは要求された操作の準備が整い次第、すぐに実行されます。 コールバック関数の実行は通常、直ちに行われますが、戻るまでに若干の遅延が生じることがあります。

次の図は、サーバーベースの Word または Excel で開いているドキュメントでユーザーが選択したデータを読み取る、"Async" メソッドへの呼び出しの実行フローを示しています。 "Async" 呼び出しが行われる時点で、JavaScript 実行スレッドは、追加のクライアント側の処理を自由に実行できます (ただし、図には何も表示されません)。 "Async" メソッドが戻ると、コールバックはスレッドの実行を再開し、アドインはアクセスデータを操作して、その結果を表示することができます。 Word 2013 または Excel 2013 などの Office リッチクライアントアプリケーションを使用する場合、同じ非同期実行パターンが保持されます。

*図 1. 非同期プログラミング実行フロー*

![非同期プログラミング スレッドの実行フロー](../images/office-addins-asynchronous-programming-flow.png)

リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。

## <a name="writing-the-callback-function-for-an-async-method"></a>"Async" メソッドのコールバック関数を記述する


"Async" メソッドにコール _バック_ 引数として渡すコールバック関数は、コールバック関数の実行時にアドインランタイムが [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトへのアクセスを提供するために使用する1つのパラメーターを宣言する必要があります。 次のように記述することができます。


- "Async" メソッドの _callback_ パラメーターとして "async" メソッドを呼び出すことで、行内で直接記述して渡す必要がある匿名関数。

- "Async" メソッドの _callback_ パラメーターとしてその関数の名前を渡す、名前付き関数。

匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。


### <a name="writing-an-anonymous-callback-function"></a>匿名コールバック関数を記述する

次の匿名のコールバック関数は、 `result` コールバックが戻るときに、 [AsyncResult](/javascript/api/office/office.asyncresult#value) プロパティからデータを取得するという名前の単一のパラメーターを宣言します。


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

次の例は、メソッドへの完全な "Async" メソッド呼び出しのコンテキストで、この匿名のコールバック関数をインラインで渡す方法を示して `Document.getSelectedDataAsync` います。


- 最初の _coercionType_ 引数は、 `Office.CoercionType.Text` 選択されているデータをテキストの文字列として返すように指定します。

- 2番目の _callback_ 引数は、メソッドにインラインで渡される匿名関数です。 関数が実行されると、 _result_ パラメーターを使用して `value` オブジェクトのプロパティにアクセスし、 `AsyncResult` ドキュメント内のユーザーが選択したデータを表示します。


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

また、コールバック関数のパラメーターを使用して、オブジェクトの他のプロパティにアクセスすることもでき `AsyncResult` ます。 呼び出しの成功または失敗を判断する場合は [AsyncResult.status](/javascript/api/office/office.asyncresult#status) プロパティを使用します。 呼び出しが失敗した場合は [AsyncResult.error](/javascript/api/office/office.asyncresult#error) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、エラーの詳細を確認できます。

メソッドの使用法の詳細については `getSelectedDataAsync` 、「 [文書またはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。 


### <a name="writing-a-named-callback-function"></a>名前付き関数を記述する

別の方法として、名前付き関数を記述して、その名前を "Async" メソッドの _callback_ パラメーターに渡すこともできます。 たとえば、前の例は次のように `writeDataCallback` という名前の関数を _callback_ パラメーターとして渡すように書き換えることができます。


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


`asyncContext` `status` オブジェクトの、、およびプロパティは、 `error` `AsyncResult` すべての "Async" メソッドに渡されたコールバック関数に対して同じ種類の情報を返します。 ただし、プロパティに返される内容は、 `AsyncResult.value` "Async" メソッドの機能によって異なります。

たとえば、 `addHandlerAsync` メソッド ( [Binding](/javascript/api/office/office.binding)、 [CustomXmlPart](/javascript/api/office/office.customxmlpart)、 [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、および [Settings](/javascript/api/office/office.settings) オブジェクトの) を使用して、これらのオブジェクトによって表されるアイテムにイベントハンドラー関数を追加します。 `AsyncResult.value`メソッドに渡すコールバック関数からプロパティにアクセスできます `addHandlerAsync` が、イベントハンドラーを追加するときにデータまたはオブジェクトにアクセスされていないため、 `value` このプロパティにアクセスしようとすると、常に**undefined**が返されます。

一方、メソッドを呼び出すと `Document.getSelectedDataAsync` 、コールバックのプロパティにドキュメントでユーザーが選択したデータが返され `AsyncResult.value` ます。 または、getAllAsync メソッドを呼び出すと、ドキュメント内のすべてのオブジェクトの配列が返されます[。](/javascript/api/office/office.bindings#getallasync-options--callback-) `Binding` また、バインディングを呼び出した場合は、1つのオブジェクトを返し[ます。](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) `Binding`

メソッドのプロパティに返される内容の詳細につい `AsyncResult.value` `Async` ては、そのメソッドのリファレンストピックの「コールバック値」セクションを参照してください。 メソッドを提供するすべてのオブジェクトの概要については、「 `Async` [AsyncResult](/javascript/api/office/office.asyncresult) object」の下にある表を参照してください。


## <a name="asynchronous-programming-patterns"></a>非同期プログラミング パターン


Office JavaScript API は、次の2種類の非同期プログラミングパターンをサポートしています。


- 入れ子のコールバックの使用
    
- promise パターンの使用
    
コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。

入れ子のコールバックを使用することは、ほとんどの JavaScript 開発者にとってなじみのあるプログラミング パターンですが、コールバックが深い入れ子になっているコードは読みにくく、理解しにくいものです。 Office JavaScript API は、ネストされたコールバックの代わりに、約束パターンの実装もサポートしています。 ただし、現在のバージョンの Office JavaScript API では、約束パターンは、 [Excel スプレッドシートと Word 文書内のバインディング](bind-to-regions-in-a-document-or-spreadsheet.md)のコードでのみ機能します。

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>入れ子のコールバック関数を使用する非同期プログラミング


多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。

次のコード例では、2 つの非同期呼び出しを入れ子にしています。


- 最初に、[Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) メソッドが呼び出され、"MyBinding" という名前のドキュメントのバインドにアクセスします。 `AsyncResult`そのコールバックのパラメーターに返されるオブジェクトは、 `result` 指定された binding オブジェクトへのアクセスをプロパティから提供し `AsyncResult.value` ます。

- 次に、最初のパラメーターからアクセスされる binding オブジェクトを使用して、 `result` [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) メソッドを呼び出します。

- 最後に、 `result2` メソッドに渡されたコールバックのパラメーターを使用して、 `Binding.getDataAsync` バインド内のデータを表示します。


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

この基本的な入れ子のコールバックパターンは、Office JavaScript API のすべての非同期メソッドに使用できます。

次のセクションでは、非同期メソッドの入れ子のコールバックで匿名関数または名前付き関数を使用する方法を示します。


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>入れ子のコールバックとして匿名関数を使用する

次の例では、2つの匿名関数をインラインで宣言し、入れ子にされた `getByIdAsync` `getDataAsync` コールバックとしてメソッドに渡します。 関数は単純でインラインのため、実装の意図は明白です。


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

複雑な実装の場合、名前付き関数を使用すると、読みやすく、保守管理がしやすく、再利用しやすくなります。 次の例では、前のセクションの例にある2つの匿名関数がという名前の関数として書き換えられてい `deleteAllData` `showResult` ます。 これらの名前付き関数は、 `getByIdAsync` `deleteAllDataValuesAsync` 名前によってコールバックとしてメソッドに渡されます。


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


Office JavaScript API は、既存の binding オブジェクトを使用するための約束パターンをサポートするために、 [office の select](/javascript/api/office#office-select-expression--callback-) メソッドを提供します。 メソッドに返される promise オブジェクトは、 `Office.select` [Binding](/javascript/api/office/office.binding) オブジェクトから直接アクセスできる、 [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)、 [setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)、 [addハンドラ async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)、および [removeハンドラ async](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)の4つのメソッドのみをサポートしています。


バインドと連携する promise パターンは次のような形式になります。

 **Office。 [(**_selectorExpression_、 _onError_**)** ] を選択します。_BindingObjectAsyncMethod_

_SelectorExpression_パラメーターは、フォームを取得します。 `"bindings#bindingId"` ここで、 _bindingid_は、 `id` ドキュメントまたはスプレッドシートで以前に作成したバインドの名前 () です (このコレクションの "addfrom" メソッドのいずれかを使用します `Bindings` `addFromNamedItemAsync` `addFromPromptAsync` `addFromSelectionAsync` )。 たとえば、selector 式は、 `bindings#cities` **id** が ' 都市 ' であるバインドにアクセスすることを指定します。

_OnError_パラメーターは、 `AsyncResult` 指定された `Error` `select` バインドへのアクセスに失敗した場合に、オブジェクトへのアクセスに使用できる1つのパラメーター型のパラメーターを受け取るエラー処理関数です。 次の例は、_onError_パラメーターに渡すことができる基本的なエラー処理関数を示しています。




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

_BindingObjectAsyncMethod_ placeholder を `Binding` 、promise オブジェクトでサポートされている4つのオブジェクトのいずれかのメソッド (、、、 `getDataAsync` `setDataAsync` `addHandlerAsync` または `removeHandlerAsync` ) を呼び出して置き換えます。 これらのメソッドの呼び出しでは追加の promise がサポートされません。 これらは[入れ子のコールバック関数パターン](#AsyncProgramming_NestedCallbacks)を使用して呼び出す必要があります。

オブジェクトの `Binding` promise が履行された後は、バインドされている場合と同様に、連鎖メソッドの呼び出しで再利用できます (アドインランタイムは、promise を実行しても非同期に再試行することはありません)。 オブジェクトの `Binding` 約束を履行できない場合、アドインランタイムは、次にその非同期メソッドが呼び出されたときに binding オブジェクトへのアクセスを再試行します。

次のコード例では、メソッドを使用して、 `select` "" というバインドを `id` `cities` コレクションから取得 `Bindings` した後、 [addhandler async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出して、バインドの [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントのイベントハンドラーを追加します。




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> `Binding`メソッドによって返されるオブジェクトの promise は、 `Office.select` オブジェクトの4つのメソッドへのアクセスのみを提供し `Binding` ます。 オブジェクトの他のメンバーにアクセスする必要がある場合は `Binding` 、その `Document.bindings` オブジェクトを取得するのにプロパティとメソッドのどちらか一方または両方を使用する必要があり `Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding` ます。 たとえば、オブジェクトのプロパティ (、、またはプロパティ) にアクセスする必要がある場合や、MatrixBinding または Tablebinding オブジェクトのプロパティにアクセスする必要がある場合は、 `Binding` `document` `id` `type` またはメソッドを使用して[MatrixBinding](/javascript/api/office/office.matrixbinding) [TableBinding](/javascript/api/office/office.tablebinding) `getByIdAsync` オブジェクトを取得する必要があり `getAllAsync` `Binding` ます。


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>オプションのパラメーターを非同期メソッドに渡す


すべての "Async" メソッドの一般的な構文は、次のパターンに従います。

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

すべての非同期メソッドは、オプションのパラメーターをサポートします。これらは、1 つまたは複数のオプションのパラメーターが格納された JavaScript Object Notation (JSON) オブジェクトとして渡されます。オプションのパラメーターを格納している JSON オブジェクトは、キーと値のペアの順不同のコレクションで、キーと値は ":" 文字で区切られています。オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体はかっこで囲まれます。キーはパラメーター名で、値はそのパラメーターに渡す値です。

省略可能なパラメーターを含む JSON オブジェクトはインラインで作成するか、オブジェクトを作成 `options` し、それを _options_ パラメーターとして渡すことによって作成できます。


### <a name="passing-optional-parameters-inline"></a>オプションのパラメーターをインラインで渡す

たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドを呼び出す場合の構文は、次のようになります。

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

この呼び出し構文では、 _coercionType_ および _asynccontext_という2つのオプションのパラメーターが、かっこで囲まれてインラインで JSON オブジェクトとして定義されています。

次の例は、 `Document.setSelectedDataAsync` 省略可能なパラメーターをインラインで指定することによってメソッドを呼び出す方法を示しています。


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

または、メソッド呼び出しとは別に、オプションのパラメーターを指定するという名前のオブジェクトを作成し、その `options` `options` オブジェクトを _options_ 引数として渡すことができます。

次の例は、オブジェクトを作成する1つの方法を示してい `options` ます。ここで、 `parameter1` `value1` 、、などは、実際のパラメーター名と値のプレースホルダーです。




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

オブジェクトを作成する別の方法を次に示し `options` ます。




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

パラメーターとパラメーターの指定に使用する場合は、次の例のように `ValueFormat` `FilterType` なります。


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> いずれかの方法でオブジェクトを作成する場合は `options` 、名前が正しく指定されていれば、任意の順序でオプションのパラメーターを指定できます。

次の例は、 `Document.setSelectedDataAsync` オブジェクトで省略可能なパラメーターを指定してメソッドを呼び出す方法を示して `options` います。




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


## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
