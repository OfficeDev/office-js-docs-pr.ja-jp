---
title: Office アドインにおける非同期プログラミング
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: fc39bddbe050f8253769a0013be2d48b26dcb599
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324646"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office アドインにおける非同期プログラミング

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office アドイン API が非同期プログラミングを使用するのはなぜですか?JavaScript はシングルスレッド言語なので、スクリプトが長時間実行されている同期プロセスを呼び出すと、そのプロセスが完了するまで、以降のすべてのスクリプトの実行がブロックされます。Office web クライアントに対する特定の操作 (ただし、リッチクライアントでも) が同期的に実行されると、実行がブロックされる可能性があるため、ほとんどの Office JavaScript Api は非同期で実行するように設計されています。これにより、Office アドインが迅速に応答できるようになります。また、これらの非同期メソッドを使用するときには、コールバック関数を作成することもよくあります。

API の最後にあるすべての非同期メソッドの名前。 "Async" (、、 `Document.getSelectedDataAsync`、 `Binding.getDataAsync`または`Item.loadCustomPropertiesAsync`メソッドなど)。"Async" メソッドが呼び出されると、すぐに実行され、それ以降のスクリプト実行を続行できます。"Async" メソッドに渡すオプションのコールバック関数は、データまたは要求された操作の準備が完了するとすぐに実行されます。これは通常はすぐに行われますが、戻るまで少し時間がかかる場合があります。

次の図は、サーバーベースの Word または Excel で開いたドキュメントでユーザーが選択したデータを読み込む "Async" メソッドの呼び出しを実行するフローを示したものです。"Async" が呼び出された時点で、JavaSript 実行スレッドは自由にクライアント側の追加処理を実行できます (ただし、この追加処理は図に示されていません)。"Async" メソッドが戻ると、コールバックはスレッドの実行を再開します。アドインはデータにアクセスし、何らかの操作を行い、結果を表示できます。Word 2013 や Excel 2013 など、Office リッチ クライアント ホスト アプリケーションを使用しているときは、同じ非同期実行パターンが当てはまります。

*図 1. 非同期プログラミング実行フロー*

![非同期プログラミング スレッドの実行フロー](../images/office-addins-asynchronous-programming-flow.png)

リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。

## <a name="writing-the-callback-function-for-an-async-method"></a>"Async" メソッドのコールバック関数を記述する


"Async" メソッドにコール_バック_引数として渡すコールバック関数は、コールバック関数の実行時にアドインランタイムが[AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトへのアクセスを提供するために使用する1つのパラメーターを宣言する必要があります。次のように記述できます。


- "Async" メソッドの_callback_パラメーターとして "async" メソッドを呼び出すことで、行内で直接記述して渡す必要がある匿名関数。

- "Async" メソッドの_callback_パラメーターとしてその関数の名前を渡す、名前付き関数。

匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。


### <a name="writing-an-anonymous-callback-function"></a>匿名コールバック関数を記述する

次の匿名のコールバック関数は、コール`result`バックが戻るときに、 [AsyncResult](/javascript/api/office/office.asyncresult#value)プロパティからデータを取得するという名前の単一のパラメーターを宣言します。


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

次の例は、 `Document.getSelectedDataAsync`メソッドへの完全な "Async" メソッド呼び出しのコンテキストで、この匿名のコールバック関数をインラインで渡す方法を示しています。


- 最初の_coercionType_引数`Office.CoercionType.Text`は、選択されているデータをテキストの文字列として返すように指定します。

- 2番目の_callback_引数は、メソッドにインラインで渡される匿名関数です。関数が実行されると、 _result_パラメーターを使用して`value` `AsyncResult`オブジェクトのプロパティにアクセスし、ドキュメント内のユーザーが選択したデータを表示します。


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

また、コールバック関数のパラメーターを使用して、 `AsyncResult`オブジェクトの他のプロパティにアクセスすることもできます。呼び出しが成功したか失敗したかを確認するには、 [AsyncResult](/javascript/api/office/office.asyncresult#status)プロパティを使用します。呼び出しが失敗した場合は、 [AsyncResult. error](/javascript/api/office/office.asyncresult#error)プロパティを使用してエラー情報の[error](/javascript/api/office/office.error)オブジェクトにアクセスできます。

メソッドの`getSelectedDataAsync`使用法の詳細については、「[文書またはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。 


### <a name="writing-a-named-callback-function"></a>名前付き関数を記述する

別の方法として、名前付き関数を記述して、その名前を "Async" メソッドの_callback_パラメーターに渡すこともできます。たとえば、前述の例では、という名前`writeDataCallback`の関数を、次のような_callback_パラメーターとして渡すように書き換えることができます。


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


オブジェクトの、、 `error`およびプロパティは`asyncContext`、すべての "Async" メソッドに渡されたコールバック関数に対して同じ種類の情報を返します。 `status` `AsyncResult`ただし、 `AsyncResult.value`プロパティに返される内容は、"Async" メソッドの機能によって異なります。

たとえば`addHandlerAsync` 、メソッド ( [Binding](/javascript/api/office/office.binding)、 [CustomXmlPart](/javascript/api/office/office.customxmlpart)、 [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、および[Settings](/javascript/api/office/office.settings)オブジェクトの) を使用して、これらのオブジェクトによって表されるアイテムにイベントハンドラー関数を追加します。メソッドに渡すコール`AsyncResult.value`バック関数からプロパティにアクセスできますが、イベントハンドラーを追加するときにデータまたはオブジェクトにアクセスされていないため、 `value`このプロパティにアクセスしようとすると、常に undefined が返されます。 **** `addHandlerAsync`

一方、 `Document.getSelectedDataAsync`メソッドを呼び出すと、コールバックの`AsyncResult.value`プロパティにドキュメントでユーザーが選択したデータが返されます。または、 [getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-)メソッドを呼び出すと、ドキュメント内のすべての`Binding`オブジェクトの配列が返されます。また、バインディングを呼び出した場合は、1つ`Binding`のオブジェクトを返し[ます。](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)

`Async`メソッドの`AsyncResult.value`プロパティに返される内容の詳細については、そのメソッドのリファレンストピックの「コールバック値」セクションを参照してください。メソッドを提供`Async`するすべてのオブジェクトの概要については、「 [AsyncResult](/javascript/api/office/office.asyncresult) object」の下にある表を参照してください。


## <a name="asynchronous-programming-patterns"></a>非同期プログラミング パターン


Office JavaScript API は、次の2種類の非同期プログラミングパターンをサポートしています。


- 入れ子のコールバックの使用
    
- promise パターンの使用
    
コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。

ネストされたコールバックは、ほとんどの JavaScript 開発者にとってよく使用されるプログラミングパターンですが、深くネストされたコールバックを使用したコードは、読みやすく理解しにくくなる可能性があります。Office JavaScript API は、ネストされたコールバックの代わりに、約束パターンの実装もサポートしています。ただし、現在のバージョンの Office JavaScript API では、約束パターンは、 [Excel スプレッドシートと Word 文書内のバインディング](bind-to-regions-in-a-document-or-spreadsheet.md)のコードでのみ機能します。

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>入れ子のコールバック関数を使用する非同期プログラミング


多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。

次のコード例では、2 つの非同期呼び出しを入れ子にしています。


- 最初に、"MyBinding" という名前のドキュメント内のバインドにアクセスするために、 [getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)メソッドが呼び出されます。その`AsyncResult`コールバックの`result`パラメーターに返されるオブジェクトは、指定された binding オブジェクトへ`AsyncResult.value`のアクセスをプロパティから提供します。

- 次に、最初`result`のパラメーターからアクセスされる binding オブジェクトを使用して、 [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)メソッドを呼び出します。

- 最後に、 `result2` `Binding.getDataAsync`メソッドに渡されたコールバックのパラメーターを使用して、バインド内のデータを表示します。


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

次の例では、2つの匿名関数をインラインで宣言`getByIdAsync`し`getDataAsync` 、入れ子にされたコールバックとしてメソッドに渡します。関数は単純でインラインなので、実装の目的はすぐにわかります。


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

複雑な実装では、名前付き関数を使用して、コードをより簡単に読み取り、維持、および再利用できるようにするのに役立ちます。次の例では、前のセクションの例にある2つの匿名関数がと`deleteAllData` `showResult`いう名前の関数として書き換えられています。これらの名前付き関数は、 `getByIdAsync`名前に`deleteAllDataValuesAsync`よってコールバックとしてメソッドに渡されます。


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


Office JavaScript API は、既存の binding オブジェクトを使用するための約束パターンをサポートするために、 [office の select](/javascript/api/office#office-select-expression--callback-)メソッドを提供します。`Office.select`メソッドに返される promise オブジェクトは、 [Binding](/javascript/api/office/office.binding)オブジェクトから直接アクセスできる、 [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)、 [Setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)、 [addハンドラ async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)、および[removeハンドラ async](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)の4つのメソッドのみをサポートしています。


バインドと連携する promise パターンは次のような形式になります。

 **Office。 [(**_selectorExpression_、 _onError_**)** ] を選択します。_BindingObjectAsyncMethod_

_SelectorExpression_パラメーターは`"bindings#bindingId"`、フォームを取得します。ここで、 _bindingid_は、ドキュメントまたはスプレッドシートで以前に作成したバインドの`id`名前 () です (この`Bindings`コレクション`addFromNamedItemAsync` `addFromPromptAsync`の "addfrom" メソッド`addFromSelectionAsync`のいずれかを使用します)。たとえば、selector 式`bindings#cities`は、 **id**が ' 都市 ' であるバインドにアクセスすることを指定します。

_OnError_パラメーターは、指定されたバインドへのアクセスに失敗し`AsyncResult`た場合`Error` `select`に、オブジェクトへのアクセスに使用できる1つのパラメーター型のパラメーターを受け取るエラー処理関数です。次の例は、 _onError_パラメーターに渡すことができる基本的なエラーハンドラ関数を示しています。




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

_BindingObjectAsyncMethod_ placeholder を`Binding` `getDataAsync`、promise オブジェクトでサポートされている4つのオブジェクトのいずれかのメソッド`setDataAsync`( `addHandlerAsync`、、 `removeHandlerAsync`、または) を呼び出して置き換えます。これらのメソッドを呼び出すことは、追加の約束をサポートしていません。ネストされた[コールバック関数パターン](#AsyncProgramming_NestedCallbacks)を使用して呼び出す必要があります。

`Binding`オブジェクトの promise が履行された後は、バインドされている場合と同様に、連鎖メソッドの呼び出しで再利用できます (アドインランタイムは、promise を実行しても非同期に再試行することはありません)。`Binding`オブジェクトの約束を履行できない場合、アドインランタイムは、次にその非同期メソッドが呼び出されたときに binding オブジェクトへのアクセスを再試行します。

次のコード例では`select` 、メソッドを使用して、 `id` "`cities`" という`Bindings`バインドをコレクションから取得した後、 [addhandler async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)メソッドを呼び出して、バインドの[dataChanged](/javascript/api/office/office.bindingdatachangedeventargs)イベントのイベントハンドラーを追加します。




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> メソッドによって返されるオブジェクトの promise は`Binding` 、 `Binding`オブジェクトの4つのメソッドへのアクセスのみを提供します。 `Office.select``Binding`オブジェクトの他のメンバーにアクセスする必要がある場合`Document.bindings`は、その`Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding`オブジェクトを取得するのにプロパティとメソッドのどちらか一方または両方を使用する必要があります。`Binding`たとえば、オブジェクトのプロパティ ( `document` `id`、、または`type`プロパティ) にアクセスする必要がある場合や、 [MatrixBinding](/javascript/api/office/office.matrixbinding)または[tablebinding](/javascript/api/office/office.tablebinding)オブジェクトのプロパティにアクセスする必要がある場合は、 `getByIdAsync`または`getAllAsync`メソッドを使用して`Binding`オブジェクトを取得する必要があります。


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>オプションのパラメーターを非同期メソッドに渡す


すべての "Async" メソッドの一般的な構文は、次のパターンに従います。

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

すべての非同期メソッドは、オプションのパラメーターをサポートします。これらは、1 つまたは複数のオプションのパラメーターが格納された JavaScript Object Notation (JSON) オブジェクトとして渡されます。オプションのパラメーターを格納している JSON オブジェクトは、キーと値のペアの順不同のコレクションで、キーと値は ":" 文字で区切られています。オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体はかっこで囲まれます。キーはパラメーター名で、値はそのパラメーターに渡す値です。

省略可能なパラメーターを含む JSON オブジェクトはインラインで作成するか、 `options`オブジェクトを作成し、それを_options_パラメーターとして渡すことによって作成できます。


### <a name="passing-optional-parameters-inline"></a>オプションのパラメーターをインラインで渡す

たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドを呼び出す場合の構文は、次のようになります。

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

この呼び出し構文では、 _coercionType_および_asynccontext_という2つのオプションのパラメーターが、かっこで囲まれてインラインで JSON オブジェクトとして定義されています。

次の例は、省略可能なパラメーター `Document.setSelectedDataAsync`をインラインで指定することによってメソッドを呼び出す方法を示しています。


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

または、メソッド呼び出しとは別`options`に、オプションのパラメーターを指定するという名前のオブジェクトを作成`options`し、そのオブジェクトを_options_引数として渡すことができます。

次の例は、 `options`オブジェクトを作成する1つの`parameter1`方法`value1`を示しています。ここで、、、などは、実際のパラメーター名と値のプレースホルダーです。




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

オブジェクトを`options`作成する別の方法を次に示します。




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

パラメーターと`ValueFormat` `FilterType`パラメーターの指定に使用する場合は、次の例のようになります。


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> いずれかの方法で`options`オブジェクトを作成する場合は、名前が正しく指定されていれば、任意の順序でオプションのパラメーターを指定できます。

次の例は、 `Document.setSelectedDataAsync` `options`オブジェクトで省略可能なパラメーターを指定してメソッドを呼び出す方法を示しています。




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


両方の省略可能なパラメーターの例では、 _callback_パラメーターが最後のパラメーターとして指定されています (インラインオプションのパラメーターの後、または_options_引数オブジェクトの後)。または、_コールバック_パラメーターをインライン JSON オブジェクトまたは`options`オブジェクトの内部で指定することもできます。ただし、 _callback_パラメーターを渡すことができるのは、 _options_オブジェクト (インラインまたは外部で作成されている)、または最後のパラメーター (両方ではない) のいずれかです。


## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office)
