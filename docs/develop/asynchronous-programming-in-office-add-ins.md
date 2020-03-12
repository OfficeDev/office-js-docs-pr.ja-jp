---
title: Office アドインにおける非同期プログラミング
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 931ef17115885c8f96d41bf00143b3269a515d56
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596691"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="50ca1-102">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="50ca1-102">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="50ca1-p101">Office アドイン API が非同期プログラミングを使用するのはなぜですか?JavaScript はシングルスレッド言語なので、スクリプトが長時間実行されている同期プロセスを呼び出すと、そのプロセスが完了するまで、以降のすべてのスクリプトの実行がブロックされます。Office web クライアントに対する特定の操作 (ただし、リッチクライアントでも) が同期的に実行されると、実行がブロックされる可能性があるため、ほとんどの Office JavaScript Api は非同期で実行するように設計されています。これにより、Office アドインが迅速に応答できるようになります。また、これらの非同期メソッドを使用するときには、コールバック関数を作成することもよくあります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p101">Why does the Office Add-ins API use asynchronous programming? Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and fast. It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="50ca1-p102">API の最後にあるすべての非同期メソッドの名前。 "Async" (、、 `Document.getSelectedDataAsync`、 `Binding.getDataAsync`または`Item.loadCustomPropertiesAsync`メソッドなど)。"Async" メソッドが呼び出されると、すぐに実行され、それ以降のスクリプト実行を続行できます。"Async" メソッドに渡すオプションのコールバック関数は、データまたは要求された操作の準備が完了するとすぐに実行されます。これは通常はすぐに行われますが、戻るまで少し時間がかかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p102">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="50ca1-p103">次の図は、サーバーベースの Word または Excel で開いたドキュメントでユーザーが選択したデータを読み込む "Async" メソッドの呼び出しを実行するフローを示したものです。"Async" が呼び出された時点で、JavaSript 実行スレッドは自由にクライアント側の追加処理を実行できます (ただし、この追加処理は図に示されていません)。"Async" メソッドが戻ると、コールバックはスレッドの実行を再開します。アドインはデータにアクセスし、何らかの操作を行い、結果を表示できます。Word 2013 や Excel 2013 など、Office リッチ クライアント ホスト アプリケーションを使用しているときは、同じ非同期実行パターンが当てはまります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p103">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram). When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="50ca1-116">*図 1. 非同期プログラミング実行フロー*</span><span class="sxs-lookup"><span data-stu-id="50ca1-116">*Figure 1. Asynchronous programming execution flow*</span></span>

![非同期プログラミング スレッドの実行フロー](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="50ca1-p104">リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="50ca1-120">"Async" メソッドのコールバック関数を記述する</span><span class="sxs-lookup"><span data-stu-id="50ca1-120">Writing the callback function for an "Async" method</span></span>


<span data-ttu-id="50ca1-p105">"Async" メソッドにコール_バック_引数として渡すコールバック関数は、コールバック関数の実行時にアドインランタイムが[AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトへのアクセスを提供するために使用する1つのパラメーターを宣言する必要があります。次のように記述できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p105">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes. You can write:</span></span>


- <span data-ttu-id="50ca1-123">"Async" メソッドの_callback_パラメーターとして "async" メソッドを呼び出すことで、行内で直接記述して渡す必要がある匿名関数。</span><span class="sxs-lookup"><span data-stu-id="50ca1-123">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="50ca1-124">"Async" メソッドの_callback_パラメーターとしてその関数の名前を渡す、名前付き関数。</span><span class="sxs-lookup"><span data-stu-id="50ca1-124">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="50ca1-p106">匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>


### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="50ca1-127">匿名コールバック関数を記述する</span><span class="sxs-lookup"><span data-stu-id="50ca1-127">Writing an anonymous callback function</span></span>

<span data-ttu-id="50ca1-128">次の匿名のコールバック関数は、コール`result`バックが戻るときに、 [AsyncResult](/javascript/api/office/office.asyncresult#value)プロパティからデータを取得するという名前の単一のパラメーターを宣言します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-128">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="50ca1-129">次の例は、 `Document.getSelectedDataAsync`メソッドへの完全な "Async" メソッド呼び出しのコンテキストで、この匿名のコールバック関数をインラインで渡す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-129">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>


- <span data-ttu-id="50ca1-130">最初の_coercionType_引数`Office.CoercionType.Text`は、選択されているデータをテキストの文字列として返すように指定します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-130">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="50ca1-p107">2番目の_callback_引数は、メソッドにインラインで渡される匿名関数です。関数が実行されると、 _result_パラメーターを使用して`value` `AsyncResult`オブジェクトのプロパティにアクセスし、ドキュメント内のユーザーが選択したデータを表示します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p107">The second _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>


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

<span data-ttu-id="50ca1-p108">また、コールバック関数のパラメーターを使用して、 `AsyncResult`オブジェクトの他のプロパティにアクセスすることもできます。呼び出しが成功したか失敗したかを確認するには、 [AsyncResult](/javascript/api/office/office.asyncresult#status)プロパティを使用します。呼び出しが失敗した場合は、 [AsyncResult. error](/javascript/api/office/office.asyncresult#error)プロパティを使用してエラー情報の[error](/javascript/api/office/office.error)オブジェクトにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p108">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object. Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="50ca1-136">メソッドの`getSelectedDataAsync`使用法の詳細については、「[文書またはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="50ca1-136">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 


### <a name="writing-a-named-callback-function"></a><span data-ttu-id="50ca1-137">名前付き関数を記述する</span><span class="sxs-lookup"><span data-stu-id="50ca1-137">Writing a named callback function</span></span>

<span data-ttu-id="50ca1-p109">別の方法として、名前付き関数を記述して、その名前を "Async" メソッドの_callback_パラメーターに渡すこともできます。たとえば、前述の例では、という名前`writeDataCallback`の関数を、次のような_callback_パラメーターとして渡すように書き換えることができます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p109">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="50ca1-140">AsyncResult.value プロパティに返される内容の違い</span><span class="sxs-lookup"><span data-stu-id="50ca1-140">Differences in what's returned to the AsyncResult.value property</span></span>


<span data-ttu-id="50ca1-p110">オブジェクトの、、 `error`およびプロパティは`asyncContext`、すべての "Async" メソッドに渡されたコールバック関数に対して同じ種類の情報を返します。 `status` `AsyncResult`ただし、 `AsyncResult.value`プロパティに返される内容は、"Async" メソッドの機能によって異なります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p110">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="50ca1-p111">たとえば`addHandlerAsync` 、メソッド ( [Binding](/javascript/api/office/office.binding)、 [CustomXmlPart](/javascript/api/office/office.customxmlpart)、 [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)、および[Settings](/javascript/api/office/office.settings)オブジェクトの) を使用して、これらのオブジェクトによって表されるアイテムにイベントハンドラー関数を追加します。メソッドに渡すコール`AsyncResult.value`バック関数からプロパティにアクセスできますが、イベントハンドラーを追加するときにデータまたはオブジェクトにアクセスされていないため、 `value`このプロパティにアクセスしようとすると、常に undefined が返されます。 **undefined** `addHandlerAsync`</span><span class="sxs-lookup"><span data-stu-id="50ca1-p111">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects. You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="50ca1-p112">一方、 `Document.getSelectedDataAsync`メソッドを呼び出すと、コールバックの`AsyncResult.value`プロパティにドキュメントでユーザーが選択したデータが返されます。または、 [getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-)メソッドを呼び出すと、ドキュメント内のすべての`Binding`オブジェクトの配列が返されます。また、バインディングを呼び出した場合は、1つ`Binding`のオブジェクトを返し[ます。](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="50ca1-p112">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback. Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document. And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="50ca1-p113">`Async`メソッドの`AsyncResult.value`プロパティに返される内容の詳細については、そのメソッドのリファレンストピックの「コールバック値」セクションを参照してください。メソッドを提供`Async`するすべてのオブジェクトの概要については、「 [AsyncResult](/javascript/api/office/office.asyncresult) object」の下にある表を参照してください。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p113">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>


## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="50ca1-150">非同期プログラミング パターン</span><span class="sxs-lookup"><span data-stu-id="50ca1-150">Asynchronous programming patterns</span></span>


<span data-ttu-id="50ca1-151">Office JavaScript API は、次の2種類の非同期プログラミングパターンをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-151">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>


- <span data-ttu-id="50ca1-152">入れ子のコールバックの使用</span><span class="sxs-lookup"><span data-stu-id="50ca1-152">Using nested callbacks</span></span>
    
- <span data-ttu-id="50ca1-153">promise パターンの使用</span><span class="sxs-lookup"><span data-stu-id="50ca1-153">Using the promises pattern</span></span>
    
<span data-ttu-id="50ca1-p114">コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="50ca1-p115">ネストされたコールバックは、ほとんどの JavaScript 開発者にとってよく使用されるプログラミングパターンですが、深くネストされたコールバックを使用したコードは、読みやすく理解しにくくなる可能性があります。Office JavaScript API は、ネストされたコールバックの代わりに、約束パターンの実装もサポートしています。ただし、現在のバージョンの Office JavaScript API では、約束パターンは、 [Excel スプレッドシートと Word 文書内のバインディング](bind-to-regions-in-a-document-or-spreadsheet.md)のコードでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p115">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern. However, in the current version of the Office JavaScript API, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="50ca1-159">入れ子のコールバック関数を使用する非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="50ca1-159">Asynchronous programming using nested callback functions</span></span>


<span data-ttu-id="50ca1-p116">多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p116">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="50ca1-162">次のコード例では、2 つの非同期呼び出しを入れ子にしています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-162">The following code example nests two asynchronous calls.</span></span>


- <span data-ttu-id="50ca1-p117">最初に、"MyBinding" という名前のドキュメント内のバインドにアクセスするために、 [getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)メソッドが呼び出されます。その`AsyncResult`コールバックの`result`パラメーターに返されるオブジェクトは、指定された binding オブジェクトへ`AsyncResult.value`のアクセスをプロパティから提供します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p117">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding". The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>

- <span data-ttu-id="50ca1-165">次に、最初`result`のパラメーターからアクセスされる binding オブジェクトを使用して、 [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-165">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>

- <span data-ttu-id="50ca1-166">最後に、 `result2` `Binding.getDataAsync`メソッドに渡されたコールバックのパラメーターを使用して、バインド内のデータを表示します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-166">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>


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

<span data-ttu-id="50ca1-167">この基本的な入れ子のコールバックパターンは、Office JavaScript API のすべての非同期メソッドに使用できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-167">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="50ca1-168">次のセクションでは、非同期メソッドの入れ子のコールバックで匿名関数または名前付き関数を使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-168">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>


#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="50ca1-169">入れ子のコールバックとして匿名関数を使用する</span><span class="sxs-lookup"><span data-stu-id="50ca1-169">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="50ca1-p118">次の例では、2つの匿名関数をインラインで宣言`getByIdAsync`し`getDataAsync` 、入れ子にされたコールバックとしてメソッドに渡します。関数は単純でインラインなので、実装の目的はすぐにわかります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p118">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>


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


#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="50ca1-172">入れ子のコールバックとして名前付き関数を使用する</span><span class="sxs-lookup"><span data-stu-id="50ca1-172">Using named functions for nested callbacks</span></span>

<span data-ttu-id="50ca1-p119">複雑な実装では、名前付き関数を使用して、コードをより簡単に読み取り、維持、および再利用できるようにするのに役立ちます。次の例では、前のセクションの例にある2つの匿名関数がと`deleteAllData` `showResult`いう名前の関数として書き換えられています。これらの名前付き関数は、 `getByIdAsync`名前に`deleteAllDataValuesAsync`よってコールバックとしてメソッドに渡されます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p119">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`. These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="50ca1-176">promise パターンを使用してバインドのデータにアクセスする非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="50ca1-176">Asynchronous programming using the promises pattern to access data in bindings</span></span>


<span data-ttu-id="50ca1-p120">コールバック関数を渡し、その関数が戻るのを待ってから実行を続行する代わりに、promise プログラミング パターンを使用すれば、その意図した結果を表す promise オブジェクトがすぐに返されます。ただし、本物の同期プログラミングとは異なり、実際には Office アドインのランタイム環境が要求を完了できるまでは、約束された結果の履行は実際には延期されます。要求が履行されない状況に対処するために _onError_ ハンドラーが用意されています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p120">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>


<span data-ttu-id="50ca1-p121">Office JavaScript API は、既存の binding オブジェクトを使用するための約束パターンをサポートするために、 [office の select](/javascript/api/office#office-select-expression--callback-)メソッドを提供します。`Office.select`メソッドに返される promise オブジェクトは、 [Binding](/javascript/api/office/office.binding)オブジェクトから直接アクセスできる、 [getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)、 [Setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)、 [addハンドラ async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)、および[removeハンドラ async](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)の4つのメソッドのみをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p121">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects. The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>


<span data-ttu-id="50ca1-182">バインドと連携する promise パターンは次のような形式になります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-182">The promises pattern for working with bindings takes this form:</span></span>

 <span data-ttu-id="50ca1-183">**Office。 [(**_selectorExpression_、 _onError_**)** ] を選択します。_BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="50ca1-183">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="50ca1-p122">_SelectorExpression_パラメーターは`"bindings#bindingId"`、フォームを取得します。ここで、 _bindingid_は、ドキュメントまたはスプレッドシートで以前に作成したバインドの`id`名前 () です (この`Bindings`コレクション`addFromNamedItemAsync` `addFromPromptAsync`の "addfrom" メソッド`addFromSelectionAsync`のいずれかを使用します)。たとえば、selector 式`bindings#cities`は、 **id**が ' 都市 ' であるバインドにアクセスすることを指定します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p122">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`). For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="50ca1-p123">_OnError_パラメーターは、指定されたバインドへのアクセスに失敗し`AsyncResult`た場合`Error` `select`に、オブジェクトへのアクセスに使用できる1つのパラメーター型のパラメーターを受け取るエラー処理関数です。次の例は、 _onError_パラメーターに渡すことができる基本的なエラーハンドラ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p123">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>




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

<span data-ttu-id="50ca1-p124">_BindingObjectAsyncMethod_ placeholder を`Binding` `getDataAsync`、promise オブジェクトでサポートされている4つのオブジェクトのいずれかのメソッド`setDataAsync`( `addHandlerAsync`、、 `removeHandlerAsync`、または) を呼び出して置き換えます。これらのメソッドを呼び出すことは、追加の約束をサポートしていません。ネストされた[コールバック関数パターン](#AsyncProgramming_NestedCallbacks)を使用して呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p124">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](#AsyncProgramming_NestedCallbacks).</span></span>

<span data-ttu-id="50ca1-p125">`Binding`オブジェクトの promise が履行された後は、バインドされている場合と同様に、連鎖メソッドの呼び出しで再利用できます (アドインランタイムは、promise を実行しても非同期に再試行することはありません)。`Binding`オブジェクトの約束を履行できない場合、アドインランタイムは、次にその非同期メソッドが呼び出されたときに binding オブジェクトへのアクセスを再試行します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p125">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="50ca1-193">次のコード例では`select` 、メソッドを使用して、 `id` "`cities`" という`Bindings`バインドをコレクションから取得した後、 [addhandler async](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)メソッドを呼び出して、バインドの[dataChanged](/javascript/api/office/office.bindingdatachangedeventargs)イベントのイベントハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-193">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> <span data-ttu-id="50ca1-p126">メソッドによって返されるオブジェクトの promise は`Binding` 、 `Binding`オブジェクトの4つのメソッドへのアクセスのみを提供します。 `Office.select``Binding`オブジェクトの他のメンバーにアクセスする必要がある場合`Document.bindings`は、その`Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding`オブジェクトを取得するのにプロパティとメソッドのどちらか一方または両方を使用する必要があります。`Binding`たとえば、オブジェクトのプロパティ ( `document` `id`、、または`type`プロパティ) にアクセスする必要がある場合や、 [MatrixBinding](/javascript/api/office/office.matrixbinding)または[tablebinding](/javascript/api/office/office.tablebinding)オブジェクトのプロパティにアクセスする必要がある場合は、 `getByIdAsync`または`getAllAsync`メソッドを使用して`Binding`オブジェクトを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p126">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object. If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object. For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>


## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="50ca1-197">オプションのパラメーターを非同期メソッドに渡す</span><span class="sxs-lookup"><span data-stu-id="50ca1-197">Passing optional parameters to asynchronous methods</span></span>


<span data-ttu-id="50ca1-198">すべての "Async" メソッドの一般的な構文は、次のパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="50ca1-198">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="50ca1-199">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="50ca1-199">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="50ca1-p127">すべての非同期メソッドは、オプションのパラメーターをサポートします。これらは、1 つまたは複数のオプションのパラメーターが格納された JavaScript Object Notation (JSON) オブジェクトとして渡されます。オプションのパラメーターを格納している JSON オブジェクトは、キーと値のペアの順不同のコレクションで、キーと値は ":" 文字で区切られています。オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体はかっこで囲まれます。キーはパラメーター名で、値はそのパラメーターに渡す値です。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p127">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="50ca1-204">省略可能なパラメーターを含む JSON オブジェクトはインラインで作成するか、 `options`オブジェクトを作成し、それを_options_パラメーターとして渡すことによって作成できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-204">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>


### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="50ca1-205">オプションのパラメーターをインラインで渡す</span><span class="sxs-lookup"><span data-stu-id="50ca1-205">Passing optional parameters inline</span></span>

<span data-ttu-id="50ca1-206">たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドを呼び出す場合の構文は、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-206">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="50ca1-207">この呼び出し構文では、 _coercionType_および_asynccontext_という2つのオプションのパラメーターが、かっこで囲まれてインラインで JSON オブジェクトとして定義されています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-207">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="50ca1-208">次の例は、省略可能なパラメーター `Document.setSelectedDataAsync`をインラインで指定することによってメソッドを呼び出す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-208">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>


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
> <span data-ttu-id="50ca1-209">オプションのパラメーターは、名前さえ正しければ、任意の順序で JSON オブジェクトに指定できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-209">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>


### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="50ca1-210">オプションのパラメーターを options オブジェクトで渡す</span><span class="sxs-lookup"><span data-stu-id="50ca1-210">Passing optional parameters in an options object</span></span>

<span data-ttu-id="50ca1-211">または、メソッド呼び出しとは別`options`に、オプションのパラメーターを指定するという名前のオブジェクトを作成`options`し、そのオブジェクトを_options_引数として渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-211">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="50ca1-212">次の例は、 `options`オブジェクトを作成する1つの`parameter1`方法`value1`を示しています。ここで、、、などは、実際のパラメーター名と値のプレースホルダーです。</span><span class="sxs-lookup"><span data-stu-id="50ca1-212">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="50ca1-213">[ValueFormat](/javascript/api/office/office.valueformat) パラメーターおよび [FilterType](/javascript/api/office/office.filtertype) パラメーターを指定する場合は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-213">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="50ca1-214">オブジェクトを`options`作成する別の方法を次に示します。</span><span class="sxs-lookup"><span data-stu-id="50ca1-214">Here's another way of creating the `options` object.</span></span>




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="50ca1-215">パラメーターと`ValueFormat` `FilterType`パラメーターの指定に使用する場合は、次の例のようになります。</span><span class="sxs-lookup"><span data-stu-id="50ca1-215">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> <span data-ttu-id="50ca1-216">いずれかの方法で`options`オブジェクトを作成する場合は、名前が正しく指定されていれば、任意の順序でオプションのパラメーターを指定できます。</span><span class="sxs-lookup"><span data-stu-id="50ca1-216">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="50ca1-217">次の例は、 `Document.setSelectedDataAsync` `options`オブジェクトで省略可能なパラメーターを指定してメソッドを呼び出す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="50ca1-217">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>




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


<span data-ttu-id="50ca1-p128">両方の省略可能なパラメーターの例では、 _callback_パラメーターが最後のパラメーターとして指定されています (インラインオプションのパラメーターの後、または_options_引数オブジェクトの後)。または、_コールバック_パラメーターをインライン JSON オブジェクトまたは`options`オブジェクトの内部で指定することもできます。ただし、 _callback_パラメーターを渡すことができるのは、 _options_オブジェクト (インラインまたは外部で作成されている)、または最後のパラメーター (両方ではない) のいずれかです。</span><span class="sxs-lookup"><span data-stu-id="50ca1-p128">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>


## <a name="see-also"></a><span data-ttu-id="50ca1-221">関連項目</span><span class="sxs-lookup"><span data-stu-id="50ca1-221">See also</span></span>

- [<span data-ttu-id="50ca1-222">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="50ca1-222">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="50ca1-223">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="50ca1-223">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
