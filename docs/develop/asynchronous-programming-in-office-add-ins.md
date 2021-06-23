---
title: Office アドインにおける非同期プログラミング
description: JavaScript ライブラリがOfficeアドインで非同期プログラミングを使用する方法Office説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 42cf2d8e1b0d5185866a55152517683031da3b3d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076267"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="5a561-103">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="5a561-103">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="5a561-104">アドイン API Office非同期プログラミングを使用する理由</span><span class="sxs-lookup"><span data-stu-id="5a561-104">Why does the Office Add-ins API use asynchronous programming?</span></span> <span data-ttu-id="5a561-105">JavaScript はシングルスレッドの言語であるため、スクリプトで実行時間の長い同期プロセスが呼び出されると、そのプロセスが完了するまで後続のすべてのスクリプト実行がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="5a561-105">Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes.</span></span> <span data-ttu-id="5a561-106">Office Web クライアント (ただしリッチ クライアントも含む) に対する特定の操作は、同期的に実行される場合に実行をブロックする可能性があるため、Office JavaScript API の大部分は非同期的に実行するように設計されています。</span><span class="sxs-lookup"><span data-stu-id="5a561-106">Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously.</span></span> <span data-ttu-id="5a561-107">これにより、アドインOffice応答性と速度が向上します。</span><span class="sxs-lookup"><span data-stu-id="5a561-107">This makes sure that Office Add-ins are responsive and fast.</span></span> <span data-ttu-id="5a561-108">このような非同期メソッドを利用するときは、多くの場合、コールバック関数の記述も必要です。</span><span class="sxs-lookup"><span data-stu-id="5a561-108">It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="5a561-109">API 内のすべての非同期メソッドの名前は、"Async" (、 、メソッドなど) `Document.getSelectedDataAsync` `Binding.getDataAsync` で `Item.loadCustomPropertiesAsync` 終わる。</span><span class="sxs-lookup"><span data-stu-id="5a561-109">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods.</span></span> <span data-ttu-id="5a561-110">"Async" メソッドは呼び出されるとすぐに実行され、後続のスクリプトも続けて実行することができます。</span><span class="sxs-lookup"><span data-stu-id="5a561-110">When an "Async" method is called, it executes immediately and any subsequent script execution can continue.</span></span> <span data-ttu-id="5a561-111">"Async" メソッドに渡す任意のコールバック関数は、データまたは要求された操作の準備が整い次第、すぐに実行されます。</span><span class="sxs-lookup"><span data-stu-id="5a561-111">The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready.</span></span> <span data-ttu-id="5a561-112">コールバック関数の実行は通常、直ちに行われますが、戻るまでに若干の遅延が生じることがあります。</span><span class="sxs-lookup"><span data-stu-id="5a561-112">This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="5a561-113">次の図は、サーバー ベースの Word または Excel で開いているドキュメントでユーザーが選択したデータを読み取る "Async" メソッドへの呼び出しの実行フローを示しています。</span><span class="sxs-lookup"><span data-stu-id="5a561-113">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel.</span></span> <span data-ttu-id="5a561-114">"Async" 呼び出しが行われた時点で、JavaScript 実行スレッドは追加のクライアント側処理を自由に実行できます (ただし、図には何も表示されません)。</span><span class="sxs-lookup"><span data-stu-id="5a561-114">At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram).</span></span> <span data-ttu-id="5a561-115">"Async" メソッドが返されると、コールバックはスレッドでの実行を再開し、アドインはデータにアクセスし、データにアクセスして何かを行い、結果を表示できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-115">When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result.</span></span> <span data-ttu-id="5a561-116">同じ非同期実行パターンは、Word 2013 や 2013 Officeなど、リッチ クライアント アプリケーションを操作するときに保持Excelします。</span><span class="sxs-lookup"><span data-stu-id="5a561-116">The same asynchronous execution pattern holds when working with the Office rich client applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="5a561-117">*図 1. 非同期プログラミング実行フロー*</span><span class="sxs-lookup"><span data-stu-id="5a561-117">*Figure 1. Asynchronous programming execution flow*</span></span>

![ユーザー、アドイン ページ、およびアドインをホストする Web アプリ サーバーとの時間の間のコマンド実行の相互作用を示す図。](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="5a561-p104">リッチ クライアントと Web クライアントの両方でこの非同期設計をサポートすることは、Office アドイン開発モデルの "write once-run cross-platform (一度書けばどんなプラットフォームも実行できる)" 設計目的の一部です。たとえば、Excel 2013 と Excel Online の両方で実行されるシングル コード ベースのコンテンツ アドインまたは作業ウィンドウ アドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="5a561-121">"Async" メソッドのコールバック関数を記述する</span><span class="sxs-lookup"><span data-stu-id="5a561-121">Writing the callback function for an "Async" method</span></span>

<span data-ttu-id="5a561-122">コールバック引数として "Async" メソッドに渡すコールバック関数は、コールバック関数の実行時にアドイン ランタイムが[AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトへのアクセスを提供するために使用する 1 つのパラメーターを宣言する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="5a561-122">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes.</span></span> <span data-ttu-id="5a561-123">次のように記述することができます。</span><span class="sxs-lookup"><span data-stu-id="5a561-123">You can write:</span></span>

- <span data-ttu-id="5a561-124">"Async" メソッドのコールバック パラメーターとして"Async" メソッドへの呼び出しに沿って直接書き込み、渡す必要がある匿名関数。</span><span class="sxs-lookup"><span data-stu-id="5a561-124">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="5a561-125">"Async" メソッドのコールバック パラメーターとしてその関数の名前を渡す名前付き関数。</span><span class="sxs-lookup"><span data-stu-id="5a561-125">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="5a561-p106">匿名関数は、そのコードを一度だけ使用する場合に便利です。関数には名前がないため、コードの別の部分で参照できないためです。名前付き関数は、複数の "Async" メソッドにコールバック関数を再利用する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="5a561-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>

### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="5a561-128">匿名コールバック関数を記述する</span><span class="sxs-lookup"><span data-stu-id="5a561-128">Writing an anonymous callback function</span></span>

<span data-ttu-id="5a561-129">次の匿名コールバック関数は、コールバックが返されたときに `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティからデータを取得するという名前の単一のパラメーターを宣言します。</span><span class="sxs-lookup"><span data-stu-id="5a561-129">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="5a561-130">次の例は、完全な "Async" メソッド呼び出しのコンテキストで、この匿名コールバック関数を行に渡す方法を示 `Document.getSelectedDataAsync` しています。</span><span class="sxs-lookup"><span data-stu-id="5a561-130">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>

- <span data-ttu-id="5a561-131">最初の _coercionType_ 引数 , は、選択したデータを文字列として `Office.CoercionType.Text` 返す値を指定します。</span><span class="sxs-lookup"><span data-stu-id="5a561-131">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="5a561-132">2 _番目のコールバック_ 引数は、メソッドに行で渡される匿名関数です。</span><span class="sxs-lookup"><span data-stu-id="5a561-132">The second _callback_ argument is the anonymous function passed in-line to the method.</span></span> <span data-ttu-id="5a561-133">関数を実行すると _、result_ パラメーターを使用してオブジェクトのプロパティにアクセスし、ユーザーが選択したデータをドキュメント `value` `AsyncResult` に表示します。</span><span class="sxs-lookup"><span data-stu-id="5a561-133">When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>

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

<span data-ttu-id="5a561-134">コールバック関数のパラメーターを使用して、オブジェクトの他のプロパティにアクセス `AsyncResult` することもできます。</span><span class="sxs-lookup"><span data-stu-id="5a561-134">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object.</span></span> <span data-ttu-id="5a561-135">呼び出しの成功または失敗を判断する場合は [AsyncResult.status](/javascript/api/office/office.asyncresult#status) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="5a561-135">Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed.</span></span> <span data-ttu-id="5a561-136">呼び出しが失敗した場合は [AsyncResult.error](/javascript/api/office/office.asyncresult#error) プロパティを使用して [Error](/javascript/api/office/office.error) オブジェクトにアクセスし、エラーの詳細を確認できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-136">If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="5a561-137">メソッドの使用の詳細については、「ドキュメントまたはスプレッドシート内のアクティブな選択範囲に対するデータの読み取りおよび書き `getSelectedDataAsync` [込み」を参照してください](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="5a561-137">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 

### <a name="writing-a-named-callback-function"></a><span data-ttu-id="5a561-138">名前付き関数を記述する</span><span class="sxs-lookup"><span data-stu-id="5a561-138">Writing a named callback function</span></span>

<span data-ttu-id="5a561-139">または、名前付き関数を記述し、その名前を "Async" メソッドの _コールバック_ パラメーターに渡します。</span><span class="sxs-lookup"><span data-stu-id="5a561-139">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method.</span></span> <span data-ttu-id="5a561-140">たとえば、前の例は次のように `writeDataCallback` という名前の関数を _callback_ パラメーターとして渡すように書き換えることができます。</span><span class="sxs-lookup"><span data-stu-id="5a561-140">For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>

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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="5a561-141">AsyncResult.value プロパティに返される内容の違い</span><span class="sxs-lookup"><span data-stu-id="5a561-141">Differences in what's returned to the AsyncResult.value property</span></span>

<span data-ttu-id="5a561-142">オブジェクトの 、 、 プロパティは、すべての "Async" メソッドに渡されるコールバック関数に同じ種類の情報 `asyncContext` `status` `error` `AsyncResult` を返します。</span><span class="sxs-lookup"><span data-stu-id="5a561-142">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods.</span></span> <span data-ttu-id="5a561-143">ただし、プロパティに返される内容は `AsyncResult.value` 、"Async" メソッドの機能によって異なります。</span><span class="sxs-lookup"><span data-stu-id="5a561-143">However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="5a561-144">たとえば、これらのオブジェクトで表されるアイテムにイベント ハンドラー関数を追加するには、メソッド `addHandlerAsync` (Binding、CustomXmlPart、Document、RoamingSettings、および[設定](/javascript/api/office/office.settings)オブジェクト) を使用[](/javascript/api/outlook/office.roamingsettings)[](/javascript/api/office/office.document)[](/javascript/api/office/office.binding)[](/javascript/api/office/office.customxmlpart)します。</span><span class="sxs-lookup"><span data-stu-id="5a561-144">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects.</span></span> <span data-ttu-id="5a561-145">任意のメソッドに渡すコールバック関数からプロパティにアクセスできますが、イベント ハンドラーを追加するときにデータやオブジェクトがアクセスされていないので、アクセスしようとすると、プロパティは常に未定義を返します。 `AsyncResult.value` `addHandlerAsync` `value` </span><span class="sxs-lookup"><span data-stu-id="5a561-145">You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="5a561-146">一方、メソッドを呼び出した場合は、ドキュメントで選択したユーザーのデータをコールバック `Document.getSelectedDataAsync` `AsyncResult.value` 内のプロパティに返します。</span><span class="sxs-lookup"><span data-stu-id="5a561-146">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback.</span></span> <span data-ttu-id="5a561-147">または [、Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) メソッドを呼び出した場合は、ドキュメント内のすべてのオブジェクトの配列 `Binding` を返します。</span><span class="sxs-lookup"><span data-stu-id="5a561-147">Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document.</span></span> <span data-ttu-id="5a561-148">[Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)メソッドを呼び出した場合は、1 つのオブジェクトを返 `Binding` します。</span><span class="sxs-lookup"><span data-stu-id="5a561-148">And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="5a561-149">メソッドのプロパティに返される内容の説明については、そのメソッドのリファレンス トピックの `AsyncResult.value` `Async` 「Callback value」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a561-149">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic.</span></span> <span data-ttu-id="5a561-150">メソッドを提供するオブジェクトの概要については `Async` [、AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトのトピックの下部にある表を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a561-150">For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>

## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="5a561-151">非同期プログラミング パターン</span><span class="sxs-lookup"><span data-stu-id="5a561-151">Asynchronous programming patterns</span></span>

<span data-ttu-id="5a561-152">JavaScript API Officeは、次の 2 種類の非同期プログラミング パターンをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="5a561-152">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>

- <span data-ttu-id="5a561-153">入れ子のコールバックの使用</span><span class="sxs-lookup"><span data-stu-id="5a561-153">Using nested callbacks</span></span>
- <span data-ttu-id="5a561-154">promise パターンの使用</span><span class="sxs-lookup"><span data-stu-id="5a561-154">Using the promises pattern</span></span>

<span data-ttu-id="5a561-p114">コールバック関数のある非同期プログラミングでは、多くの場合、2 つ以上のコールバック内に 1 つのコールバックで返された結果を入れ子にすることが必要となります。その場合、API のすべての "Async" メソッドからの入れ子のコールバックを使用できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="5a561-157">入れ子のコールバックを使用することは、ほとんどの JavaScript 開発者にとってなじみのあるプログラミング パターンですが、コールバックが深い入れ子になっているコードは読みにくく、理解しにくいものです。</span><span class="sxs-lookup"><span data-stu-id="5a561-157">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand.</span></span> <span data-ttu-id="5a561-158">入れ子になったコールバックの代わりに、javaScript API Office約束パターンの実装もサポートしています。</span><span class="sxs-lookup"><span data-stu-id="5a561-158">As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern.</span></span>

> [!NOTE]
> <span data-ttu-id="5a561-159">Office JavaScript API の現在のバージョンでは、promises パターンの組み込みサポートは、Excel スプレッドシートと Word ドキュメントのバインド[のコードでのみ機能します](bind-to-regions-in-a-document-or-spreadsheet.md)。 </span><span class="sxs-lookup"><span data-stu-id="5a561-159">In the current version of the Office JavaScript API, *built-in* support for the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span> <span data-ttu-id="5a561-160">ただし、独自のカスタム Promise-returning 関数内にコールバックを持つ他の関数をラップできます。</span><span class="sxs-lookup"><span data-stu-id="5a561-160">However, you can wrap other functions that have callbacks inside your own custom Promise-returning function.</span></span> <span data-ttu-id="5a561-161">詳細については [、「Promise-returning 関数で一般的な API をラップする」を参照してください](#wrap-common-apis-in-promise-returning-functions)。</span><span class="sxs-lookup"><span data-stu-id="5a561-161">For more information, see [Wrap Common APIs in Promise-returning functions](#wrap-common-apis-in-promise-returning-functions).</span></span>

### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="5a561-162">入れ子のコールバック関数を使用する非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="5a561-162">Asynchronous programming using nested callback functions</span></span>

<span data-ttu-id="5a561-p117">多くの場合、タスクを完了するには、2 つ以上の非同期操作を実行する必要があります。これを実現するために、1 つの "Async" 呼び出し内で別の呼び出しを入れ子にできます。</span><span class="sxs-lookup"><span data-stu-id="5a561-p117">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="5a561-165">次のコード例では、2 つの非同期呼び出しを入れ子にしています。</span><span class="sxs-lookup"><span data-stu-id="5a561-165">The following code example nests two asynchronous calls.</span></span>

- <span data-ttu-id="5a561-166">最初に、[Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) メソッドが呼び出され、"MyBinding" という名前のドキュメントのバインドにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="5a561-166">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding".</span></span> <span data-ttu-id="5a561-167">その `AsyncResult` コールバックのパラメーターに返 `result` されるオブジェクトは、プロパティから指定されたバインド オブジェクトへのアクセスを提供 `AsyncResult.value` します。</span><span class="sxs-lookup"><span data-stu-id="5a561-167">The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>
- <span data-ttu-id="5a561-168">次に、最初のパラメーターからアクセスするバインド オブジェクトを使用して `result` [Binding.getDataAsync メソッドを呼び出](/javascript/api/office/office.binding#getdataasync-options--callback-) します。</span><span class="sxs-lookup"><span data-stu-id="5a561-168">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>
- <span data-ttu-id="5a561-169">最後に、メソッド `result2` に渡されるコールバックのパラメーターを使用して、バインド `Binding.getDataAsync` 内のデータを表示します。</span><span class="sxs-lookup"><span data-stu-id="5a561-169">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>

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

<span data-ttu-id="5a561-170">この基本的な入れ子になったコールバック パターンは、JavaScript API 内のすべての非同期メソッドOffice使用できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-170">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="5a561-171">次のセクションでは、非同期メソッドの入れ子のコールバックで匿名関数または名前付き関数を使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="5a561-171">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>

#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="5a561-172">入れ子のコールバックとして匿名関数を使用する</span><span class="sxs-lookup"><span data-stu-id="5a561-172">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="5a561-173">次の例では、2 つの匿名関数がインラインで宣言され、and メソッドに入れ子になった `getByIdAsync` `getDataAsync` コールバックとして渡されます。</span><span class="sxs-lookup"><span data-stu-id="5a561-173">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks.</span></span> <span data-ttu-id="5a561-174">関数は単純でインラインのため、実装の意図は明白です。</span><span class="sxs-lookup"><span data-stu-id="5a561-174">Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>

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

#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="5a561-175">入れ子のコールバックとして名前付き関数を使用する</span><span class="sxs-lookup"><span data-stu-id="5a561-175">Using named functions for nested callbacks</span></span>

<span data-ttu-id="5a561-176">複雑な実装の場合、名前付き関数を使用すると、読みやすく、保守管理がしやすく、再利用しやすくなります。</span><span class="sxs-lookup"><span data-stu-id="5a561-176">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse.</span></span> <span data-ttu-id="5a561-177">次の例では、前のセクションの例の 2 つの匿名関数が、という名前の関数として書き換 `deleteAllData` えされています `showResult` 。</span><span class="sxs-lookup"><span data-stu-id="5a561-177">In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`.</span></span> <span data-ttu-id="5a561-178">これらの名前付き関数は、名前によって `getByIdAsync` コールバックとして and `deleteAllDataValuesAsync` メソッドに渡されます。</span><span class="sxs-lookup"><span data-stu-id="5a561-178">These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>

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

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="5a561-179">promise パターンを使用してバインドのデータにアクセスする非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="5a561-179">Asynchronous programming using the promises pattern to access data in bindings</span></span>

<span data-ttu-id="5a561-p121">コールバック関数を渡し、その関数が戻るのを待ってから実行を続行する代わりに、promise プログラミング パターンを使用すれば、その意図した結果を表す promise オブジェクトがすぐに返されます。ただし、本物の同期プログラミングとは異なり、実際には Office アドインのランタイム環境が要求を完了できるまでは、約束された結果の履行は実際には延期されます。要求が履行されない状況に対処するために _onError_ ハンドラーが用意されています。</span><span class="sxs-lookup"><span data-stu-id="5a561-p121">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>

<span data-ttu-id="5a561-183">JavaScript API Officeは[、Office.select](/javascript/api/office#office-select-expression--callback-)メソッドを使用して、既存のバインド オブジェクトを操作する約束パターンをサポートします。</span><span class="sxs-lookup"><span data-stu-id="5a561-183">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects.</span></span> <span data-ttu-id="5a561-184">メソッドに返される promise オブジェクトは、Binding オブジェクトから直接アクセスできる 4 つのメソッド `Office.select` [(getDataAsync、setDataAsync、addHandlerAsync、removeHandlerAsync)](/javascript/api/office/office.binding#getdataasync-options--callback-)のみを[サポート](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)[](/javascript/api/office/office.binding)[](/javascript/api/office/office.binding#setdataasync-data--options--callback-)[](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)します。</span><span class="sxs-lookup"><span data-stu-id="5a561-184">The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>

<span data-ttu-id="5a561-185">バインドと連携する promise パターンは次のような形式になります。</span><span class="sxs-lookup"><span data-stu-id="5a561-185">The promises pattern for working with bindings takes this form:</span></span>

<span data-ttu-id="5a561-186">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="5a561-186">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="5a561-187">_selectorExpression_ パラメーターはフォームを受け取ります。bindingId は、ドキュメントまたはスプレッドシートで以前に作成したバインドの名前 ( ) です (コレクションの `"bindings#bindingId"`  `id` "addFrom" メソッドの 1 つを使用します `Bindings` 。 `addFromNamedItemAsync` `addFromPromptAsync` `addFromSelectionAsync`</span><span class="sxs-lookup"><span data-stu-id="5a561-187">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`).</span></span> <span data-ttu-id="5a561-188">たとえば、セレクター式は、'cities' の ID を持つバインドにアクセス `bindings#cities` する場合に指定します。 </span><span class="sxs-lookup"><span data-stu-id="5a561-188">For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="5a561-189">_onError パラメーターは_、メソッドが指定されたバインドにアクセスできない場合に、オブジェクトにアクセスするために使用できる型の 1 つのパラメーターを受け取るエラー処理 `AsyncResult` `Error` `select` 関数です。</span><span class="sxs-lookup"><span data-stu-id="5a561-189">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding.</span></span> <span data-ttu-id="5a561-190">次の例は、_onError_ パラメーターに渡すことができる基本的なエラー処理関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="5a561-190">The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>

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

<span data-ttu-id="5a561-191">_BindingObjectAsyncMethod_ プレースホルダーを promise オブジェクトでサポートされている 4 つのオブジェクト メソッドの呼び出しに置き換 `Binding` える: `getDataAsync` `setDataAsync` 、、、 `addHandlerAsync` または `removeHandlerAsync` 。</span><span class="sxs-lookup"><span data-stu-id="5a561-191">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`.</span></span> <span data-ttu-id="5a561-192">これらのメソッドの呼び出しでは追加の promise がサポートされません。</span><span class="sxs-lookup"><span data-stu-id="5a561-192">Calls to these methods don't support additional promises.</span></span> <span data-ttu-id="5a561-193">これらは[入れ子のコールバック関数パターン](#asynchronous-programming-using-nested-callback-functions)を使用して呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="5a561-193">You must call them using the [nested callback function pattern](#asynchronous-programming-using-nested-callback-functions).</span></span>

<span data-ttu-id="5a561-194">オブジェクトの約束が満たされた後、チェーンメソッド呼び出しで、バインドである場合と同様に再利用できます (アドイン ランタイムは、約束を満たすのを非同期的に再試行する必要はありません `Binding` )。</span><span class="sxs-lookup"><span data-stu-id="5a561-194">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise).</span></span> <span data-ttu-id="5a561-195">オブジェクトの約束を満たすことができない場合、次に非同期メソッドの 1 つが呼び出されると、アドイン ランタイムはバインド オブジェクトへのアクセスを再試行 `Binding` します。</span><span class="sxs-lookup"><span data-stu-id="5a561-195">If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="5a561-196">次のコード例では、メソッドを使用してコレクションから " とバインドを取得し `select` `id` `cities` `Bindings` [、addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出して、バインドの [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="5a561-196">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> <span data-ttu-id="5a561-197">メソッド `Binding` によって返されるオブジェクトの約束 `Office.select` は、オブジェクトの 4 つのメソッドにのみアクセス `Binding` できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-197">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object.</span></span> <span data-ttu-id="5a561-198">オブジェクトの他のメンバーにアクセスする必要がある場合は、代わりにプロパティまたはメソッドを使用してオブジェクト `Binding` `Document.bindings` `Bindings.getByIdAsync` `Bindings.getAllAsync` を取得する必要 `Binding` があります。</span><span class="sxs-lookup"><span data-stu-id="5a561-198">If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object.</span></span> <span data-ttu-id="5a561-199">たとえば、オブジェクトのプロパティ (、 、またはプロパティ) にアクセスする必要がある場合、 `Binding` `document` または `id` MatrixBinding オブジェクトまたは `type` [TableBinding](/javascript/api/office/office.matrixbinding)[](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` `Binding` オブジェクトのプロパティにアクセスする必要がある場合は、or メソッドを使用してオブジェクトを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5a561-199">For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>

## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="5a561-200">オプションのパラメーターを非同期メソッドに渡す</span><span class="sxs-lookup"><span data-stu-id="5a561-200">Passing optional parameters to asynchronous methods</span></span>

<span data-ttu-id="5a561-201">すべての "Async" メソッドの一般的な構文は、次のパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="5a561-201">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="5a561-202">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="5a561-202">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="5a561-p128">すべての非同期メソッドは、オプションのパラメーターをサポートします。これらは、1 つまたは複数のオプションのパラメーターが格納された JavaScript Object Notation (JSON) オブジェクトとして渡されます。オプションのパラメーターを格納している JSON オブジェクトは、キーと値のペアの順不同のコレクションで、キーと値は ":" 文字で区切られています。オブジェクト内の各ペアはコンマで区切られ、ペアのセット全体はかっこで囲まれます。キーはパラメーター名で、値はそのパラメーターに渡す値です。</span><span class="sxs-lookup"><span data-stu-id="5a561-p128">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="5a561-207">オプションのパラメーターをインラインで含む JSON オブジェクトを作成するか、オブジェクトを作成して options パラメーター `options` として渡します。 </span><span class="sxs-lookup"><span data-stu-id="5a561-207">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>

### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="5a561-208">オプションのパラメーターをインラインで渡す</span><span class="sxs-lookup"><span data-stu-id="5a561-208">Passing optional parameters inline</span></span>

<span data-ttu-id="5a561-209">たとえば、オプションのパラメーターをインラインで指定して [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドを呼び出す場合の構文は、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="5a561-209">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="5a561-210">呼び出し元の構文のこの形式では _、coercionType_ と _asyncContext_ という 2 つの省略可能なパラメーターは、中かっこでインラインで囲まれた JSON オブジェクトとして定義されます。</span><span class="sxs-lookup"><span data-stu-id="5a561-210">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="5a561-211">次の例は、オプションのパラメーターをインラインで `Document.setSelectedDataAsync` 指定してメソッドを呼び出す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5a561-211">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>

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
> <span data-ttu-id="5a561-212">オプションのパラメーターは、名前さえ正しければ、任意の順序で JSON オブジェクトに指定できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-212">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>

### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="5a561-213">オプションのパラメーターを options オブジェクトで渡す</span><span class="sxs-lookup"><span data-stu-id="5a561-213">Passing optional parameters in an options object</span></span>

<span data-ttu-id="5a561-214">または、省略可能なパラメーターをメソッド呼び出しとは別に指定するオブジェクトを作成し、options 引数としてオブジェクト `options` `options` を _渡_ します。</span><span class="sxs-lookup"><span data-stu-id="5a561-214">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="5a561-215">次の例は、オブジェクトを作成する方法の 1 つを示しています。ここで、実際のパラメーター名と値のプレースホルダー `options` `parameter1` `value1` です。</span><span class="sxs-lookup"><span data-stu-id="5a561-215">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="5a561-216">[ValueFormat](/javascript/api/office/office.valueformat) パラメーターおよび [FilterType](/javascript/api/office/office.filtertype) パラメーターを指定する場合は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="5a561-216">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="5a561-217">オブジェクトを作成する別の方法を次に示 `options` します。</span><span class="sxs-lookup"><span data-stu-id="5a561-217">Here's another way of creating the `options` object.</span></span>

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="5a561-218">and パラメーターの指定に使用する場合、次の例のようになります `ValueFormat` `FilterType` 。</span><span class="sxs-lookup"><span data-stu-id="5a561-218">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> <span data-ttu-id="5a561-219">オブジェクトを作成する方法を使用する場合は、名前が正しく指定されている限り、任意の順序でオプションのパラメーター `options` を指定できます。</span><span class="sxs-lookup"><span data-stu-id="5a561-219">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="5a561-220">次の例は、オブジェクトでオプションのパラメーター `Document.setSelectedDataAsync` を指定してメソッドを呼び出す方法を示 `options` しています。</span><span class="sxs-lookup"><span data-stu-id="5a561-220">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>

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

<span data-ttu-id="5a561-221">どちらのオプションのパラメーター例でも、_callback_ パラメーターが最後のパラメーターとして (インラインのオプションのパラメーターまたは _options_ 引数オブジェクトに続けて) 指定されています。</span><span class="sxs-lookup"><span data-stu-id="5a561-221">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object).</span></span> <span data-ttu-id="5a561-222">_callback_ パラメーターは、インライン JSON オブジェクトの中で、または `options` オブジェクト内で指定することもできます。</span><span class="sxs-lookup"><span data-stu-id="5a561-222">Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object.</span></span> <span data-ttu-id="5a561-223">ただし、 _callback_ パラメーターを渡せるのは _options_ オブジェクト (インラインまたは外部で作成) または最後のパラメーターのどちらか一方であり、両方に渡すことはできません。</span><span class="sxs-lookup"><span data-stu-id="5a561-223">However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>

## <a name="wrap-common-apis-in-promise-returning-functions"></a><span data-ttu-id="5a561-224">Promise-returning 関数で一般的な API をラップする</span><span class="sxs-lookup"><span data-stu-id="5a561-224">Wrap Common APIs in Promise-returning functions</span></span>

<span data-ttu-id="5a561-225">Common API (および Outlook API) メソッドは Promises を[返します](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。</span><span class="sxs-lookup"><span data-stu-id="5a561-225">The Common API (and Outlook API) methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="5a561-226">したがって、非同期操作が完了するまで [、await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) を使用して実行を一時停止することはできません。</span><span class="sxs-lookup"><span data-stu-id="5a561-226">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="5a561-227">動作が必要 `await` な場合は、明示的に作成された Promise でメソッド呼び出しをラップできます。</span><span class="sxs-lookup"><span data-stu-id="5a561-227">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span> 

<span data-ttu-id="5a561-228">基本的なパターンは、Promise オブジェクトをすぐに返す非同期メソッドを作成し、内部メソッドの完了時に Promise オブジェクトを解決するか、メソッドが失敗した場合にオブジェクトを拒否します。</span><span class="sxs-lookup"><span data-stu-id="5a561-228">The basic pattern is to create an asynchronous method that returns a Promise object immediately and *resolves* that Promise object when the inner method completes, or *rejects* the object if the method fails.</span></span> <span data-ttu-id="5a561-229">次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="5a561-229">The following is a simple example</span></span>

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

<span data-ttu-id="5a561-230">このメソッドを待つ必要がある場合は、キーワードを使用するか、関数に渡 `await` す関数として呼び出 `then` すことができます。</span><span class="sxs-lookup"><span data-stu-id="5a561-230">When this method needs to be awaited, it can be called either with the `await` keyword or as the function passed to a `then` function.</span></span>

> [!NOTE]
> <span data-ttu-id="5a561-231">この手法は、アプリケーション固有のオブジェクト モデルの 1 つでメソッドの呼び出し内で共通 API の 1 つを呼び出す必要がある場合に特に `run` 便利です。</span><span class="sxs-lookup"><span data-stu-id="5a561-231">This technique is especially useful when you need to call one of the Common APIs inside a call of the `run` method in one of the application-specific object models.</span></span> <span data-ttu-id="5a561-232">この方法で使用されている上記の関数の例については、 [ サンプル Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)のファイルHome.jsを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5a561-232">For an example of the function above being used in this way, see the file [Home.js in the sample Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).</span></span>

<span data-ttu-id="5a561-233">TypeScript を使用する例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="5a561-233">The following is an example using TypeScript.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5a561-234">関連項目</span><span class="sxs-lookup"><span data-stu-id="5a561-234">See also</span></span>

- [<span data-ttu-id="5a561-235">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="5a561-235">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="5a561-236">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="5a561-236">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
