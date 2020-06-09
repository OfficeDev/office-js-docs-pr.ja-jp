---
title: ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み
description: Word 文書または Excel スプレッドシートでアクティブな選択範囲に対してデータを読み書きする方法について説明します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: eb6c4d89e9c66ee3cda012c21601cb7454e73ae8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609395"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="95147-103">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="95147-103">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="95147-104">[Document](/javascript/api/office/office.document) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="95147-104">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet.</span></span> <span data-ttu-id="95147-105">そのために、 `Document` オブジェクトは `getSelectedDataAsync` メソッドとメソッドを提供し `setSelectedDataAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="95147-105">To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods.</span></span> <span data-ttu-id="95147-106">このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="95147-106">This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="95147-107">この `getSelectedDataAsync` メソッドは、ユーザーの現在の選択範囲に対してのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="95147-107">The `getSelectedDataAsync` method only works against the user's current selection.</span></span> <span data-ttu-id="95147-108">実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) メソッドを使用 (または、[Bindings](/javascript/api/office/office.bindings) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="95147-108">If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object).</span></span> <span data-ttu-id="95147-109">ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="95147-109">For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="95147-110">選択されたデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="95147-110">Read selected data</span></span>


<span data-ttu-id="95147-111">次の例は、ドキュメント内の選択範囲のデータを [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドで取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95147-111">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="95147-112">この例では、最初の_coercionType_パラメーターをとして指定し `Office.CoercionType.Text` ます (このパラメーターは、リテラル文字列を使用して指定することもでき `"text"` ます)。</span><span class="sxs-lookup"><span data-stu-id="95147-112">In this example, the first  _coercionType_ parameter is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`).</span></span> <span data-ttu-id="95147-113">この場合、コールバック関数の [asyncResult](/javascript/api/office/office.asyncresult#status) パラメーターから取得できる [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。</span><span class="sxs-lookup"><span data-stu-id="95147-113">This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document.</span></span> <span data-ttu-id="95147-114">別の型変換を指定すると、別の値が取得されます。</span><span class="sxs-lookup"><span data-stu-id="95147-114">Specifying different coercion types will result in different values.</span></span> <span data-ttu-id="95147-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) は、使用できる型変換の値を表す列挙型です。</span><span class="sxs-lookup"><span data-stu-id="95147-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values.</span></span> <span data-ttu-id="95147-116">`Office.CoercionType.Text`文字列 "text" に評価されます。</span><span class="sxs-lookup"><span data-stu-id="95147-116">`Office.CoercionType.Text` evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="95147-117">**データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。**</span><span class="sxs-lookup"><span data-stu-id="95147-117">**When should you use the matrix versus table coercionType for data access?**</span></span> <span data-ttu-id="95147-118">行と列が追加されたときに、選択した表形式のデータを動的に拡張する必要がある場合に、テーブルのヘッダーを操作する必要がある場合は、テーブルのデータ型を使用する必要があります (メソッドの_coercionType_パラメーターをまたはに指定することもでき `getSelectedDataAsync` `"table"` `Office.CoercionType.Table` ます)。</span><span class="sxs-lookup"><span data-stu-id="95147-118">If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`).</span></span> <span data-ttu-id="95147-119">データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="95147-119">Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data.</span></span> <span data-ttu-id="95147-120">行と列の追加を計画しておらず、データにヘッダー機能が必要ない場合は、マトリックスデータ型を使用する必要があります (メソッドの_coercionType_パラメーターを指定することにより `getSelectedDataAsync` )。これにより、データを操作する `"matrix"` `Office.CoercionType.Matrix` ためのより簡単なモデルが提供されます。</span><span class="sxs-lookup"><span data-stu-id="95147-120">If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="95147-121">2番目の_callback_パラメーターとして関数に渡される匿名関数は、操作が完了したときに実行され `getSelectedDataAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="95147-121">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the `getSelectedDataAsync` operation is completed.</span></span> <span data-ttu-id="95147-122">この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="95147-122">The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call.</span></span> <span data-ttu-id="95147-123">呼び出しが失敗した場合、オブジェクトの[error](/javascript/api/office/office.asyncresult#asynccontext)プロパティは `AsyncResult` [error](/javascript/api/office/office.error)オブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="95147-123">If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object.</span></span> <span data-ttu-id="95147-124">[Error.name](/javascript/api/office/office.error#name) プロパティと [Error.message](/javascript/api/office/office.error#message) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。</span><span class="sxs-lookup"><span data-stu-id="95147-124">You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed.</span></span> <span data-ttu-id="95147-125">呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="95147-125">Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="95147-126">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded.</span><span class="sxs-lookup"><span data-stu-id="95147-126">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded.</span></span> <span data-ttu-id="95147-127">[AsyncResultStatus](/javascript/api/office/office.asyncresult#status)は、使用可能なプロパティ値の列挙体です `AsyncResult.status` 。</span><span class="sxs-lookup"><span data-stu-id="95147-127">[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available `AsyncResult.status` property values.</span></span> <span data-ttu-id="95147-128">`Office.AsyncResultStatus.Failed`文字列 "failed" に評価されます (つまり、そのリテラル文字列としても指定できます)。</span><span class="sxs-lookup"><span data-stu-id="95147-128">`Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="95147-129">選択範囲にデータを書き込む</span><span class="sxs-lookup"><span data-stu-id="95147-129">Write data to the selection</span></span>


<span data-ttu-id="95147-130">次の例は、"Hello World!" を表示するために選択範囲を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95147-130">The following example shows how to set the selection to show "Hello World!".</span></span>


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="95147-p107">_data_ パラメーターに異なるオブジェクト型を渡すと、結果が異なります。結果は、ドキュメントで現在選択されているもの、アドインをホストしているアプリケーション、および渡されたデータを現在の選択範囲に型変換できるかどうかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="95147-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="95147-133">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed.</span><span class="sxs-lookup"><span data-stu-id="95147-133">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed.</span></span> <span data-ttu-id="95147-134">メソッドを使用して選択範囲にデータを書き込む場合 `setSelectedDataAsync` 、コールバックの_asyncResult_パラメーターは、呼び出しの状態だけでなく、呼び出しが失敗した場合は[Error](/javascript/api/office/office.error)オブジェクトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="95147-134">When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="95147-135">Excel 2013 SP1 および Excel on the web の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。</span><span class="sxs-lookup"><span data-stu-id="95147-135">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="95147-136">選択範囲の変更を検出する</span><span class="sxs-lookup"><span data-stu-id="95147-136">Detect changes in the selection</span></span>


<span data-ttu-id="95147-137">次の例は、[Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) メソッドを使用して、[SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) イベントのイベント ハンドラーをドキュメント上に追加することで、選択範囲の変更を検出する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95147-137">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


```js
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){}
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="95147-138">最初の  _eventType_ パラメーターでは、サブスクライブするイベントの名前を指定しています。</span><span class="sxs-lookup"><span data-stu-id="95147-138">The first  _eventType_ parameter specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="95147-139">このパラメーターの文字列を渡すこと `"documentSelectionChanged"` は、イベントの種類として `Office.EventType.DocumentSelectionChanged` [EventType](/javascript/api/office/office.eventtype)列挙を渡すことと同じです。</span><span class="sxs-lookup"><span data-stu-id="95147-139">Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="95147-p110">2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](/javascript/api/office/office.documentselectionchangedeventargs) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="95147-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="95147-143">特定のイベントに対して複数のイベントハンドラーを追加するには、 `addHandlerAsync` メソッドをもう一度呼び出し、 _handler_パラメーターに対して追加のイベントハンドラー関数を渡します。</span><span class="sxs-lookup"><span data-stu-id="95147-143">You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter.</span></span> <span data-ttu-id="95147-144">この場合、各イベント ハンドラー関数の名前は一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="95147-144">This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="95147-145">選択範囲の変更の検出を中止する</span><span class="sxs-lookup"><span data-stu-id="95147-145">Stop detecting changes in the selection</span></span>


<span data-ttu-id="95147-146">次の例は、[document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) メソッドを呼び出して、[Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) イベントのリッスンを中止する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="95147-146">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="95147-147">`myHandler`2 番目の_handler_パラメーターとして渡される関数名は、イベントから削除されるイベントハンドラーを指定し `SelectionChanged` ます。</span><span class="sxs-lookup"><span data-stu-id="95147-147">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the `SelectionChanged` event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="95147-148">メソッドを呼び出すときにオプションの_handler_パラメーターを省略すると `removeHandlerAsync` 、指定された_eventType_のすべてのイベントハンドラーが削除されます。</span><span class="sxs-lookup"><span data-stu-id="95147-148">If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
