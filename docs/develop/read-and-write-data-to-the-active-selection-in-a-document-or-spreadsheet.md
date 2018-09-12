---
title: ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d1c8fcdeec8d92fd3f77e169dc24715f7c5e9964
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944987"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="cb1ba-102">ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み</span><span class="sxs-lookup"><span data-stu-id="cb1ba-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="cb1ba-p101">[Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。これは、**Document** オブジェクトの **getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドで行います。このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p101">The [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="cb1ba-p102">
  \*\*getSelectedDataAsync\*\* メソッドは、現在のユーザーの選択範囲のみに実行されます。実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) メソッドを使用 (または、[Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p102">The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="cb1ba-109">選択されたデータを読み取る</span><span class="sxs-lookup"><span data-stu-id="cb1ba-109">Read selected data</span></span>


<span data-ttu-id="cb1ba-110">次の例は、ドキュメント内の選択範囲のデータを [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) メソッドで取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="cb1ba-p103">この例では、最初の _coercionType_ パラメーターには **Office.CoercionType.Text** が指定されています (このパラメーターはリテラル文字列 `"text"` で指定することもできます)。この場合、コールバック関数の [asyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) パラメーターから取得できる [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。別の型変換を指定すると、別の値が取得されます。[Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) は、使用できる型変換の値を表す列挙型です。**Office.CoercionType.Text** は文字列 "text" として評価されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p103">In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) property of the [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="cb1ba-p104">**データ アクセスにおけるマトリックスとテーブルの coercionType の使い分けについて。** 行と列が追加されたときに選択済みの表形式データが動的に増えるようにし、またテーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用します (**getSelectedDataAsync** メソッドの _coercionType_ パラメーターに `"table"` または **Office.CoercionType.Table** を指定)。データ構造体内での行と列の追加はテーブル データとマトリックス データの両方でサポートされていますが、行と列の追加はテーブル データでのみサポートされています。行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用します (**getSelecteDataAsync** メソッドの _coercionType_パラメーターに `"matrix"` または **Office.CoercionType.Matrix** を指定)。このデータ型では、データとのやり取りについて、より単純なモデルを採用しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="cb1ba-p105">2 番目の _callback_ パラメーターとして関数に渡される匿名関数は、**getSelectedDataAsync** 操作の完了時に実行されます。この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。呼び出しが失敗した場合は、[AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) オブジェクトの **error** プロパティから [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) オブジェクトにアクセスできます。[Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) プロパティと [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) property of the **AsyncResult** object provides access to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object. You can check the value of the [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) and [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="cb1ba-p106">[if](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) ステートメントでは、呼び出しが成功したかどうかの判定に **AsyncResult.status** プロパティを使用します。[Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) は **AsyncResult.status** プロパティが取ることのできる値を表す列挙型です。**Office.AsyncResultStatus.Failed** は文字列 "failed" として評価されます (こちらもリテラル文字列で指定することもできます)。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p106">The [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="cb1ba-128">選択範囲にデータを書き込む</span><span class="sxs-lookup"><span data-stu-id="cb1ba-128">Write data to the selection</span></span>


<span data-ttu-id="cb1ba-129">次の例は、"Hello World!" を表示するために選択範囲を設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-129">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="cb1ba-p107">_data_ パラメーターに異なるオブジェクト型を渡すと、結果が異なります。結果は、ドキュメントで現在選択されているもの、アドインをホストしているアプリケーション、および渡されたデータを現在の選択範囲に型変換できるかどうかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="cb1ba-p108">[callback](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) パラメーターとして _setSelectedDataAsync_ メソッドに渡される匿名関数は、非同期呼び出しの完了時に実行されます。**setSelectedDataAsync** メソッドを使用して選択範囲にデータを書き込む場合、コールバックの _asyncResult_ パラメーターは呼び出しのステータスへのアクセスのみを提供し、呼び出しが失敗した場合は[Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) オブジェクトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p108">The anonymous function passed into the [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ba-134">Excel 2013 SP1 および Excel Online の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="cb1ba-135">選択範囲の変更を検出する</span><span class="sxs-lookup"><span data-stu-id="cb1ba-135">Detect changes in the selection</span></span>


<span data-ttu-id="cb1ba-136">次の例は、[Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) メソッドを使用して、[SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) イベントのイベント ハンドラーをドキュメント上に追加することで、選択範囲の変更を検出する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event on the document.</span></span>


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

<span data-ttu-id="cb1ba-p109">最初の  _eventType_ パラメーターでは、サブスクライブするイベントの名前を指定しています。文字列 `"documentSelectionChanged"` をこのパラメーターに指定するのは、 **Office.EventType** 列挙型のイベントの種類 [Office.EventType.DocumentSelectionChanged](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js) を渡すことに相当します。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js) enumeration.</span></span>

<span data-ttu-id="cb1ba-p110">2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="cb1ba-p111">**addHandlerAsync** メソッドを再び呼び出して、_handler_ パラメーターに追加のイベント ハンドラー関数を指定すると、特定のイベントに複数のイベント ハンドラーを追加できます。この場合、各イベント ハンドラー関数の名前は一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-p111">You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="cb1ba-144">選択範囲の変更の検出を中止する</span><span class="sxs-lookup"><span data-stu-id="cb1ba-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="cb1ba-145">次の例は、[document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) メソッドを呼び出して、[Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-) イベントのリッスンを中止する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-145">The following example shows how to stop listening to the [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event by calling the [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="cb1ba-146">2 番目の _handler_ パラメーターとして渡される `myHandler` 関数名は、**SelectionChanged** イベントから削除されるイベント ハンドラーを指定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="cb1ba-147">**removeHandlerAsync** メソッドを呼び出すときにオプションの _handler_ パラメーターを省略すると、指定された _eventType_ のすべてのイベント ハンドラーが削除されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ba-147">If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.</span></span>

