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
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み

[Document](/javascript/api/office/office.document) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。 そのために、 `Document` オブジェクトは `getSelectedDataAsync` メソッドとメソッドを提供し `setSelectedDataAsync` ます。 このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。

この `getSelectedDataAsync` メソッドは、ユーザーの現在の選択範囲に対してのみ機能します。 実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) メソッドを使用 (または、[Bindings](/javascript/api/office/office.bindings) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。 ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="read-selected-data"></a>選択されたデータを読み取る


次の例は、ドキュメント内の選択範囲のデータを [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドで取得する方法を示しています。


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

この例では、最初の_coercionType_パラメーターをとして指定し `Office.CoercionType.Text` ます (このパラメーターは、リテラル文字列を使用して指定することもでき `"text"` ます)。 この場合、コールバック関数の [asyncResult](/javascript/api/office/office.asyncresult#status) パラメーターから取得できる [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。 別の型変換を指定すると、別の値が取得されます。 [Office.CoercionType](/javascript/api/office/office.coerciontype) は、使用できる型変換の値を表す列挙型です。 `Office.CoercionType.Text`文字列 "text" に評価されます。


> [!TIP]
> **データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。** 行と列が追加されたときに、選択した表形式のデータを動的に拡張する必要がある場合に、テーブルのヘッダーを操作する必要がある場合は、テーブルのデータ型を使用する必要があります (メソッドの_coercionType_パラメーターをまたはに指定することもでき `getSelectedDataAsync` `"table"` `Office.CoercionType.Table` ます)。 データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。 行と列の追加を計画しておらず、データにヘッダー機能が必要ない場合は、マトリックスデータ型を使用する必要があります (メソッドの_coercionType_パラメーターを指定することにより `getSelectedDataAsync` )。これにより、データを操作する `"matrix"` `Office.CoercionType.Matrix` ためのより簡単なモデルが提供されます。

2番目の_callback_パラメーターとして関数に渡される匿名関数は、操作が完了したときに実行され `getSelectedDataAsync` ます。 この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。 呼び出しが失敗した場合、オブジェクトの[error](/javascript/api/office/office.asyncresult#asynccontext)プロパティは `AsyncResult` [error](/javascript/api/office/office.error)オブジェクトへのアクセスを提供します。 [Error.name](/javascript/api/office/office.error#name) プロパティと [Error.message](/javascript/api/office/office.error#message) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。 呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。

The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded. [AsyncResultStatus](/javascript/api/office/office.asyncresult#status)は、使用可能なプロパティ値の列挙体です `AsyncResult.status` 。 `Office.AsyncResultStatus.Failed`文字列 "failed" に評価されます (つまり、そのリテラル文字列としても指定できます)。


## <a name="write-data-to-the-selection"></a>選択範囲にデータを書き込む


次の例は、"Hello World!" を表示するために選択範囲を設定する方法を示しています。


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

_data_ パラメーターに異なるオブジェクト型を渡すと、結果が異なります。結果は、ドキュメントで現在選択されているもの、アドインをホストしているアプリケーション、および渡されたデータを現在の選択範囲に型変換できるかどうかによって異なります。

The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. メソッドを使用して選択範囲にデータを書き込む場合 `setSelectedDataAsync` 、コールバックの_asyncResult_パラメーターは、呼び出しの状態だけでなく、呼び出しが失敗した場合は[Error](/javascript/api/office/office.error)オブジェクトにアクセスします。

> [!NOTE]
> Excel 2013 SP1 および Excel on the web の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。


## <a name="detect-changes-in-the-selection"></a>選択範囲の変更を検出する


次の例は、[Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) メソッドを使用して、[SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) イベントのイベント ハンドラーをドキュメント上に追加することで、選択範囲の変更を検出する方法を示しています。


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

最初の  _eventType_ パラメーターでは、サブスクライブするイベントの名前を指定しています。 このパラメーターの文字列を渡すこと `"documentSelectionChanged"` は、イベントの種類として `Office.EventType.DocumentSelectionChanged` [EventType](/javascript/api/office/office.eventtype)列挙を渡すことと同じです。

2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](/javascript/api/office/office.documentselectionchangedeventargs) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。


> [!NOTE]
> 特定のイベントに対して複数のイベントハンドラーを追加するには、 `addHandlerAsync` メソッドをもう一度呼び出し、 _handler_パラメーターに対して追加のイベントハンドラー関数を渡します。 この場合、各イベント ハンドラー関数の名前は一意である必要があります。


## <a name="stop-detecting-changes-in-the-selection"></a>選択範囲の変更の検出を中止する


次の例は、[document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) メソッドを呼び出して、[Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) イベントのリッスンを中止する方法を示しています。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

`myHandler`2 番目の_handler_パラメーターとして渡される関数名は、イベントから削除されるイベントハンドラーを指定し `SelectionChanged` ます。


> [!IMPORTANT]
> メソッドを呼び出すときにオプションの_handler_パラメーターを省略すると `removeHandlerAsync` 、指定された_eventType_のすべてのイベントハンドラーが削除されます。
