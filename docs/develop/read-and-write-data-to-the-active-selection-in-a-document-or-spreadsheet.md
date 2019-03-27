---
title: ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 2fe847fcc04e3670db294a421388dbd2faad6f2f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871452"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み

[Document](/javascript/api/office/office.document) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。これは、**Document** オブジェクトの **getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドで行います。このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。


  **getSelectedDataAsync** メソッドは、現在のユーザーの選択範囲のみに実行されます。実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) メソッドを使用 (または、[Bindings](/javascript/api/office/office.bindings) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


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

この例では、最初の _coercionType_ パラメーターには **Office.CoercionType.Text** が指定されています (このパラメーターはリテラル文字列 `"text"` で指定することもできます)。この場合、コールバック関数の [asyncResult](/javascript/api/office/office.asyncresult#status) パラメーターから取得できる [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。別の型変換を指定すると、別の値が取得されます。[Office.CoercionType](/javascript/api/office/office.coerciontype) は、使用できる型変換の値を表す列挙型です。**Office.CoercionType.Text** は文字列 "text" として評価されます。


> [!TIP]
> **データ アクセスにおけるマトリックスとテーブルの coercionType の使い分けについて。** 行と列が追加されたときに選択済みの表形式データが動的に増えるようにし、またテーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用します (**getSelectedDataAsync** メソッドの _coercionType_ パラメーターに `"table"` または **Office.CoercionType.Table** を指定)。データ構造体内での行と列の追加はテーブル データとマトリックス データの両方でサポートされていますが、行と列の追加はテーブル データでのみサポートされています。行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用します (**getSelecteDataAsync** メソッドの _coercionType_パラメーターに `"matrix"` または **Office.CoercionType.Matrix** を指定)。このデータ型では、データとのやり取りについて、より単純なモデルを採用しています。

2 番目の _callback_ パラメーターとして関数に渡される匿名関数は、**getSelectedDataAsync** 操作の完了時に実行されます。この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。呼び出しが失敗した場合は、[AsyncResult](/javascript/api/office/office.asyncresult#asynccontext) オブジェクトの **error** プロパティから [Error](/javascript/api/office/office.error) オブジェクトにアクセスできます。[Error.name](/javascript/api/office/office.error#name) プロパティと [Error.message](/javascript/api/office/office.error#message) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。

[if](/javascript/api/office/office.asyncresult#error) ステートメントでは、呼び出しが成功したかどうかの判定に **AsyncResult.status** プロパティを使用します。[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) は **AsyncResult.status** プロパティが取ることのできる値を表す列挙型です。**Office.AsyncResultStatus.Failed** は文字列 "failed" として評価されます (こちらもリテラル文字列で指定することもできます)。


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

[callback](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) パラメーターとして _setSelectedDataAsync_ メソッドに渡される匿名関数は、非同期呼び出しの完了時に実行されます。**setSelectedDataAsync** メソッドを使用して選択範囲にデータを書き込む場合、コールバックの _asyncResult_ パラメーターは呼び出しのステータスへのアクセスのみを提供し、呼び出しが失敗した場合は[Error](/javascript/api/office/office.error) オブジェクトにアクセスします。

> [!NOTE]
> Excel 2013 SP1 および Excel Online の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。


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

最初の  _eventType_ パラメーターでは、サブスクライブするイベントの名前を指定しています。文字列 `"documentSelectionChanged"` をこのパラメーターに指定するのは、 **Office.EventType** 列挙型のイベントの種類 [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.eventtype) を渡すことに相当します。

2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](/javascript/api/office/office.documentselectionchangedeventargs) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。


> [!NOTE]
> **addHandlerAsync** メソッドを再び呼び出して、_handler_ パラメーターに追加のイベント ハンドラー関数を指定すると、特定のイベントに複数のイベント ハンドラーを追加できます。この場合、各イベント ハンドラー関数の名前は一意である必要があります。


## <a name="stop-detecting-changes-in-the-selection"></a>選択範囲の変更の検出を中止する


次の例は、[document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) メソッドを呼び出して、[Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) イベントのリッスンを中止する方法を示しています。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

2 番目の _handler_ パラメーターとして渡される `myHandler` 関数名は、**SelectionChanged** イベントから削除されるイベント ハンドラーを指定します。


> [!IMPORTANT]
> **removeHandlerAsync** メソッドを呼び出すときにオプションの _handler_ パラメーターを省略すると、指定された _eventType_ のすべてのイベント ハンドラーが削除されます。
