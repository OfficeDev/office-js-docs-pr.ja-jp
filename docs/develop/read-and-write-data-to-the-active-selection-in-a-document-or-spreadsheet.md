---
title: ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 039631e935d2ff6fadb4eab9d99df73ac30dae4d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325004"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み

[Document](/javascript/api/office/office.document)オブジェクトは、ドキュメントまたはスプレッドシートでユーザーの現在の選択範囲を読み書きできるメソッドを公開します。そのために、オブジェクト`Document`はメソッド`getSelectedDataAsync`と`setSelectedDataAsync`メソッドを提供します。このトピックでは、ユーザーの選択内容の変更を検出するためのイベントハンドラーの読み取り、書き込み、および作成の方法についても説明します。

この`getSelectedDataAsync`メソッドは、ユーザーの現在の選択範囲に対してのみ機能します。ドキュメントの選択を保持する必要がある場合は、アドインを実行するセッション間で同じ選択を読み書きできるようにするには、Bindings メソッドを使用してバインドを追加する必要があり[ます。](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-)また[は、binding オブジェクトの](/javascript/api/office/office.bindings)他の "addfrom" メソッドのいずれかを使用してバインドを作成します。ドキュメントの領域へのバインドを作成し、バインドの読み取りおよび書き込みを行う方法については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


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

この例では、最初の_coercionType_パラメーターをと`Office.CoercionType.Text`して指定します (このパラメーターは、リテラル文字列`"text"`を使用して指定することもできます)。つまり、コールバック関数の_asyncresult_パラメーターから使用できる[asyncresult](/javascript/api/office/office.asyncresult)オブジェクトの[value](/javascript/api/office/office.asyncresult#status)プロパティは、ドキュメント内で選択されたテキストを含む**文字列**を返します。さまざまな強制型変換を指定すると、別の値になります。[CoercionType](/javascript/api/office/office.coerciontype)は、使用可能な強制型変換型の値の列挙体です。`Office.CoercionType.Text`文字列 "text" に評価されます。


> [!TIP]
> **データアクセスにマトリックスとテーブル coercionType のどちらを使用する必要があるか。** 行と列が追加されたときに、選択した表形式のデータを動的に拡張する必要がある場合に、テーブルのヘッダーを操作する必要がある場合は`getSelectedDataAsync` 、テーブルのデータ型を使用する必要があります (メソッドの`"table"` _coercionType_パラメーターをまたは`Office.CoercionType.Table`に指定することもできます)。データ構造内での行と列の追加は、テーブルデータとマトリックスデータの両方でサポートされていますが、行と列の追加はテーブルデータに対してのみサポートされています。行と列の追加を計画しておらず、データにヘッダー機能が必要ない場合は、マトリックスデータ型を使用する必要があります__ (メソッド`getSelectedDataAsync` `"matrix"`の coercionType パラメーターを`Office.CoercionType.Matrix`指定することにより)。これにより、データを操作するためのより簡単なモデルが提供されます。

2番目の_callback_パラメーターとして関数に渡される匿名関数は、 `getSelectedDataAsync`操作が完了したときに実行されます。この関数は、1つのパラメーター _asyncResult_を指定して呼び出されます。これには、結果と呼び出しの状態が含まれます。呼び出しが失敗した場合[](/javascript/api/office/office.asyncresult#asynccontext) 、 `AsyncResult`オブジェクトの error プロパティは[error](/javascript/api/office/office.error)オブジェクトへのアクセスを提供します。[Error.name](/javascript/api/office/office.error#name)および[Error のメッセージ](/javascript/api/office/office.error#message)のプロパティの値を確認して、set 操作が失敗した理由を確認できます。それ以外の場合は、ドキュメント内で選択されているテキストが表示されます。

**If**ステートメントでは、呼び出しが成功したかどうかをテストするために、 [AsyncResult](/javascript/api/office/office.asyncresult#error)プロパティを使用します。[AsyncResultStatus](/javascript/api/office/office.asyncresult#status)は、使用可能な`AsyncResult.status`プロパティ値の列挙体です。`Office.AsyncResultStatus.Failed`文字列 "failed" に評価されます (つまり、そのリテラル文字列としても指定できます)。


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

_コールバック_パラメーターとして[Setselecteddataasync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)メソッドに渡される匿名関数は、非同期呼び出しが完了したときに実行されます。メソッドを使用して選択範囲にデータを書き込む場合、コールバックの_asyncResult_パラメーターは、呼び出しの状態だけでなく、呼び出しが失敗した場合は Error オブジェクトにアクセスします。 [](/javascript/api/office/office.error) `setSelectedDataAsync`

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

最初の_eventType_パラメーターは、サブスクライブするイベントの名前を指定します。このパラメーターの`"documentSelectionChanged"`文字列を渡すことは、イベントの`Office.EventType.DocumentSelectionChanged`種類として[EventType](/javascript/api/office/office.eventtype)列挙を渡すことと同じです。

2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](/javascript/api/office/office.documentselectionchangedeventargs) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。


> [!NOTE]
> 特定のイベントに対して複数のイベントハンドラーを追加する`addHandlerAsync`には、メソッドをもう一度呼び出し、 _handler_パラメーターに対して追加のイベントハンドラー関数を渡します。各イベントハンドラー関数の名前が一意である限り、これは正常に動作します。


## <a name="stop-detecting-changes-in-the-selection"></a>選択範囲の変更の検出を中止する


次の例は、[document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) メソッドを呼び出して、[Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) イベントのリッスンを中止する方法を示しています。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

2 `myHandler`番目の_handler_パラメーターとして渡される関数名は、 `SelectionChanged`イベントから削除されるイベントハンドラーを指定します。


> [!IMPORTANT]
> メソッドを呼び出すときにオプションの_handler_パラメーターを省略すると、指定された eventType のすべてのイベントハンドラーが削除されます。 __ `removeHandlerAsync`
