---
title: ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み
description: Word 文書またはスプレッドシートでアクティブな選択範囲にデータを読み書きする方法Excel説明します。
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 360701bc43a7fc63f8447ff9a068256d187e2a70
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187316"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み

[Document](/javascript/api/office/office.document) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。 これを行うために、オブジェクトは`Document`メソッドと`setSelectedDataAsync`メソッドを提供します`getSelectedDataAsync`。 このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。

このメソッドは `getSelectedDataAsync` 、ユーザーの現在の選択に対してのみ機能します。 実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) メソッドを使用 (または、[Bindings](/javascript/api/office/office.bindings) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。 ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。

## <a name="read-selected-data"></a>選択されたデータを読み取る

次の例は、ドキュメント内の選択範囲のデータを [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドで取得する方法を示しています。

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

この例では、最初のパラメーター _である coercionType_ を次のように `Office.CoercionType.Text` 指定します (リテラル文字列 `"text"`を使用してこのパラメーターを指定することもできます)。 この場合、コールバック関数の [asyncResult](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) パラメーターから取得できる [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。 別の型変換を指定すると、別の値が取得されます。 [Office.CoercionType](/javascript/api/office/office.coerciontype) は、使用できる型変換の値を表す列挙型です。 `Office.CoercionType.Text` は文字列 "text" に評価されます。

> [!TIP]
> **データ アクセスにマトリックスを使用する場合と、テーブルの coercionType を使用する場合。** 行と列が追加されたときに選択した表形式データを動的に拡張する必要があり、テーブル ヘッダーを操作する必要がある場合は、(メソッド`"table"`の _coercionType_ パラメーターを指定するか`Office.CoercionType.Table`) テーブル データ型を使用する`getSelectedDataAsync`必要があります。 データ構造内の行と列の追加は、テーブルとマトリックス データの両方でサポートされますが、行と列の追加はテーブル データでのみサポートされます。 行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用する必要があります (メソッドの coercionType パラメーターを指定するか、または`Office.CoercionType.Matrix`メソッド`"matrix"`の _coercionType_ パラメーター`getSelectedDataAsync`を指定します)。

2 番目のパラメーターである _コールバック_ として関数に渡される匿名関数は、操作が `getSelectedDataAsync` 完了すると実行されます。 この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。 呼び出しが失敗した場合、オブジェクトの `AsyncResult` [error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) プロパティは [Error](/javascript/api/office/office.error) オブジェクトへのアクセスを提供します。 [Error.name](/javascript/api/office/office.error#office-office-error-name-member) プロパティと [Error.message](/javascript/api/office/office.error#office-office-error-message-member) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。 呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。

The [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) property is used in the **if** statement to test whether the call succeeded. [Office。AsyncResultStatus](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) は、使用可能`AsyncResult.status`なプロパティ値の列挙体です。 `Office.AsyncResultStatus.Failed` は文字列 "failed" に評価されます (また、そのリテラル文字列として指定することもできます)。

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

_data_ パラメーターに異なるオブジェクト型を渡すと、結果が異なります。 結果は、ドキュメントで現在選択されている内容、アドインをホストしているクライアント アプリケーションOffice、および渡されたデータを現在の選択に強制できるかどうかによって異なります。

The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) method as the _callback_ parameter is executed when the asynchronous call is completed. メソッドを使用して `setSelectedDataAsync` 選択範囲にデータを書き込むと、コールバックの _asyncResult_ パラメーターは呼び出しの状態にのみアクセスし、呼び出しが失敗した場合 [は Error](/javascript/api/office/office.error) オブジェクトにアクセスできます。

> [!NOTE]
> Excel 2013 SP1 および Excel on the web の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../excel/excel-add-ins-tables.md)ができるようになりました。

## <a name="detect-changes-in-the-selection"></a>選択範囲の変更を検出する

次の例は、[Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) メソッドを使用して、[SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) イベントのイベント ハンドラーをドキュメント上に追加することで、選択範囲の変更を検出する方法を示しています。

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

最初のパラメーター _eventType_ は、サブスクライブするイベントの名前を指定します。 このパラメーターの文字列`"documentSelectionChanged"`を渡すことは、Officeの`Office.EventType.DocumentSelectionChanged`イベントの種類を渡すことと同じです[。EventType](/javascript/api/office/office.eventtype) 列挙体。

`myHandler()` 2 番目のパラメーターハンドラーとして関数に渡される関数 _は、ドキュメント_ で選択範囲が変更されたときに実行されるイベント ハンドラーです。 この関数は、非同期処理の完了時に、 _DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](/javascript/api/office/office.documentselectionchangedeventargs) という 1 つのパラメーターを使用して呼び出されます。 [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。

> [!NOTE]
> メソッドを再度呼び出し、ハンドラー パラメーターの追加のイベント ハンドラー関数を渡すことで、特定の `addHandlerAsync` イベントに対して複数のイベント _ハンドラー_ を追加できます。 この場合、各イベント ハンドラー関数の名前は一意である必要があります。

## <a name="stop-detecting-changes-in-the-selection"></a>選択範囲の変更の検出を中止する

次の例は、[document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) メソッドを呼び出して、[Document.SelectionChanged](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)) イベントのリッスンを中止する方法を示しています。

```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

2 番目のパラメーターハンドラーとして渡される関数名は`myHandler`_、イベント_ から削除されるイベント ハンドラーを`SelectionChanged`指定します。

> [!IMPORTANT]
> メソッドの呼び出し時に省略可能な  _ハンドラー_ パラメーターを `removeHandlerAsync` 省略すると、指定した _eventType のすべてのイベント_ ハンドラーが削除されます。
