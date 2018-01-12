
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>ドキュメントまたはスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み

[Document](../../reference/shared/document.md) オブジェクトが公開しているメソッドを使用すると、ユーザーのドキュメントまたはスプレッドシート内の現在の選択範囲への読み取りと書き込みを行うことができます。これは、**Document** オブジェクトの **getSelectedDataAsync** メソッドと **setSelectedDataAsync** メソッドで行います。このトピックでは、ユーザーの選択範囲の読み取り方法、書き込み方法、およびその変更を検出するイベント ハンドラーの作成方法についても説明します。


  **getSelectedDataAsync** メソッドは、現在のユーザーの選択範囲のみに実行されます。実行中のアドインのセッション間で読み取りおよび書き取りに同じ選択範囲を利用できるように、ドキュメントの選択範囲を保持する必要がある場合、[Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) メソッドを使用 (または、[Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) オブジェクトの他の "addFrom" メソッドの 1 つでバインドを作成) して、バインドを追加する必要があります。ドキュメントの領域にバインドを作成して、バインドの読み取りおよび書き込みを行う詳細については、「[ドキュメントまたはスプレッドシート内の領域へのバインド](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)」を参照してください。


## <a name="read-selected-data"></a>選択されたデータを読み取る


次の例は、ドキュメント内の選択範囲のデータを [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) メソッドで取得する方法を示しています。


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

この例では、最初の _coercionType_ パラメーターには **Office.CoercionType.Text** が指定されています (このパラメーターはリテラル文字列 `"text"` で指定することもできます)。この場合、コールバック関数の [asyncResult](../../reference/shared/asyncresult.status.md) パラメーターから取得できる [AsyncResult](../../reference/shared/asyncresult.md) オブジェクトの _value_ プロパティは、ドキュメント内で選択されているテキストを格納している **string** を返します。別の型変換を指定すると、別の値が取得されます。[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) は、使用できる型変換の値を表す列挙型です。**Office.CoercionType.Text** は文字列 "text" として評価されます。


 >**ヒント:** **どのようなタイミングでデータ アクセスにマトリックスを使用し、どのような場合にテーブルの coercionType を使用するか。**行と列が追加されたときに選択済みの表形式データが動的に増えるようにし、またテーブル ヘッダーを使用する必要がある場合は、テーブル データ型を使用します (**getSelectedDataAsync** メソッドの _coercionType_ パラメーターに`"table"` または **Office.CoercionType.Table** を指定)。データ構造体内での行と列の追加はテーブル データとマトリックス データの両方でサポートされていますが、行と列の追加はテーブル データでのみサポートされています。行と列を追加する予定がなく、データにヘッダー機能が必要ない場合は、マトリックス データ型を使用します (**getSelecteDataAsync** メソッドの _coercionType_ パラメーターに `"matrix"` または **Office.CoercionType.Matrix** を指定)。このデータ型では、データとのやり取りについて、より単純なモデルを採用しています。

2 番目の _callback_ パラメーターとして関数に渡される匿名関数は、**getSelectedDataAsync** 操作の完了時に実行されます。この関数は、結果および呼び出しのステータスが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。呼び出しが失敗した場合は、[AsyncResult](../../reference/shared/asyncresult.context.md) オブジェクトの **error** プロパティから [Error](../../reference/shared/error.md) オブジェクトにアクセスできます。[Error.name](../../reference/shared/error.name.md) プロパティと [Error.message](../../reference/shared/error.message.md) プロパティの値をチェックして、設定の操作が失敗した理由を判断できます。呼び出しが成功した場合は、ドキュメント内で選択されているテキストが表示されます。

[if](../../reference/shared/asyncresult.error.md) ステートメントでは、呼び出しが成功したかどうかの判定に **AsyncResult.status** プロパティを使用します。[Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) は **AsyncResult.status** プロパティが取ることのできる値を表す列挙型です。**Office.AsyncResultStatus.Failed** は文字列 "failed" として評価されます (こちらもリテラル文字列で指定することもできます)。


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

[callback](../../reference/shared/document.setselecteddataasync.md) パラメーターとして _setSelectedDataAsync_ メソッドに渡される匿名関数は、非同期呼び出しの完了時に実行されます。**setSelectedDataAsync** メソッドを使用して選択範囲にデータを書き込む場合、コールバックの _asyncResult_ パラメーターは呼び出しのステータスへのアクセスのみを提供し、呼び出しが失敗した場合は[Error](../../reference/shared/error.md) オブジェクトにアクセスします。

 **注意:** Excel 2013 SP1 および Excel Online の関連するビルドのリリースから、[現在の選択範囲にテーブルを書き込む際に書式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)ができるようになりました。


## <a name="detect-changes-in-the-selection"></a>選択範囲の変更を検出する


次の例は、[Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) メソッドを使用して、[SelectionChanged](../../reference/shared/document.selectionchanged.event.md) イベントのイベント ハンドラーをドキュメント上に追加することで、選択範囲の変更を検出する方法を示しています。


```
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

最初の  _eventType_ パラメーターでは、サブスクライブするイベントの名前を指定しています。文字列 `"documentSelectionChanged"` をこのパラメーターに指定するのは、 **Office.EventType** 列挙型のイベントの種類 [Office.EventType.DocumentSelectionChanged](../../reference/shared/eventtype-enumeration.md) を渡すことに相当します。

2 番目の _handler_ パラメーターとして関数に渡される `myHander()` 関数は、ドキュメントで選択範囲が変更されたときに実行されるイベント ハンドラーです。この関数は、非同期処理の完了時に、_DocumentSelectionChangedEventArgs_ オブジェクトへの参照が格納される [eventArgs](../../reference/shared/document.selectionchangedeventargs.md) という 1 つのパラメーターを使用して呼び出されます。[DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) プロパティを使用すると、このイベントが発生したドキュメントにアクセスできます。


 >**メモ**   **addHandlerAsync** メソッドを再び呼び出して、 _handler_ パラメーターに追加のイベント ハンドラー関数を指定すると、特定のイベントに複数のイベント ハンドラーを追加できます。この場合、各イベント ハンドラー関数の名前は一意である必要があります。


## <a name="stop-detecting-changes-in-the-selection"></a>選択範囲の変更の検出を中止する


次の例は、[document.removeHandlerAsync](../../reference/shared/document.selectionchanged.event.md) メソッドを呼び出して、[Document.SelectionChanged](../../reference/shared/document.removehandlerasync.md) イベントのリッスンを中止する方法を示しています。


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

2 番目の  _handler_ パラメーターとして渡される `myHandler` 関数名は、 **SelectionChanged** イベントから削除されるイベント ハンドラーを指定します。


 >**重要:****removeHandlerAsync** メソッドを呼び出すときにオプションの _handler_ パラメーターを省略すると、指定された _eventType_ のすべてのイベント ハンドラーが削除されます。

