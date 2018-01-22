
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>ドキュメントまたはスプレッドシート内の領域へのバインド

バインドベースのデータ アクセスにより、コンテンツ アドインおよび作業ウィンドウ アドインは、ドキュメントまたはスプレッドシートの特定の領域に ID を通じて一貫性をもってアクセスできます。アドインは、最初に、ドキュメントの部分と一意の ID を関連付けるいずれかのメソッド ([addFromPromptAsync]、[addFromSelectionAsync]、または [addFromNamedItemAsync]) を呼び出すことによって、バインドを確立する必要があります。バインドが確立されると、アドインは提供された ID を使用して、ドキュメントまたはスプレッドシート内の関連付けられた領域に含まれるデータにアクセスできます。バインドの作成により、アドインに次の効果がもたらされます。


- 表、範囲、またはテキスト (隣接する一連の文字) など、サポートされている Office アプリケーション全体に共通のデータ構造へのアクセスを許可します。
    
- ユーザーによる選択を必要とせずに、読み取り/書き込み操作ができます。
    
- アドインとドキュメント内のデータの間にリレーションシップが確立されます。バインドはドキュメント内に保持され、後でアクセスできます。
    
また、バインドを確立すると、ドキュメントまたはスプレッドシートの特定の領域を範囲とする、データおよび選択範囲の変更イベントをサブスクライブできます。つまり、ドキュメントまたはスプレッドシート全体の全般的な変更ではなく、バインドされた領域内で発生する変更のみがアドインに通知されます。

[Bindings] オブジェクトの [getAllAsync] メソッドを使用すると、ドキュメントまたはスプレッドシートに設定されているすべてのバインドのセットにアクセスできます。個々のバインドには、Bindings.[getByIdAsync] メソッドまたは [Office.select] メソッドを使用して ID でアクセスできます。[Bindings] オブジェクトの [addFromSelectionAsync]、[addFromPromptAsync]、[addFromNamedItemAsync]、または [releaseByIdAsync] メソッドのいずれかを使用して、新しいバインドを設定したり、既存のバインドを削除したりできます。


## <a name="binding-types"></a>バインドの種類

[addFromSelectionAsync] メソッド、[addFromPromptAsync] メソッド、または [addFromNamedItemAsync] メソッドでバインドを作成する場合、[3 種類のバインド][Office.BindingType] を _bindingType_ パラメーターで指定できます。

1. **[テキスト バインド][TextBinding]** - テキストとして表現できるドキュメントの領域にバインドします。

    Word では、連続する選択範囲の大部分が有効ですが、Excel では、単一セルの範囲のみがテキスト バインドの対象です。Excel では、プレーン テキストのみがサポートされます。Word では、3 つの形式 (プレーン テキスト、HTML、および Open XML for Office) がサポートされます。

2. **[マトリックス バインド][MatrixBinding]** - ヘッダーのない表形式データが含まれるドキュメントの固定領域にバインドします。マトリックス バインド内のデータは、2 次元の **Array** として読み書きされます。JavaScript では、これは配列の配列として実装されています。たとえば、2 列の **string** 値が 2 行ある場合は ` [['a', 'b'], ['c', 'd']]` のように書き込みまたは読み取りが行われ、1 列が 3 行ある場合は `[['a'], ['b'], ['c']]` のように行われます。

    Excel では、セルの連続する選択範囲を使用してマトリックス バインドを設定できます。Word では、表のみがマトリックス バインドをサポートします。

3. **[テーブル バインド][TableBinding]** - ヘッダーがある表が含まれるドキュメントの領域にバインドします。テーブル バインド内のデータは、[TableData](../../reference/shared/tabledata.md) オブジェクトとして書き込みまたは読み取りが行われます。`TableData` オブジェクトは `headers` および `rows` プロパティを通じてデータを公開します。

    Excel または Word の表はすべて、テーブル バインドの基礎にできます。テーブル バインドを確立すると、ユーザーが表に追加する新しい各行または各列が、自動的にバインドに含まれます。

`Bindings` オブジェクトの 3 つの "addFrom" メソッドのいずれかを使用してバインドを作成すると、[MatrixBinding]、[TableBinding]、または [TextBinding] のうち対応するオブジェクトのメソッドを使用して、バインドのデータとプロパティを操作できます。この 3 つのオブジェクトはすべて、`Binding` オブジェクトの [getDataAsync] メソッドおよび [setDataAsync] メソッドを継承しているので、バインドされたデータを操作できます。

> **マトリックス バインドとテーブル バインドの使い分け**作業中の表形式のデータに集計行が含まれ、アドインのスクリプトが集計行の値にアクセスする必要がある場合、またはユーザーの選択が集計行にあることを検出する必要がある場合は、マトリックス バインドを使用する必要があります。集計行を含む表形式データに対するテーブル バインドを設定する場合、[TableBinding.rowCount] プロパティおよびイベント ハンドラーの [BindingSelectionChangedEventArgs] オブジェクトの `rowCount` および `startRow` プロパティは、集計行のそれらの値に反映されません。この制限を回避するには、集計行を処理するマトリックス バインドを設定する必要があります。

## <a name="add-a-binding-to-the-users-current-selection"></a>ユーザーの現在の選択範囲にバインドを追加する

次の例は、[addFromSelectionAsync] メソッドを使用して、ドキュメントの現在の選択範囲に `myBinding` というテキスト バインドを追加する方法を示しています。


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例では、指定したバインドの種類はテキストです。つまり、選択範囲に対して [TextBinding] が作成されます。バインドが備えているデータと操作はバインドの種類ごとに異なります。[Office.BindingType] は、使用できるバインドの種類の値を示す列挙型です。

2 番目のオプションのパラメーターは、作成している新しいバインドの ID を指定するオブジェクトです。指定しない場合、ID は自動的に生成されます。

最後の _callback_ パラメーターで関数に渡される匿名関数は、バインドの作成が完了したときに実行されます。この関数の単一のパラメーター [ を通じて、呼び出しの状態を示す ]AsyncResult`asyncResult` オブジェクトにアクセスできます。`AsyncResult.value` プロパティには、新規作成するバインドとして指定した種類の [Binding] オブジェクトへの参照が格納されます。この [Binding] オブジェクトを使用して、データを取得および設定できます。

## <a name="add-a-binding-from-a-prompt"></a>プロンプトからバインドを追加する

次の例は、[addFromPromptAsync] メソッドを使用して `myBinding` という名前のテキスト バインドを追加する方法を示しています。このメソッドでは、ユーザーはアプリケーションの組み込み範囲選択プロンプトを使用してバインドの範囲を指定できます。


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例では、指定されているバインドの種類はテキストです。つまり、ユーザーがプロンプトで指定した選択範囲に対して [TextBinding] が作成されます。

2 番目のパラメーターは、作成している新しいバインドの ID を含むオブジェクトです。指定しない場合、ID は自動的に生成されます。

3 番目の _callback_ パラメーターとして関数に渡される匿名関数は、バインドの作成が完了すると実行されます。コールバック関数が実行されると、[AsyncResult] オブジェクトには呼び出しのステータスおよび新しく作成されたバインドが格納されます。

図 1 は、Excel の組み込み範囲選択プロンプトを示しています。


**図 1.Excel のデータ選択 UI**

![Excel のデータ選択 UI](../images/AgaveAPIOverview_ExcelSelectionUI.png)


## <a name="add-a-binding-to-a-named-item"></a>名前付きアイテムにバインドを追加する


次の例は、[addFromNamedItemAsync] メソッドを使用して、既存の `myRange` という名前のアイテムにマトリックス ("matrix") バインドを追加し、そのバインドの `id` に "myMatrix" を割り当てる方法を示しています。


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

 **Excel の場合**、[addFromNamedItemAsync] メソッドの `itemName` パラメーターは、既存の名前付き範囲 (`A1` スタイルの参照 `("A1:A3")` で指定された範囲) またはテーブルを参照できます。既定では、Excel のテーブルを追加すると、最初に追加したテーブルには "Table1"、次に追加したテーブルには "Table2" という名前が割り当てられます。Excel UI で意味のあるテーブル名を割り当てるには、リボンの **[テーブル ツール | デザイン]** タブの **[テーブル名]** プロパティを使用します。


 >**注**: Excel では、テーブルを名前付きアイテムとして指定する場合、`"Sheet1!Table1"` の形式で完全修飾名を指定して、テーブルの名前にワークシートの名前を含める必要があります。

以下の例では、Excel のバインドを列 A の最初の 3 つのセル (`"A1:A3"`) に対して作成し、id `"MyCities"` を割り当て、バインドに 3 つの都市名を書き込みます。


```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **Word の場合**、[addFromNamedItemAsync] メソッドの `itemName` パラメーターは、`Rich Text` コンテンツ コントロールの `Title` プロパティを参照します。(`Rich Text` コンテンツ コントロール以外のコンテンツ コントロールにはバインドできません)。

既定では、コンテンツ コントロールには `Title*` 値は割り当てられません。Word UI で意味のあるテーブル名を割り当てるには、リボンの **[開発]** タブの **[コントロール]** グループから **[リッチ テキスト]** コンテンツ コントロールを挿入した後、**[コントロール]** グループの **[プロパティ]** コマンドを使用して **[コンテンツ コントロールのプロパティ]** ダイアログ ボックスを表示します。次に、コンテンツ コントロールの **[タイトル]** プロパティに、コードから参照する名前を設定します。

次の例では、 `"FirstName"` という名前のリッチ テキスト コンテンツ コントロールに Word のテキスト バインドを作成し、 **id**`"firstName"` を割り当て、その情報を表示します。


```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a>すべてのバインドを取得する


次の例は、Bindings.[getAllAsync] メソッドを使用して、ドキュメント内のすべてのバインドを取得する方法を示しています。


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

`callback` パラメーターとして関数に渡される匿名関数は、操作の完了時に実行されます。この関数は、ドキュメント内のバインドの配列が格納される `asyncResult` という 1 つのパラメーターを使用して呼び出されます。配列は反復処理されて、バインドの ID を含む文字列が作成されます。この文字列がメッセージ ボックスに表示されます。


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>Bindings オブジェクトの getByIdAsync メソッドを使用して ID でバインドを取得する


次の例は、[getByIdAsync] メソッドを使用し、ID を指定してドキュメント内のバインドを取得する方法を示しています。この例では、前述のメソッドのいずれかを使用して `'myBinding'` という名前のバインドがドキュメントに追加されたと想定しています。


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例で、最初の `id` パラメーターは取得するバインドの ID です。

2 番目の  _callback_ パラメーターとして関数に渡される匿名関数は、操作の完了時に実行されます。この関数は、呼び出しのステータスおよび ID が "myBinding" であるバインドが格納される _asyncResult_ という 1 つのパラメーターを使用して呼び出されます。


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>Office オブジェクトの select メソッドを使用して ID でバインドを取得する


次の例は、[Office.select] メソッドを使用してセレクター文字列に ID を指定することによって、ドキュメント内の [Binding] オブジェクトの promise を取得する方法を示しています。その後、Binding.[getDataAsync] メソッドを呼び出して、指定したバインドからデータを取得します。この例では、前述のメソッドのいずれかを使用して `'myBinding'` という名前のバインドがドキュメントに追加されたと想定しています。


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


 > **注:**`select` メソッドの promise が正常に [Binding] オブジェクトを返す場合、このオブジェクトはオブジェクトの [getDataAsync]、[setDataAsync]、[addHandlerAsync]、および [removeHandlerAsync] の 4 つのメソッドのみを公開します。promise が Binding オブジェクトを返すことができない場合は、`onError` コールバックを使用して [asyncResult].error オブジェクトにアクセスし、詳細情報を取得できます。`select` メソッドによって返される Binding オブジェクトの promise によって公開される 4 つのメソッド以外の Binding オブジェクトのメンバーを呼び出す必要がある場合は、代わりに [getByIdAsync] メソッドを使用します。[Document.bindings] プロパティと Bindings.[getByIdAsync] メソッドを使用して Binding** オブジェクトを取得します。

## <a name="release-a-binding-by-id"></a>ID でバインドを解除する


次の例は、[releaseByIdAsync] メソッドを使用して ID を指定し、ドキュメント内のバインドを解除する方法を示しています。

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

この例で、最初の `id` パラメーターは解除するバインドの ID です。

2 番目のパラメーターとして関数に渡される匿名関数は、操作の完了時に実行されます。この関数は、呼び出しのステータスが格納される  [asyncResult] という 1 つのパラメーターを使用して呼び出されます。


## <a name="read-data-from-a-binding"></a>バインドからデータを読み取る


次の例は、[getDataAsync] メソッドを使用して既存のバインドからデータを取得する方法を示しています。


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。代わりに、[Office.select] メソッドを使用して ID によってバインドにアクセスし、次のように [getDataAsync] メソッドの呼び出しを開始できます。 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


関数に渡される匿名関数は、操作の完了時に実行されるコールバックです。[AsyncResult].value プロパティには、`myBinding` 内のデータが格納されます。その値の型は、バインドの種類により異なります。この例のバインドはテキスト バインドです。そのため、値には文字列が格納されます。マトリックス バインドおよびテーブル バインドを使用して作業する追加の例については、[getDataAsync] メソッドのトピックを参照してください。


## <a name="write-data-to-a-binding"></a>バインドにデータを書き込む

次の例は、[setDataAsync] メソッドを使用して既存のバインドにデータを設定する方法を示しています。

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。

この例で、最初のパラメーターは `myBinding` に設定する値です。これはテキスト バインドのため、値は `string` です。バインドの種類が異なる場合、異なる型のデータが使用されます。

関数に渡される匿名関数は、操作の完了時に実行されるコールバックです。この関数は、結果のステータスが格納される `asyncResult` という 1 つのパラメーターを使用して呼び出されます。

 > **メモ:** Excel 2013 SP1 および Excel Online の関連するビルドのリリースから、[バインド テーブルでデータの書き込みと更新を行う際に書式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)ができるようになりました。


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>バインド内のデータまたは選択範囲の変更を検出する


次の例は、ID が "MyBinding" であるバインドの [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) イベントにイベント ハンドラーを関連付ける方法を示しています。


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` は、ドキュメント内の既存のテキスト バインドを格納している変数です。

[addHandlerAsync] メソッドの最初の `eventType` パラメーターは、サブスクライブするイベントの名前を指定します。[Office.EventType] は、使用できるイベントの種類の値の列挙型です。`Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"`。

2 番目の  _handler_ パラメーターとして関数に渡される `dataChanged` 関数は、バインド内のデータが変更されたときに実行されるイベント ハンドラーです。この関数は、バインドへの参照が格納される _eventArgs_ という 1 つのパラメーターを使用して呼び出されます。このバインドを使用して、更新されたデータを取得できます。

同様に、バインドの [SelectionChanged] イベントにイベント ハンドラーを関連付けることによって、バインド内の選択範囲の変更を検出できます。これを行うには、[addHandlerAsync] メソッドの `eventType` パラメーターを `Office.EventType.BindingSelectionChanged` または `"bindingSelectionChanged"` と指定します。

[addHandlerAsync] メソッドを再び呼び出して、`handler` パラメーターに追加のイベント ハンドラー関数を指定すると、特定のイベントに複数のイベント ハンドラーを追加できます。この場合、各イベント ハンドラー関数の名前は一意である必要があります。


### <a name="remove-an-event-handler"></a>イベント ハンドラーを削除する


イベントのイベント ハンドラーを削除するには、最初の _eventType_ パラメーターにイベントの種類を指定し、2 番目の _handler_ パラメーターに削除するイベント ハンドラー関数の名前を指定して、[removeHandlerAsync] メソッドを呼び出します。たとえば、次の例では、前のセクションの例で追加した `dataChanged` イベント ハンドラー関数が削除されます。


```
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


 >**重要:**[removeHandlerAsync] メソッドを呼び出すときにオプションの _handler_ パラメーターを省略すると、指定された `eventType` のすべてのイベント ハンドラーが削除されます。


## <a name="additional-resources"></a>その他のリソース

- [JavaScript API for Office について](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office アドインにおける非同期プログラミング](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りと書き込みを行います](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Binding]:               ../../reference/shared/binding.md
[MatrixBinding]:         ../../reference/shared/binding.matrixbinding.md
[TableBinding]:          ../../reference/shared/binding.tablebinding.md
[TextBinding]:           ../../reference/shared/binding.textbinding.md
[getDataAsync]:          ../../reference/shared/binding.getdataasync.md
[setDataAsync]:          ../../reference/shared/binding.setdataasync.md
[SelectionChanged]:      ../../reference/shared/binding.bindingselectionchangedevent.md
[addHandlerAsync]:       ../../reference/shared/binding.addhandlerasync.md
[removeHandlerAsync]:    ../../reference/shared/binding.removehandlerasync.md

[Bindings]:              ../../reference/shared/bindings.bindings.md
[getByIdAsync]:          ../../reference/shared/bindings.getbyidasync.md 
[getAllAsync]:           ../../reference/shared/bindings.getallasync.md
[addFromNamedItemAsync]: ../../reference/shared/bindings.addfromnameditemasync.md
[addFromSelectionAsync]: ../../reference/shared/bindings.addfromselectionasync.md
[addFromPromptAsync]:    ../../reference/shared/bindings.addfrompromptasync.md
[releaseByIdAsync]:      ../../reference/shared/bindings.releasebyidasync.md

[AsyncResult]:          ../../reference/shared/asyncresult.md
[Office.BindingType]:   ../../reference/shared/bindingtype-enumeration.md
[Office.select]:        ../../reference/shared/office.select.md 
[Office.EventType]:     ../../reference/shared/eventtype-enumeration.md 
[Document.bindings]:    ../../reference/shared/document.bindings.md


[TableBinding.rowCount]: ../../reference/shared/binding.tablebinding.rowcount.1md
[BindingSelectionChangedEventArgs]: ../../reference/shared/binding.bindingselectionchangedeventargs.md
