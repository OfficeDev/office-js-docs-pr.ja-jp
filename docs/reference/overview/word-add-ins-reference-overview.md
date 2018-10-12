# <a name="word-javascript-api-overview"></a>Word の JavaScript API の使用状況の概要

Word には、ドキュメント コンテンツおよびメタデータとデータをやり取りするアドインを作成するために使用できる豊富な API のセットが用意されています。これらの API を使用して、Word を統合および拡張する魅力的なエクスペリエンスを作成します。コンテンツのインポートとエクスポート、別のデータ ソースから新しいドキュメントのアセンブル、カスタムのドキュメント ソリューションを作成するドキュメント ワークフローとの統合を行えます。

2 つの JavaScript API を使用して、Word 文書のオブジェクトおよびメタデータと対話できます。

- Word JavaScript API - Office 2016 で導入。
- [JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Office 2013 で導入。

## <a name="word-javascript-api"></a>Word JavaScript API

Word JavaScript API は Office.js によって読み込まれます。Word JavaScript API では、ドキュメントや段落などのオブジェクトとの対話方法が変わります。Word JavaScript API は、これらのそれぞれのオブジェクトの取得や更新をする個々の非同期の API を提供するのではなく、Word で実行されている実際のオブジェクトに対応する JavaScript の “プロキシ” オブジェクトを提供します。プロキシ オブジェクトのプロパティの読み取りと書き込みを同期的に行い、プロキシ オブジェクトに操作を実行する同期メソッドを呼び出すことによって、それらのプロキシ オブジェクトを操作することができます。プロキシ オブジェクトに対するこうした操作は実行中のスクリプトですぐには認識されません。**context.sync** メソッドは、キューに入れられた命令を実行し、また読み込まれた Word オブジェクトのプロパティをスクリプトで使用するために取得することで、実行中の JavaScript オブジェクトと Office の実際のオブジェクトとの間で状態を同期します。

## <a name="javascript-api-for-office"></a>JavaScript API for Office

Office.js は、次の場所から参照できます。

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js -生産のアドインの場合、このリソースを使用します。
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js -プレビュー機能を開こうとしているときにこのリソースを使用します。

[Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs) を使用している場合、[Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) をダウンロードして、Office.js を含むプロジェクト テンプレートを取得できます。[nuget から Office.js を取得する](https://www.nuget.org/packages/Microsoft.Office.js/)こともできます。

TypeScript を使用していて npm がある場合、コマンド ライン インターフェイスにこれを入力すると、TypeScript の定義を取得できます: `typings install office-js --ambient`。

## <a name="running-word-add-ins"></a>Word アドインを実行します

アドインを実行するには、Office.initialize イベント ハンドラーを使用します。アドインの初期化の詳細については、「[API について](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。

Word 2016以降をターゲットとするアドインは、関数を **Word.run()** メソッドに渡すことで実行されます。  **Run** メソッドに渡される関数は、コンテキストの引数が必要です。 この [コンテキスト オブジェクト](/javascript/api/word/word.requestcontext) は、 Office オブジェクトから取得するコンテキストオブジェクトとは異なりますが、 Word　のランタイム環境と対話するためにも使用されます。 コンテキスト オブジェクトでは、JavaScript API の Word オブジェクト モデルへのアクセスを提供します。 次の例では、 **Word.run()** メソッドを使用して、アドインを初期化し、Word を実行する方法を示します。

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Word 文書を Word JavaScript API のプロキシ オブジェクトと同期します

Word JavaScript API オブジェクト モデルは、Word 内のオブジェクトと緩く結合されています。Word JavaScript API のオブジェクトは、Word 文書内のオブジェクトのプロキシです。プロキシ オブジェクトで実行されたアクションは、ドキュメントの状態が同期されるまで、Word では認識されません。逆に、Word 文書の状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトでは認識されません。ドキュメントの状態を同期するには、**context.sync()** メソッドを実行します。次の例では、本文のプロキシ オブジェクトと、その本文プロキシ オブジェクトにテキスト プロパティを読み込むためのキューに登録済みのコマンドを作成し、さらに **context.sync()** メソッドを使用してWord 文書内の本文と本文プロキシ オブジェクトとを同期します。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>コマンドのバッチを実行する

Word のプロキシ オブジェクトには、オブジェクト モデルにアクセスして更新するためのメソッドが用意されています。これらのメソッドは、バッチでキューに入れられた順序で順番に実行されます。context.sync() 呼び出しが行われると、キューに入れられたすべてのコマンドが実行されます。

次の例では、コマンドのキューが機能する仕組みを示しています。**context.sync()** が呼び出されると、本文テキストを読み込むコマンドが Word で実行されます。次に、Word の本文にテキストを挿入するコマンドが生成されます。その結果は本文のプロキシ オブジェクトに返されます。Word JavaScript API の **body.text** プロパティの値は、テキストが Word 文書に挿入される<u>前</u>の Word 文書本文の値です。


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a>単語の JavaScript API 仕様を開く

新しい Word アドイン用の API の設計と開発にあたり、[Open API の仕様](../openspec.md) ページでこれらに対するフィードバックの提供が可能になります。Word JavaScript API 用のパイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="word-javascript-api-reference"></a>Word JavaScript API リファレンス

単語の JavaScript API の詳細については、 [Word の JavaScript API リファレンス ドキュメント](/javascript/api/word)を参照してください。

## <a name="see-also"></a>関連項目

* [Word アドインの概要](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [Office アドイン プラットフォームの概要](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [GitHub の Word アドインのサンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
