---
title: Word JavaScript API を使用した基本的なプログラミングの概念
description: Word JavaScript API を使用して、Word 用アドインを構築します。
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 697f3068a039caa8ae60ed449bacb05f3999a1ee
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608566"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a>Word JavaScript API を使用した基本的なプログラミングの概念

この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用して Word 2016 以降のアドインを構築する場合の基本的な概念について説明します。

## <a name="referencing-officejs"></a>Office.js を参照する

Office.js は、次の場所から参照できます。

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 運用環境のアドインには、このリソースを使用します。

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - このリソースを使用してプレビュー機能を試します。

## <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。 Word JavaScript API 要件セットの詳細については、「[Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)」を参照してください。

## <a name="running-word-add-ins"></a>Word アドインを実行する

アドインを実行するには、`Office.initialize` イベント ハンドラーを使用します。 アドインの初期化の詳細については、「[API について](../develop/understanding-the-javascript-api-for-office.md)」を参照してください。

Word 2016 以降を対象とするアドインは、関数を `Word.run()` メソッドに渡すことによって実行されます。 `run` メソッドに渡される関数には、context 引数を含める必要があります。 この[コンテキスト オブジェクト](/javascript/api/word/word.requestcontext)は、Office オブジェクトから取得するコンテキスト オブジェクトとは異なりますが、これは Word ランタイム環境とやりとりするためにも使用されます。 コンテキスト オブジェクトを使用して、Word JavaScript API オブジェクト モデルにアクセスできます。 次の例では、`Word.run()` メソッドを使用することにより、Word アドインを初期化して実行する方法について示します。

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

### <a name="asynchronous-nature-of-word-apis"></a>Word API の非同期性

Word JavaScript API は Office.js で読み込まれます。 Word JavaScript API では、ドキュメントや段落などのオブジェクトとの対話方法が変わります。 Word JavaScript API は、これらの各オブジェクトを取得および更新するための個々の非同期 API を提供するのではなく、Word で実行されているライブ オブジェクトに対応する「プロキシ」JavaScript オブジェクトを提供します。 プロキシ オブジェクトのプロパティの読み取りと書き込みを同期的に行い、プロキシ オブジェクトに操作を実行する同期メソッドを呼び出すことによって、それらのプロキシ オブジェクトを操作することができます。 プロキシ オブジェクトに対するこうした操作は実行中のスクリプトですぐには認識されません。 `context.sync`context.sync メソッドは、キューに入れられた命令を実行し、また読み込まれた Word オブジェクトのプロパティをスクリプトで使用するために取得することで、実行中の JavaScript オブジェクトと Office の実際のオブジェクトとの間で状態を同期します。

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Word 文書を Word JavaScript API のプロキシ オブジェクトと同期する

Word JavaScript API オブジェクト モデルは、Word 内のオブジェクトと緩く結合されています。Word JavaScript API のオブジェクトは、Word 文書内のオブジェクトのプロキシです。プロキシ オブジェクトで実行されたアクションは、ドキュメントの状態が同期されるまで、Word では認識されません。逆に、Word 文書の状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトでは認識されません。ドキュメントの状態を同期するには、`context.sync()`context.sync()`context.sync()` メソッドを実行します。次の例では、本文のプロキシ オブジェクトと、その本文プロキシ オブジェクトにテキスト プロパティを読み込むためのキューに登録済みのコマンドを作成し、さらに context.sync() メソッドを使用してWord 文書内の本文と本文プロキシ オブジェクトとを同期します。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>コマンドのバッチを実行する

Word のプロキシ オブジェクトには、オブジェクト モデルにアクセスして更新するためのメソッドが用意されています。 これらのメソッドは、バッチでキューに入れられた順序で順番に実行されます。 `context.sync()` 呼び出しが行われると、バッチでキューに入れられたすべてのコマンドが実行されます。

次の例は、コマンド キューの仕組みを示します。 `context.sync()` 呼び出しが行われると、本文を読み込むコマンドが Word で実行されます。 次に、Word の本文にテキストを挿入するコマンドが生成されます。 その結果は本文のプロキシ オブジェクトに返されます。 Word JavaScript の `body.text`body.text<u> プロパティの値は、テキストが Word 文書に挿入される</u>前の Word 文書本文の値になります。

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

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

## <a name="see-also"></a>関連項目

- [Word JavaScript API の概要](../reference/overview/word-add-ins-reference-overview.md)
- [最初の Word アドインをビルドする](../quickstarts/word-quickstart.md)
- [Word アドインのチュートリアル](../tutorials/word-tutorial.md)
- [Word JavaScript API リファレンス](/javascript/api/word)