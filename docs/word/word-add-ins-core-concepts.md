---
title: Word JavaScript API を使用した基本的なプログラミングの概念
description: Word JavaScript API を使用して、Word 用アドインを構築します。
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 7849780c1aed48152355c3fdbf350d798b2de1f2
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325018"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="44116-103">Word JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="44116-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="44116-104">この記事では、[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) を使用して Word 2016 以降のアドインを構築する場合の基本的な概念について説明します。</span><span class="sxs-lookup"><span data-stu-id="44116-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="44116-105">Office.js を参照する</span><span class="sxs-lookup"><span data-stu-id="44116-105">Referencing Office.js</span></span>

<span data-ttu-id="44116-106">Office.js は、次の場所から参照できます。</span><span class="sxs-lookup"><span data-stu-id="44116-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="44116-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - 運用環境のアドインには、このリソースを使用します。</span><span class="sxs-lookup"><span data-stu-id="44116-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="44116-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - このリソースを使用してプレビュー機能を試します。</span><span class="sxs-lookup"><span data-stu-id="44116-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="44116-109">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="44116-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="44116-110">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="44116-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="44116-111">Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="44116-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="44116-112">Word JavaScript API 要件セットの詳細については、「[Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44116-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="44116-113">Word アドインを実行する</span><span class="sxs-lookup"><span data-stu-id="44116-113">Running Word add-ins</span></span>

<span data-ttu-id="44116-114">アドインを実行するには、`Office.initialize` イベント ハンドラーを使用します。</span><span class="sxs-lookup"><span data-stu-id="44116-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="44116-115">アドインの初期化の詳細については、「[API について](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44116-115">For more information about add-in initialization, see [Understanding the API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

<span data-ttu-id="44116-116">Word 2016 以降を対象とするアドインは、関数を `Word.run()` メソッドに渡すことによって実行されます。</span><span class="sxs-lookup"><span data-stu-id="44116-116">Add-ins that target Word 2016 or later run by passing a function into the `Word.run()` method.</span></span> <span data-ttu-id="44116-117">`run` メソッドに渡される関数には、context 引数を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="44116-117">The function passed into the `run` method must have a context argument.</span></span> <span data-ttu-id="44116-118">この[コンテキスト オブジェクト](/javascript/api/word/word.requestcontext)は、Office オブジェクトから取得するコンテキスト オブジェクトとは異なりますが、これは Word ランタイム環境とやりとりするためにも使用されます。</span><span class="sxs-lookup"><span data-stu-id="44116-118">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="44116-119">コンテキスト オブジェクトを使用して、Word JavaScript API オブジェクト モデルにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="44116-119">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="44116-120">次の例では、`Word.run()` メソッドを使用することにより、Word アドインを初期化して実行する方法について示します。</span><span class="sxs-lookup"><span data-stu-id="44116-120">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

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

### <a name="asynchronous-nature-of-word-apis"></a><span data-ttu-id="44116-121">Word API の非同期性</span><span class="sxs-lookup"><span data-stu-id="44116-121">Asynchronous nature of Word APIs</span></span>

<span data-ttu-id="44116-122">Word JavaScript API は Office.js で読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="44116-122">The Word JavaScript API is loaded by Office.js.</span></span> <span data-ttu-id="44116-123">Word JavaScript API では、ドキュメントや段落などのオブジェクトとの対話方法が変わります。</span><span class="sxs-lookup"><span data-stu-id="44116-123">The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs.</span></span> <span data-ttu-id="44116-124">Word JavaScript API は、これらの各オブジェクトを取得および更新するための個々の非同期 API を提供するのではなく、Word で実行されているライブ オブジェクトに対応する「プロキシ」JavaScript オブジェクトを提供します。</span><span class="sxs-lookup"><span data-stu-id="44116-124">Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides "proxy" JavaScript objects that correspond to the live objects running in Word.</span></span> <span data-ttu-id="44116-125">プロキシ オブジェクトのプロパティの読み取りと書き込みを同期的に行い、プロキシ オブジェクトに操作を実行する同期メソッドを呼び出すことによって、それらのプロキシ オブジェクトを操作することができます。</span><span class="sxs-lookup"><span data-stu-id="44116-125">You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.</span></span> <span data-ttu-id="44116-126">プロキシ オブジェクトに対するこうした操作は実行中のスクリプトですぐには認識されません。</span><span class="sxs-lookup"><span data-stu-id="44116-126">These interactions with proxy objects aren't immediately realized in the running script.</span></span> <span data-ttu-id="44116-127">`context.sync`context.sync メソッドは、キューに入れられた命令を実行し、また読み込まれた Word オブジェクトのプロパティをスクリプトで使用するために取得することで、実行中の JavaScript オブジェクトと Office の実際のオブジェクトとの間で状態を同期します。</span><span class="sxs-lookup"><span data-stu-id="44116-127">The `context.sync` method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="44116-128">Word 文書を Word JavaScript API のプロキシ オブジェクトと同期する</span><span class="sxs-lookup"><span data-stu-id="44116-128">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="44116-p105">Word JavaScript API オブジェクト モデルは、Word 内のオブジェクトと緩く結合されています。Word JavaScript API のオブジェクトは、Word 文書内のオブジェクトのプロキシです。プロキシ オブジェクトで実行されたアクションは、ドキュメントの状態が同期されるまで、Word では認識されません。逆に、Word 文書の状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトでは認識されません。ドキュメントの状態を同期するには、`context.sync()`context.sync()`context.sync()` メソッドを実行します。次の例では、本文のプロキシ オブジェクトと、その本文プロキシ オブジェクトにテキスト プロパティを読み込むためのキューに登録済みのコマンドを作成し、さらに context.sync() メソッドを使用してWord 文書内の本文と本文プロキシ オブジェクトとを同期します。</span><span class="sxs-lookup"><span data-stu-id="44116-p105">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the `context.sync()` method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the `context.sync()` method to synchronize the body of the Word document with the body proxy object.</span></span>

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

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="44116-135">コマンドのバッチを実行する</span><span class="sxs-lookup"><span data-stu-id="44116-135">Executing a batch of commands</span></span>

<span data-ttu-id="44116-136">Word のプロキシ オブジェクトには、オブジェクト モデルにアクセスして更新するためのメソッドが用意されています。</span><span class="sxs-lookup"><span data-stu-id="44116-136">The Word proxy objects have methods for accessing and updating the object model.</span></span> <span data-ttu-id="44116-137">これらのメソッドは、バッチでキューに入れられた順序で順番に実行されます。</span><span class="sxs-lookup"><span data-stu-id="44116-137">These methods are run sequentially in the order in which they were queued in the batch.</span></span> <span data-ttu-id="44116-138">`context.sync()` 呼び出しが行われると、バッチでキューに入れられたすべてのコマンドが実行されます。</span><span class="sxs-lookup"><span data-stu-id="44116-138">All of the commands that are queued in the batch are run when `context.sync()` is called.</span></span>

<span data-ttu-id="44116-139">次の例は、コマンド キューの仕組みを示します。</span><span class="sxs-lookup"><span data-stu-id="44116-139">The following example shows how the command queue works.</span></span> <span data-ttu-id="44116-140">`context.sync()` 呼び出しが行われると、本文を読み込むコマンドが Word で実行されます。</span><span class="sxs-lookup"><span data-stu-id="44116-140">When `context.sync()` is called, the command to load the body text is run in Word.</span></span> <span data-ttu-id="44116-141">次に、Word の本文にテキストを挿入するコマンドが生成されます。</span><span class="sxs-lookup"><span data-stu-id="44116-141">Then, the command to insert text into the body in Word occurs.</span></span> <span data-ttu-id="44116-142">その結果は本文のプロキシ オブジェクトに返されます。</span><span class="sxs-lookup"><span data-stu-id="44116-142">The results are then returned to the body proxy object.</span></span> <span data-ttu-id="44116-143">Word JavaScript の `body.text`body.text<u> プロパティの値は、テキストが Word 文書に挿入される</u>前の Word 文書本文の値になります。</span><span class="sxs-lookup"><span data-stu-id="44116-143">The value of the `body.text` property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="44116-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="44116-144">See also</span></span>

- [<span data-ttu-id="44116-145">Word JavaScript API の概要</span><span class="sxs-lookup"><span data-stu-id="44116-145">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="44116-146">最初の Word アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="44116-146">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="44116-147">Word アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="44116-147">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="44116-148">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="44116-148">Word JavaScript API reference</span></span>](/javascript/api/word)