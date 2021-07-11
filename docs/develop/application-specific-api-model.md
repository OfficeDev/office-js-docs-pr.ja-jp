---
title: アプリケーション固有の API モデルの使用
description: Excel、OneNote、および Word アドインの Promise ベースの API モデルについて説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5cf1d088dfa883e5df9eaba25e395857cfce9f5c
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350065"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="0225e-103">アプリケーション固有の API モデルの使用</span><span class="sxs-lookup"><span data-stu-id="0225e-103">Using the application-specific API model</span></span>

<span data-ttu-id="0225e-104">この記事では、Excel、Word、OneNote でアドインを構築するために API モデルを使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="0225e-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="0225e-105">この説明では、Promise ベースの API の使用に基本的な主要な概念を説明します。</span><span class="sxs-lookup"><span data-stu-id="0225e-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="0225e-106">このモデルは、Office 2013 クライアントではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0225e-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="0225e-107">これらの Office バージョンを使用しながら、[共通のAPIモデル](office-javascript-api-object-model.md) を使用します。</span><span class="sxs-lookup"><span data-stu-id="0225e-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="0225e-108">フル プラットフォーム可用性のノートについては、「[Office アドイン用 Office クライアント アプリケーションとプラットフォームの可用性](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0225e-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="0225e-109">このページの例では Excel JavaScript API を使用しますが、概念は OneNote、Visio、Word JavaScript API にも適用されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="0225e-110">Promise ベース API の非同期の性質</span><span class="sxs-lookup"><span data-stu-id="0225e-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="0225e-111">Office アドインは、Excel などの Office アプリケーション内のブラウザー コンテナー内に表示される Web サイトです。</span><span class="sxs-lookup"><span data-stu-id="0225e-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="0225e-112">コンテナーは、Office on Windows などのデスクトップ ベースのプラットフォーム上の Office アプリケーションに組み込まれ、Office on the Web の HTML iFrame 内で実行されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="0225e-113">パフォーマンスの考慮事項により、Office.js API は、すべてのプラットフォームの Office アプリケーションと同期して対話することはできません。</span><span class="sxs-lookup"><span data-stu-id="0225e-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="0225e-114">このため、`sync()`Office.js 内の API 呼び出しは Office アプリケーションが要求された読み取りまたは書き込み操作を完了したときに解決された[ Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返します。</span><span class="sxs-lookup"><span data-stu-id="0225e-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="0225e-115">また、操作ごとに別個の要求として送信する代わりに、プロパティの設定やメソッドの起動など、複数の操作をキューに登録し、`sync()`への 1 回の呼び出しでコマンドのバッチとしてそれらを実行することもできます。</span><span class="sxs-lookup"><span data-stu-id="0225e-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="0225e-116">次のセクションでは、`run()` および `sync()` API を使用してこれを実行する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="0225e-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="0225e-117">\*.run 関数</span><span class="sxs-lookup"><span data-stu-id="0225e-117">\*.run function</span></span>

<span data-ttu-id="0225e-118">`Excel.run`、 `Word.run`、 `OneNote.run`は、Excel、Word、OneNote に対して実行するアクションを指定する関数を実行できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="0225e-119">`*.run` は Office オブジェクトと対話するために使用できる要求コンテキストを自動的に作成します。</span><span class="sxs-lookup"><span data-stu-id="0225e-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="0225e-120">`*.run`が完了すると、Promose が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="0225e-121">次の例は、`Excel.run`の使用方法を説明しています。</span><span class="sxs-lookup"><span data-stu-id="0225e-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="0225e-122">Word と OneNote でも同じパターンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-122">The same pattern is also used with Word and OneNote.</span></span>

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a><span data-ttu-id="0225e-123">要求コンテキスト</span><span class="sxs-lookup"><span data-stu-id="0225e-123">Request context</span></span>

<span data-ttu-id="0225e-124">Office アプリケーションとユーザーのアドインは、2 つの異なるプロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="0225e-125">それらは異なるランタイム環境を使用するため、アドインは、ワークシート、範囲、グラフ、表など、Office のオブジェクトにユーザーのアドインを接続するために `RequestContext` オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="0225e-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="0225e-126">この `RequestContext` オブジェクトは、`*.run`を呼び出す際に引数として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="0225e-127">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0225e-127">Proxy objects</span></span>

<span data-ttu-id="0225e-128">Promise ベースの API と共にユーザーが宣言して使用する Office JavaScript オブジェクトはプロキシ オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0225e-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="0225e-129">起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="0225e-130">要求コンテキスト上 (たとえば `context.sync()`) で `sync()`メソッドを呼び出すと、キューに入れられたコマンドは Office アプリケーションにディスパッチされて実行されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="0225e-131">これらの API は、基本的にバッチ中心です。</span><span class="sxs-lookup"><span data-stu-id="0225e-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="0225e-132">要求コンテキストに必要なだけ変更内容をキューに登録し、`sync()` メソッドを呼び出して、キューに入れられたコマンドをバッチで実行することができます。</span><span class="sxs-lookup"><span data-stu-id="0225e-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="0225e-133">たとえば、次のコード スニペットでは、ローカル JavaScript [Excel.Range](/javascript/api/excel/excel.range) オブジェクト、`selectedRange`が Excel ワークブック内の選択範囲を参照することを宣言し、そのオブジェクトでいくつかのプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="0225e-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="0225e-134">`selectedRange` オブジェクトはプロキシ オブジェクトであるため、設定されたプロパティと、そのオブジェクトに対して呼び出されたメソッドは、ユーザーのアドインが `context.sync()` を呼び出すまで Excel ドキュメントには反映されません。</span><span class="sxs-lookup"><span data-stu-id="0225e-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="0225e-135">作業のヒント: 作成されたプロキシ オブジェクトの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="0225e-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="0225e-136">同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。</span><span class="sxs-lookup"><span data-stu-id="0225e-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="0225e-137">代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。</span><span class="sxs-lookup"><span data-stu-id="0225e-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a><span data-ttu-id="0225e-138">sync()</span><span class="sxs-lookup"><span data-stu-id="0225e-138">sync()</span></span>

<span data-ttu-id="0225e-139">要求コンテキストで `sync()`メソッドを呼び出すと、プロキシ オブジェクトと Officeドキュメント内のオブジェクトの状態が同期されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="0225e-140">`sync()` メソッドは、要求コンテキストのキューに登録されたすべてのコマンドを実行し、プロキシ オブジェクトに読み込まれるプロパティの値を取得します。  </span><span class="sxs-lookup"><span data-stu-id="0225e-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="0225e-141">`sync()`メソッドは非同期で実行されて [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。これは、`sync()` メソッドが完了すると解決されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="0225e-142">次の例は、ローカル JavaScript proxy オブジェクト (`selectedRange`) を定義し、そのオブジェクトのプロパティを読み込み、JavaScript の Promises パターンを使用して `context.sync()` を呼び出し、プロキシ オブジェクトと Excel ドキュメント内のオブジェクトの状態を同期するバッチ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="0225e-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="0225e-143">前の例では、`selectedRange` が設定されており、`context.sync()` が呼び出されると `address` プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="0225e-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="0225e-144">`sync()`が非同期操作である場合、スクリプトが引き続き実行される前に、`Promise` オブジェクトを返して、`sync()`の操作が完了するのを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="0225e-145">TypeScript または ES6+ JavaScript を使用している場合は、Promise を返す代わりに `context.sync()` の呼び出しを`await` にできます。</span><span class="sxs-lookup"><span data-stu-id="0225e-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="0225e-146">作業のこつ: 同期呼び出しの数を最小限にする</span><span class="sxs-lookup"><span data-stu-id="0225e-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="0225e-147">Excel JavaScript API では、`sync()` は唯一の非同期操作で、状況によっては遅くなる可能性があり、Excel on the web の場合は特にその傾向があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="0225e-148">パフォーマンスを最適化するには、`sync()` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="0225e-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="0225e-149">パフォーマンスを`sync()`で最適化する方法の詳細については、「[ループで context.sync メソッドの使用を避ける](../concepts/correlated-objects-pattern.md)」をご参照ください。</span><span class="sxs-lookup"><span data-stu-id="0225e-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="0225e-150">load()</span><span class="sxs-lookup"><span data-stu-id="0225e-150">load()</span></span>

<span data-ttu-id="0225e-151">プロキシ オブジェクトのプロパティを読み取るには、まず Office ドキュメントからプロキシ オブジェクトとデータを入力するためにプロパティを明確に読み込み、`context.sync()`を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="0225e-152">たとえば、選択範囲を操作するプロキシ オブジェクトを作成してから選択範囲の`address` プロパティを読み取る場合、読み取る前に`address` プロパティを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="0225e-153">読み込むプロキシ オブジェクトのプロパティを要求するには、オブジェクトに対して `load()` メソッドを呼び出し、読み込むプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="0225e-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="0225e-154">次の例は、`myRange`に読み込まれているプロパティ `Range.address`を示しています 。</span><span class="sxs-lookup"><span data-stu-id="0225e-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> <span data-ttu-id="0225e-155">プロキシ オブジェクト上でメソッドを呼び出す、またはプロパティを設定するだけの場合は、`load()` メソッドを呼び出す必要はありません。  </span><span class="sxs-lookup"><span data-stu-id="0225e-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="0225e-156">`load()` メソッドは、プロキシ オブジェクト上でプロパティを読み取る場合のみ必要です。</span><span class="sxs-lookup"><span data-stu-id="0225e-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="0225e-p115">プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 `sync()` メソッドを呼び出すときに実行されます。`load()` の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。</span><span class="sxs-lookup"><span data-stu-id="0225e-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="0225e-159">スカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="0225e-159">Scalar and navigation properties</span></span>

<span data-ttu-id="0225e-160">プロパティには、**スカラー** と **ナビゲーション** という 2 つのカテゴリがあります。</span><span class="sxs-lookup"><span data-stu-id="0225e-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="0225e-161">スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。</span><span class="sxs-lookup"><span data-stu-id="0225e-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="0225e-162">ナビゲーション プロパティは、プロパティを直接割り当てるのではなく、読み取り専用のオブジェクトと、そのフィールドが割り当てられているオブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="0225e-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="0225e-163">たとえば、[Excel.Worksheet](/javascript/api/excel/excel.worksheet)のオブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はナビゲーション プロパティです。</span><span class="sxs-lookup"><span data-stu-id="0225e-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="0225e-164">アドインは、特定のスカラー プロパティを読み込むパスとしてナビゲーション プロパティを使用できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="0225e-165">次のコードは、ほかの情報を読み込む必要なく、`Excel.Range` オブジェクトで使用されるフォント名の`load` コマンドをキューに入れられます。</span><span class="sxs-lookup"><span data-stu-id="0225e-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="0225e-166">パスを詳しく調べることでナビゲーション プロパティのスカラー プロパティを設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="0225e-167">たとえば、`someRange.format.font.size = 10;`を使用して`Excel.Range` のフォント サイズを設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="0225e-168">設定前にプロパティを読み込む必要はありません。</span><span class="sxs-lookup"><span data-stu-id="0225e-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="0225e-169">オブジェクトの下の「プロパティ」の中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="0225e-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="0225e-170">例えば、`format` は`Excel.Range`オブジェクトの下のプロパティですが、`format` それ自体もオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0225e-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="0225e-171">そのため、`range.load("format")`などの呼び出しを行った場合、これは `range.format.load()` (望ましくない空の空白のステートメント`load()`) と同等になります。</span><span class="sxs-lookup"><span data-stu-id="0225e-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="0225e-172">これを避けるには、コードがオブジェクト ツリー内の "リーフノード" のみをロードするようにしてください。</span><span class="sxs-lookup"><span data-stu-id="0225e-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="0225e-173">パラメーターを使用せず (非推奨) に `load` を呼び出す</span><span class="sxs-lookup"><span data-stu-id="0225e-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="0225e-174">パラメーターを指定せずにオブジェクト (またはコレクション) の `load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティ (またはコレクション内のオブジェクト) が読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="0225e-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="0225e-175">不要なデータを読み込むと、アドインの速度が低下します。</span><span class="sxs-lookup"><span data-stu-id="0225e-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="0225e-176">常に読み込むプロパティを明示的に指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0225e-177">パラメーターのない `load` ステートメントで返されるデータの量は、サービスのサイズ制限を超える場合があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="0225e-178">古いアドインのリスクを軽減するために、明示的に要求しない限り `load` によって返されないプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="0225e-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="0225e-179">次のプロパティは、このような読み込み操作から除外されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-179">The following properties are excluded from such load operations.</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="0225e-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="0225e-180">ClientResult</span></span>

<span data-ttu-id="0225e-181">プリミティブ型を返す、Promise ベースの API 内のメソッドは、`load`/`sync`パラダイムと同様のパターンを持っています。</span><span class="sxs-lookup"><span data-stu-id="0225e-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="0225e-182">たとえば、`Excel.TableCollection.getCount` はコレクション内のテーブルの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="0225e-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="0225e-183">`getCount` は `ClientResult<number>` を返します。つまり、返される[`ClientResult`](/javascript/api/office/officeextension.clientresult)の`value`プロパティは数値になります。</span><span class="sxs-lookup"><span data-stu-id="0225e-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="0225e-184">`context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="0225e-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="0225e-185">次のコードは、Excel ワークブック内のテーブルの総数を取得し、その数をコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="0225e-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a><span data-ttu-id="0225e-186">set()</span><span class="sxs-lookup"><span data-stu-id="0225e-186">set()</span></span>

<span data-ttu-id="0225e-187">入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。</span><span class="sxs-lookup"><span data-stu-id="0225e-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="0225e-188">前述のナビゲーション パスを使用してプロパティを個別に設定する代わりに、Promise ベースの JavaScript API のオブジェクトで使用できる、`object.set()`メソッドを使用できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="0225e-189">このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="0225e-p124">次のコード サンプルは、`set()` メソッドを呼び出し、`Range`Range **オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲** B2:E2 にデータがあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="0225e-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="0225e-192">一部のプロパティを直接設定できません</span><span class="sxs-lookup"><span data-stu-id="0225e-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="0225e-193">書き込み可能であるにもかかわらず、一部のプロパティを設定できません。</span><span class="sxs-lookup"><span data-stu-id="0225e-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="0225e-194">これらのプロパティは、1 つのオブジェクトとして設定する必要がある親プロパティの一部です。</span><span class="sxs-lookup"><span data-stu-id="0225e-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="0225e-195">これは、親プロパティが特定の論理関係を持つサブプロパティに依存しているからです。</span><span class="sxs-lookup"><span data-stu-id="0225e-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="0225e-196">これらの親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクト全体を設定するためにオブジェクト リテラル表記を使用して設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="0225e-197">その 1 つの例は、[PageLayout](/javascript/api/excel/excel.pagelayout)にあります。</span><span class="sxs-lookup"><span data-stu-id="0225e-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="0225e-198">`zoom`プロパティは、1 つの [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクト を使用して、以下のように設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="0225e-199">前の例では、`zoom` 値: `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることは ***できません***。</span><span class="sxs-lookup"><span data-stu-id="0225e-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="0225e-200">このステートメントは、`zoom` が読み込まれないので、エラーを発生させます。</span><span class="sxs-lookup"><span data-stu-id="0225e-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="0225e-201">`zoom` が読み込まれるような場合でも、スケール セットは有効化されません。</span><span class="sxs-lookup"><span data-stu-id="0225e-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="0225e-202">すべてのコンテキスト操作は `zoom`上、でアドインのプロキシオブジェクトを更新し、ローカルに設定された値を上書きする場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="0225e-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="0225e-203">この動作は、[Range.format](/javascript/api/excel/excel.range#format)など、[ナビゲーション プロパティ](application-specific-api-model.md#scalar-and-navigation-properties) とは異なります。</span><span class="sxs-lookup"><span data-stu-id="0225e-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="0225e-204">ここに示されているように、`format`のプロパティはオブジェクト ナビゲーションを使用して設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="0225e-205">読み取り専用の修飾キーを確認することで、サブプロパティを直接設定できないプロパティを識別できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="0225e-206">読み取り専用プロパティはすべて、読み取り専用以外のサブプロパティを直接設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="0225e-207">`PageLayout.zoom` のような書き可能なプロパティは、そのレベルのオブジェクトで設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0225e-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="0225e-208">まとめると、以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="0225e-208">In summary:</span></span>

- <span data-ttu-id="0225e-209">読み取り専用プロパティ: ナビゲーション経由でサブプロパティを設定できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="0225e-210">書き込み可能なプロパティ: サブプロパティをナビゲーションを介して設定することはできません (最初の親オブジェクトの一部として設定する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="0225e-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="0225e-211">&#42;OrNullObject メソッドとプロパティ</span><span class="sxs-lookup"><span data-stu-id="0225e-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="0225e-212">一部のアクセサリ方法とプロパティでは、目的のオブジェクトが存在しない場合に例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="0225e-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="0225e-213">たとえば、ブックに存在しないワークシート名を指定して Excel ワークシートを取得しようとすると、`getItem()` メソッドは `ItemNotFound` 例外を返します。</span><span class="sxs-lookup"><span data-stu-id="0225e-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="0225e-214">アプリケーション固有のライブラリを使用すると、例外処理コードを必要とせずに、コードがドキュメント エンティティの存在をテストできます。</span><span class="sxs-lookup"><span data-stu-id="0225e-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="0225e-215">これは、`*OrNullObject`メソッドのバリエーションとプロパティ を使用して行います。</span><span class="sxs-lookup"><span data-stu-id="0225e-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="0225e-216">これらのバリエーションは、 `isNullObject` プロパティが `true`に設定されているオブジェクトを返します (指定したアイテムが存在しない場合は、例外をスローしません)。</span><span class="sxs-lookup"><span data-stu-id="0225e-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="0225e-217">たとえば、**Worksheets** などのコレクションで `getItemOrNullObject()` メソッドを呼び出して、コレクションからのアイテムの取得を試行できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="0225e-218">`getItemOrNullObject()` メソッドは、指定された項目が存在する場合はその項目を返し、それ以外の場合は `isNullObject`プロパティが `true`に設定されているオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0225e-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="0225e-219">コードは、このプロパティを評価して、オブジェクトが存在するかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="0225e-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="0225e-220">`*OrNullObject` のバリエーションは、JavaScript 値`null`を返すことはありません。</span><span class="sxs-lookup"><span data-stu-id="0225e-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="0225e-221">通常の Office プロキシ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0225e-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="0225e-222">オブジェクトが表すエンティティが存在しない場合は、オブジェクトの `isNullObject` プロパティが `true`に設定されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="0225e-223">返されたオブジェクトの null 値または 真偽性はテストしません。</span><span class="sxs-lookup"><span data-stu-id="0225e-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="0225e-224">これは、決して `null`、 `false`、`undefined`ではありません。</span><span class="sxs-lookup"><span data-stu-id="0225e-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="0225e-225">次のコード サンプルは `getItemOrNullObject()` メソッドを使用して、"Data" という名前のワークシートの取得を試行します。</span><span class="sxs-lookup"><span data-stu-id="0225e-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="0225e-226">その名前のワークシートが存在しない場合は、新しいシートが作成されます。</span><span class="sxs-lookup"><span data-stu-id="0225e-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="0225e-227">コードは`isNullObject`プロパティを読み込まないことにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="0225e-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="0225e-228">Office は、`context.sync`が呼ばれると、自動的にこのプロパティを読み込みます。ですから、`datasheet.load('isNullObject')`のような名前で明示的に読み込む必要はありません。</span><span class="sxs-lookup"><span data-stu-id="0225e-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a><span data-ttu-id="0225e-229">関連項目</span><span class="sxs-lookup"><span data-stu-id="0225e-229">See also</span></span>

* [<span data-ttu-id="0225e-230">共通 JavaScript API オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="0225e-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="0225e-231">Office アドインのリソースの制限とパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="0225e-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
