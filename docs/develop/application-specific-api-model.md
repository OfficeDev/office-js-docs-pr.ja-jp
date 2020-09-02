---
title: アプリケーション固有の API モデルの使用
description: Excel、OneNote、および Word のアドインの promise ベースの API モデルについて説明します。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: cabd1ea0076b672a1dbda3079a767b0e8a1a62b7
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326283"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="4fd14-103">アプリケーション固有の API モデルの使用</span><span class="sxs-lookup"><span data-stu-id="4fd14-103">Using the application-specific API model</span></span>

<span data-ttu-id="4fd14-104">この記事では、Excel、Word、および OneNote でアドインをビルドするための API モデルの使用方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="4fd14-105">Promise ベースの Api を使用するための基本的な概念について説明します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="4fd14-106">このモデルは、Office 2013 クライアントではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="4fd14-107">[共通 API モデル](office-javascript-api-object-model.md)を使用して、これらの Office バージョンを操作します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="4fd14-108">完全なプラットフォームの可用性に関する注意事項については、「office [クライアントアプリケーションおよび Office アドインのプラットフォームの可用性](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4fd14-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="4fd14-109">このページの例では、Excel JavaScript Api を使用していますが、概念は OneNote、Visio、および Word JavaScript Api にも適用されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="4fd14-110">Promise ベースの Api の非同期的な性質</span><span class="sxs-lookup"><span data-stu-id="4fd14-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="4fd14-111">Office アドインは、Excel などの Office アプリケーション内のブラウザーコンテナー内に表示される web サイトです。</span><span class="sxs-lookup"><span data-stu-id="4fd14-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="4fd14-112">このコンテナーは、office アプリケーション内のデスクトップベースのプラットフォーム (Windows 上の Office など) に組み込まれており、web 上の Office の HTML iFrame 内で実行されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="4fd14-113">パフォーマンスに関する考慮事項のため、Office.js Api は、すべてのプラットフォームで Office アプリケーションと同期して操作することはできません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="4fd14-114">したがって、 `sync()` Office.js の API 呼び出しは、Office アプリケーションが要求された読み取りまたは書き込みアクションを完了したときに解決される [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="4fd14-115">また、 `sync()` 各アクションに対して個別の要求を送信するのではなく、プロパティの設定やメソッドの呼び出しなど、複数のアクションをキューに入れて、1回の呼び出しで1つのコマンドのバッチとして実行することもできます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="4fd14-116">次のセクションでは、api を使用してこれを実現する方法について説明し `run()` `sync()` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="4fd14-117">\*. run 関数</span><span class="sxs-lookup"><span data-stu-id="4fd14-117">\*.run function</span></span>

<span data-ttu-id="4fd14-118">`Excel.run`、 `Word.run` 、 `OneNote.run` Excel、Word、および OneNote に対して実行するアクションを指定する関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="4fd14-119">`*.run` Office オブジェクトを操作するために使用できる要求コンテキストを自動的に作成します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="4fd14-120">完了すると `*.run` 、promise が解決され、実行時に割り当てられたすべてのオブジェクトが自動的に解放されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="4fd14-121">次の例は、を使用する方法を示して `Excel.run` います。</span><span class="sxs-lookup"><span data-stu-id="4fd14-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="4fd14-122">Word と OneNote でも同じパターンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-122">The same pattern is also used with Word and OneNote.</span></span>

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

## <a name="request-context"></a><span data-ttu-id="4fd14-123">要求コンテキスト</span><span class="sxs-lookup"><span data-stu-id="4fd14-123">Request context</span></span>

<span data-ttu-id="4fd14-124">Office アプリケーションとアドインは、2つの異なるプロセスで実行されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="4fd14-125">さまざまなランタイム環境を使用しているため、アドインを `RequestContext` Office のオブジェクト (ワークシート、範囲、段落、表など) に接続するために、アドインにはオブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="4fd14-126">この `RequestContext` オブジェクトは、を呼び出すときに引数として提供され `*.run` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="4fd14-127">プロキシ オブジェクト</span><span class="sxs-lookup"><span data-stu-id="4fd14-127">Proxy objects</span></span>

<span data-ttu-id="4fd14-128">Promise ベースの Api を宣言して使用する Office JavaScript オブジェクトは、プロキシオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="4fd14-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="4fd14-129">起動するメソッドや、プロキシ オブジェクトに設定または読み込まれるプロパティは、保留中のコマンドのキューに単純に追加されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="4fd14-130">`sync()`(など) 要求コンテキストでメソッドを呼び出すと `context.sync()` 、キューに入れられたコマンドが Office アプリケーションにディスパッチされて実行されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="4fd14-131">これらの Api は、基本的にバッチ中心です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="4fd14-132">要求コンテキストに対して必要な数だけ変更をキューに入れて、キューに `sync()` 入れられたコマンドのバッチを実行するメソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="4fd14-133">たとえば、次のコードスニペットでは、ローカルの JavaScript [excel. range](/javascript/api/excel/excel.range) オブジェクトを宣言して、 `selectedRange` excel ブック内の選択された範囲を参照し、そのオブジェクトにいくつかのプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="4fd14-134">`selectedRange`オブジェクトはプロキシオブジェクトなので、設定されているプロパティと、そのオブジェクトに対して呼び出されたメソッドは、アドインが呼び出されるまで Excel ドキュメントには反映されません `context.sync()` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="4fd14-135">パフォーマンスのヒント: 作成されるプロキシオブジェクトの数を最小限に抑える</span><span class="sxs-lookup"><span data-stu-id="4fd14-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="4fd14-136">同じプロキシ オブジェクトを繰り返し作成することは避けるようにします。</span><span class="sxs-lookup"><span data-stu-id="4fd14-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="4fd14-137">代わりに、複数の操作で同じプロキシ オブジェクトが必要な場合は、一度作成して変数に割り当ててから、その変数をコードで使用します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

### <a name="sync"></a><span data-ttu-id="4fd14-138">sync()</span><span class="sxs-lookup"><span data-stu-id="4fd14-138">sync()</span></span>

<span data-ttu-id="4fd14-139">`sync()`要求コンテキストに対してメソッドを呼び出すと、Office ドキュメント内のプロキシオブジェクトとオブジェクトの間で状態が同期されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="4fd14-140">この `sync()` メソッドは、要求コンテキストでキューに入れられた任意のコマンドを実行し、プロキシオブジェクトに読み込む必要があるすべてのプロパティの値を取得します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="4fd14-141">`sync()`メソッドは非同期的に実行され、 [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返します。これは、メソッドが完了したときに解決され `sync()` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="4fd14-142">次の例は、ローカルの JavaScript プロキシオブジェクト () を定義し `selectedRange` 、そのオブジェクトのプロパティを読み込んでから、JavaScript の約束パターンを使用して、 `context.sync()` Excel ドキュメント内のプロキシオブジェクトとオブジェクト間の状態を同期するバッチ関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="4fd14-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="4fd14-143">前の例では、`selectedRange` が設定されており、`context.sync()` が呼び出されると `address` プロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="4fd14-144">`sync()`は非同期操作なので、スクリプトを実行し `Promise` 続ける前に操作が完了したことを確認するために、常にオブジェクトを返してください `sync()` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="4fd14-145">TypeScript または ES6 + JavaScript を使用している場合は、promise を返す代わりに呼び出しを行うことができ `await` `context.sync()` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="4fd14-146">パフォーマンスに関するヒント: 同期呼び出しの数を最小限に抑える</span><span class="sxs-lookup"><span data-stu-id="4fd14-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="4fd14-147">Excel JavaScript API では、`sync()` は唯一の非同期操作で、状況によっては遅くなる可能性があり、Excel on the web の場合は特にその傾向があります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="4fd14-148">パフォーマンスを最適化するには、`sync()` を呼び出す前にできるだけ多くの変更をキューイングして、呼び出しの数を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="4fd14-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="4fd14-149">でパフォーマンスを最適化する方法の詳細について `sync()` は、「 [ループでのコンテキストの同期を回避する](../concepts/correlated-objects-pattern.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4fd14-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="4fd14-150">load()</span><span class="sxs-lookup"><span data-stu-id="4fd14-150">load()</span></span>

<span data-ttu-id="4fd14-151">プロキシオブジェクトのプロパティを読み取るには、その前にプロパティを明示的に読み込んで、Office ドキュメントからのデータをプロキシオブジェクトに設定してから、を呼び出して `context.sync()` ください。</span><span class="sxs-lookup"><span data-stu-id="4fd14-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="4fd14-152">たとえば、選択した範囲を参照するプロキシオブジェクトを作成してから、選択した範囲のプロパティを読み取る場合は、 `address` そのプロパティを読み込む前に読み込む必要があり `address` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="4fd14-153">プロキシオブジェクトのプロパティを読み込むには、 `load()` オブジェクトに対してメソッドを呼び出し、読み込むプロパティを指定します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="4fd14-154">次の例は、 `Range.address` 読み込むプロパティを示して `myRange` います。</span><span class="sxs-lookup"><span data-stu-id="4fd14-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

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
> <span data-ttu-id="4fd14-155">プロキシオブジェクトのメソッドを呼び出すか、プロパティを設定するだけの場合は、メソッドを呼び出す必要はありません `load()` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="4fd14-156">この `load()` メソッドは、プロキシオブジェクトのプロパティを読み取る場合にのみ必要です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="4fd14-p115">プロキシ オブジェクトに対してプロパティを設定、またはメソッドを呼び出す要求と同じように、プロキシ オブジェクトに対してプロパティを読み込む要求も、要求コンテキストで保留中のコマンドのキューに追加され、次回 `sync()` メソッドを呼び出すときに実行されます。`load()` の呼び出しは、必要なだけ要求コンテキストのキューに入れることができます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="4fd14-159">スカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="4fd14-159">Scalar and navigation properties</span></span>

<span data-ttu-id="4fd14-160">プロパティには、**スカラー**と**ナビゲーション**という 2 つのカテゴリがあります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="4fd14-161">スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="4fd14-162">ナビゲーションプロパティは、プロパティを直接代入するのではなく、読み取り専用のオブジェクトと、フィールドが割り当てられているオブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="4fd14-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="4fd14-163">たとえば、 `name` および `position` は、Excel の [Worksheet](/javascript/api/excel/excel.worksheet) オブジェクトのメンバーはスカラープロパティであり `protection` 、 `tables` ナビゲーションプロパティです。</span><span class="sxs-lookup"><span data-stu-id="4fd14-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="4fd14-164">アドインでは、ナビゲーションプロパティをパスとして使用して、特定のスカラープロパティを読み込むことができます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="4fd14-165">次のコードでは、 `load` オブジェクトで使用されているフォントの名前に対するコマンドをキューに入れ `Excel.Range` ます。その他の情報は読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="4fd14-166">また、パスを通過してナビゲーションプロパティのスカラープロパティを設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="4fd14-167">たとえば、を使用してのフォントサイズを設定でき `Excel.Range` `someRange.format.font.size = 10;` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="4fd14-168">プロパティを設定する前に、プロパティを読み込む必要はありません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="4fd14-169">オブジェクトのプロパティの中には、別のオブジェクトと同じ名前を持つものがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="4fd14-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="4fd14-170">たとえば、 `format` はオブジェクトの下にあるプロパティですが、 `Excel.Range` `format` それ自体もオブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="4fd14-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="4fd14-171">そのため、などの呼び出しを行うと `range.load("format")` 、これは `range.format.load()` (望ましくない empty ステートメント) と同じです `load()` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="4fd14-172">これを回避するには、コードでオブジェクトツリーの "葉 nodes" のみを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="4fd14-173">`load`パラメーターを使用せずに呼び出す (推奨されません)</span><span class="sxs-lookup"><span data-stu-id="4fd14-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="4fd14-174">`load()`パラメーターを指定せずにオブジェクト (またはコレクション) に対してメソッドを呼び出すと、オブジェクトまたはコレクションのオブジェクトのすべてのスカラープロパティが読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="4fd14-175">不要なデータを読み込むと、アドインの速度が低下します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="4fd14-176">読み込むプロパティを常に明示的に指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4fd14-177">パラメーターのない `load` ステートメントで返されるデータの量は、サービスのサイズ制限を超える場合があります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="4fd14-178">古いアドインのリスクを軽減するために、明示的に要求しない限り `load` によって返されないプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="4fd14-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="4fd14-179">次のプロパティは、そのような読み込み操作で除外されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="4fd14-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="4fd14-180">ClientResult</span></span>

<span data-ttu-id="4fd14-181">プリミティブ型を返す promise ベースの api のメソッドには、パラダイムに似たパターンがあり `load` / `sync` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="4fd14-182">たとえば、`Excel.TableCollection.getCount` はコレクション内のテーブルの数を取得します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="4fd14-183">`getCount` を返し `ClientResult<number>` ます。これは、返されたプロパティが数値であることを意味 `value` [`ClientResult`](/javascript/api/office/officeextension.clientresult) します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="4fd14-184">`context.sync()` が呼び出されるまで、スクリプトはその値にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="4fd14-185">次のコードは、Excel ブック内のテーブルの合計数を取得し、その番号をコンソールに記録します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

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

### <a name="set"></a><span data-ttu-id="4fd14-186">set()</span><span class="sxs-lookup"><span data-stu-id="4fd14-186">set()</span></span>

<span data-ttu-id="4fd14-187">入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="4fd14-188">前述のナビゲーションパスを使用して個々のプロパティを設定する代わりに、 `object.set()` promise ベースの JavaScript api のオブジェクトで使用可能なメソッドを使用することができます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="4fd14-189">このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="4fd14-p124">次のコード サンプルは、`set()` メソッドを呼び出し、`Range`Range\*\* オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲 \*\*B2:E2 にデータがあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

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

## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="4fd14-192">&#42;OrNullObject メソッドとプロパティ</span><span class="sxs-lookup"><span data-stu-id="4fd14-192">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="4fd14-193">必要なオブジェクトが存在しない場合、いくつかのアクセサーメソッドとプロパティは例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="4fd14-193">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="4fd14-194">たとえば、ブックにないワークシート名を指定して Excel ワークシートを取得しようとすると、 `getItem()` メソッドは例外をスロー `ItemNotFound` します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-194">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="4fd14-195">アプリケーション固有のライブラリは、コードが、例外処理コードを必要とせずにドキュメントエンティティの存在をテストするための方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-195">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="4fd14-196">これは、 `*OrNullObject` メソッドとプロパティのバリエーションを使用して実現されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-196">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="4fd14-197">これらのバリエーション `isNullObject` は、指定したアイテムが `true` 存在しない場合、例外をスローするのではなく、プロパティがに設定されているオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-197">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="4fd14-198">たとえば、 `getItemOrNullObject()` **ワークシート** などのコレクションに対してメソッドを呼び出して、コレクションからアイテムを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-198">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="4fd14-199">このメソッドは、指定されたアイテムが存在する場合はそれを返します。 `getItemOrNullObject()` それ以外の場合は、 `isNullObject` プロパティがに設定されているオブジェクトを返し `true` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-199">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="4fd14-200">コードでは、このプロパティを評価して、オブジェクトが存在するかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-200">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="4fd14-201">バリエーションは、 `*OrNullObject` JavaScript 値を返すことはありません `null` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-201">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="4fd14-202">通常の Office プロキシオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4fd14-202">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="4fd14-203">オブジェクトが表すエンティティが存在しない場合、 `isNullObject` オブジェクトのプロパティはに設定され `true` ます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-203">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="4fd14-204">全くまたは falsity に対して返されるオブジェクトはテストしません。</span><span class="sxs-lookup"><span data-stu-id="4fd14-204">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="4fd14-205">決して `null` 、 `false` 、、または `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="4fd14-205">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="4fd14-206">次のコードサンプルでは、メソッドを使用して、"Data" という名前の Excel ワークシートを取得しようとして `getItemOrNullObject()` います。</span><span class="sxs-lookup"><span data-stu-id="4fd14-206">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="4fd14-207">その名前のワークシートが存在しない場合は、新しいシートが作成されます。</span><span class="sxs-lookup"><span data-stu-id="4fd14-207">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="4fd14-208">コードによってプロパティが読み込まれないことに注意して `isNullObject` ください。</span><span class="sxs-lookup"><span data-stu-id="4fd14-208">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="4fd14-209">Office は、が呼び出されたときに、このプロパティを自動的に読み込み `context.sync` ます。したがって、などのように明示的に読み込む必要はありません `datasheet.load('isNullObject')` 。</span><span class="sxs-lookup"><span data-stu-id="4fd14-209">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4fd14-210">関連項目</span><span class="sxs-lookup"><span data-stu-id="4fd14-210">See also</span></span>

* [<span data-ttu-id="4fd14-211">共通 JavaScript API オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="4fd14-211">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* <span data-ttu-id="4fd14-212">[一般的なコーディングの問題と、予期しないプラットフォームの動作](common-coding-issues.md)。</span><span class="sxs-lookup"><span data-stu-id="4fd14-212">[Common coding issues and unexpected platform behaviors](common-coding-issues.md).</span></span>
* [<span data-ttu-id="4fd14-213">Office アドインのリソースの制限とパフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="4fd14-213">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
