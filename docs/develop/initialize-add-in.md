---
title: Office アドインを初期化する
description: Office アドインを初期化する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 8310c5efb803391f7f0d4b01fda70dc0df537b21
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608140"
---
# <a name="initialize-your-office-add-in"></a><span data-ttu-id="534fd-103">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="534fd-103">Initialize your Office Add-in</span></span>

<span data-ttu-id="534fd-104">Office アドインには、次のような処理を行うスタートアップ ロジックがよくあります。</span><span class="sxs-lookup"><span data-stu-id="534fd-104">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="534fd-105">ユーザーのバージョンの Office で、コードが呼び出すすべての Office Api をサポートしていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="534fd-105">Check that the user's version of Office supports all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="534fd-106">特定の名前のワークシートなど、特定の成果物が存在することを確認します。</span><span class="sxs-lookup"><span data-stu-id="534fd-106">Ensure the existence of certain artifacts, such as a worksheet with a specific name.</span></span>

- <span data-ttu-id="534fd-107">Excel でセルを選択するようにユーザーに求め、選択した値で初期化されたグラフを挿入します。</span><span class="sxs-lookup"><span data-stu-id="534fd-107">Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.</span></span>

- <span data-ttu-id="534fd-108">バインディングを確立します。</span><span class="sxs-lookup"><span data-stu-id="534fd-108">Establish bindings.</span></span>

- <span data-ttu-id="534fd-109">Office ダイアログ API を使用して、既定のアドイン設定値をユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="534fd-109">Use the Office Dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="534fd-110">ただし、Office アドインは、ライブラリが読み込まれるまでは、Office JavaScript Api を正常に呼び出せません。</span><span class="sxs-lookup"><span data-stu-id="534fd-110">However, an Office Add-in cannot successfully call any Office JavaScript APIs until the library has been loaded.</span></span> <span data-ttu-id="534fd-111">この記事では、ライブラリが読み込まれていることをコードが確認する2つの方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="534fd-111">This article describes the two ways your code can ensure that the library has been loaded:</span></span>

- <span data-ttu-id="534fd-112">を使用して初期化 `Office.onReady()` します。</span><span class="sxs-lookup"><span data-stu-id="534fd-112">Initialize with `Office.onReady()`.</span></span>
- <span data-ttu-id="534fd-113">を使用して初期化 `Office.initialize` します。</span><span class="sxs-lookup"><span data-stu-id="534fd-113">Initialize with `Office.initialize`.</span></span>

> [!TIP]
> <span data-ttu-id="534fd-114">`Office.initialize` の代わりに `Office.onReady()` を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="534fd-114">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="534fd-115">`Office.initialize`はまだサポートされていますが、 `Office.onReady()` より柔軟な機能を提供します。</span><span class="sxs-lookup"><span data-stu-id="534fd-115">Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="534fd-116">割り当てることができるハンドラーは1つだけ `Office.initialize` で、Office のインフラストラクチャによって一度だけ呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="534fd-116">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure.</span></span> <span data-ttu-id="534fd-117">`Office.onReady()`コード内の別の場所で呼び出し、さまざまなコールバックを使用できます。</span><span class="sxs-lookup"><span data-stu-id="534fd-117">You can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="534fd-118">これらの手法の違いの詳細については、「[Office.initialize と Office.onReady の間の主な相違点](#major-differences-between-officeinitialize-and-officeonready)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="534fd-118">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="534fd-119">アドインの初期化時のイベントのシーケンスの詳細については、「[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="534fd-119">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

## <a name="initialize-with-officeonready"></a><span data-ttu-id="534fd-120">Office.onReady() を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="534fd-120">Initialize with Office.onReady()</span></span>

<span data-ttu-id="534fd-121">`Office.onReady()`は、Office .js ライブラリが読み込まれているかどうかを確認するときに、 [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)オブジェクトを返す非同期メソッドです。</span><span class="sxs-lookup"><span data-stu-id="534fd-121">`Office.onReady()` is an asynchronous method that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="534fd-122">ライブラリが読み込まれるとき (に限り)、Office ホスト アプリケーションを `Office.HostType` 列挙値 (`Excel`、`Word` など)、およびプラットフォームを `Office.PlatformType` 列挙値 (`PC`、`Mac`、`OfficeOnline` など) で指定するオブジェクトとして Promise を解決します。</span><span class="sxs-lookup"><span data-stu-id="534fd-122">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="534fd-123">`Office.onReady()` を呼び出すときにライブラリが既に読み込まれている場合、Promise をすぐに解決します。</span><span class="sxs-lookup"><span data-stu-id="534fd-123">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="534fd-124">`Office.onReady()` を呼び出す方法の 1 つは、コールバック メソッドを渡すことです。</span><span class="sxs-lookup"><span data-stu-id="534fd-124">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="534fd-125">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="534fd-125">Here's an example:</span></span>

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

<span data-ttu-id="534fd-126">また、コールバックを渡す代わりに、`then()` メソッドを `Office.onReady()` の呼び出しにチェーン接続することもできます。</span><span class="sxs-lookup"><span data-stu-id="534fd-126">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="534fd-127">たとえば、次のコードで、ユーザーのバージョンの Excel が、アドインで呼び出す可能性があるすべての API をサポートしているかを確認します。</span><span class="sxs-lookup"><span data-stu-id="534fd-127">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="534fd-128">`async` と `await` キーワードを TypeScript で使用する同じ例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="534fd-128">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="534fd-129">独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、*通常*、そのようなフレームワークは `Office.onReady()` への応答内に配置される必要があります。</span><span class="sxs-lookup"><span data-stu-id="534fd-129">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="534fd-130">たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="534fd-130">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="534fd-131">ただし、この実習には例外があります。</span><span class="sxs-lookup"><span data-stu-id="534fd-131">However, there are exceptions to this practice.</span></span> <span data-ttu-id="534fd-132">たとえば、ブラウザーのツールを使用してご使用の UI をデバッグするため、(Office ホスト内にサイドロードする代わりに) ブラウザーでご利用のアドインを開く必要があるとします。</span><span class="sxs-lookup"><span data-stu-id="534fd-132">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="534fd-133">Office.js がブラウザーに読み込まれないため、`onReady` は実行できず、Office `onReady` 内に呼び出される場合は、`$(document).ready` は実行されません。</span><span class="sxs-lookup"><span data-stu-id="534fd-133">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> 

<span data-ttu-id="534fd-134">アドインの読み込み中に作業ウィンドウに進行状況のインジケーターが表示されるようにする場合は、別の例外があります。</span><span class="sxs-lookup"><span data-stu-id="534fd-134">Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="534fd-135">このシナリオでは、コードで jQuery を呼び出し、コールバックを使用して進行状況インジケーターをレンダリングする必要があり `ready` ます。</span><span class="sxs-lookup"><span data-stu-id="534fd-135">In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator.</span></span> <span data-ttu-id="534fd-136">その後、Office `onReady` のコールバックで、進行状況のインジケーターを最終的な UI に置き換えることができます。</span><span class="sxs-lookup"><span data-stu-id="534fd-136">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

## <a name="initialize-with-officeinitialize"></a><span data-ttu-id="534fd-137">Office.initialize を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="534fd-137">Initialize with Office.initialize</span></span>

<span data-ttu-id="534fd-138">Office.js ライブラリが読み込まれ、ユーザーとの対話の準備が完了すると、初期化イベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="534fd-138">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="534fd-139">初期化ロジックを実装する `Office.initialize` にハンドラーを割り当てることができます。</span><span class="sxs-lookup"><span data-stu-id="534fd-139">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="534fd-140">ユーザーのバージョンの Excel が、アドインで呼び出す可能性があるすべての API をサポートしているかを確認する例は、次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="534fd-140">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="534fd-141">独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合は、*通常*、これらはイベント内に配置する必要があり `Office.initialize` ます (前の手順では、「 **Office. onready ()** セクションでの初期化」で説明されている例外)。</span><span class="sxs-lookup"><span data-stu-id="534fd-141">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also).</span></span> <span data-ttu-id="534fd-142">たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="534fd-142">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="534fd-143">作業ウィンドウ アドインとコンテンツ アドインの場合、`Office.initialize` で追加の _reason_ パラメーターが提供されます。</span><span class="sxs-lookup"><span data-stu-id="534fd-143">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="534fd-144">このパラメーターでは、アドインがどのように現在のドキュメントに追加されたかが示されます。</span><span class="sxs-lookup"><span data-stu-id="534fd-144">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="534fd-145">これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに、別のロジックを提供するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="534fd-145">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

<span data-ttu-id="534fd-146">詳細については、[Office.initialize イベント](/javascript/api/office)に関するページ、および [InitializationReason 列挙型](/javascript/api/office/office.initializationreason)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="534fd-146">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

## <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="534fd-147">Office.initialize と Office.onReady の間の主な相違点</span><span class="sxs-lookup"><span data-stu-id="534fd-147">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="534fd-148">`Office.initialize` にハンドラーは 1 つだけ割り当てることができ、1 回だけは、Office のインフラストラクチャで呼び出されますが、`Office.onReady()` の呼び出しはコードと異なる場所にして、異なるコールバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="534fd-148">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="534fd-149">たとえば、ご利用のコードでは、カスタム スクリプトが初期化ロジックを実行するコールバックを読み込むとすぐに `Office.onReady()` を呼び出しますが、ご利用のコードには、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出す、ボタンを作業ウィンドウに含めることもできます。</span><span class="sxs-lookup"><span data-stu-id="534fd-149">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="534fd-150">その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。</span><span class="sxs-lookup"><span data-stu-id="534fd-150">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="534fd-151">`Office.initialize` イベントは、Office.js 自体が初期化される内部プロセスの最後に発生します。</span><span class="sxs-lookup"><span data-stu-id="534fd-151">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="534fd-152">内部のプロセスが終了した後、*すぐに*発生します。</span><span class="sxs-lookup"><span data-stu-id="534fd-152">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="534fd-153">イベントにハンドラーを割り当てるコードが、イベント発生後に長時間実行される場合、ハンドラーは実行されません。</span><span class="sxs-lookup"><span data-stu-id="534fd-153">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="534fd-154">たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読み込まれた後で、カスタム JavaScript を読み込む前に、ポリフィルのファイルを読み込むためのアドインのホーム ページを構成する場合があります。</span><span class="sxs-lookup"><span data-stu-id="534fd-154">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="534fd-155">ご使用のスクリプトでハンドラーの読み込みと割り当てが行われる時点で、初期化イベントは既に発生しています。</span><span class="sxs-lookup"><span data-stu-id="534fd-155">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="534fd-156">ですが、`Office.onReady()` を呼び出すのに "遅すぎる" ことは決してありません。</span><span class="sxs-lookup"><span data-stu-id="534fd-156">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="534fd-157">初期化イベントが既に発生している場合、コールバックがすぐに実行されます。</span><span class="sxs-lookup"><span data-stu-id="534fd-157">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="534fd-158">スタートアップ ロジックがない場合でも、アドイン JavaScript を読み込むときには、`Office.onReady()` を呼び出すか、または空の関数を `Office.initialize` に割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="534fd-158">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="534fd-159">Office ホストとプラットフォームの組み合わせによっては、これらのいずれかが発生するまでは作業ウィンドウが読み込まれないことがあります。</span><span class="sxs-lookup"><span data-stu-id="534fd-159">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="534fd-160">次の例はこの 2 つの方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="534fd-160">The following examples show these two approaches.</span></span>
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a><span data-ttu-id="534fd-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="534fd-161">See also</span></span>

- [<span data-ttu-id="534fd-162">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="534fd-162">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="534fd-163">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="534fd-163">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)