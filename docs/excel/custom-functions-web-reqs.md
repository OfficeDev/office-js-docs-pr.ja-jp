---
ms.date: 05/07/2019
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659636"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="23ac5-103">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="23ac5-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="23ac5-104">カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。</span><span class="sxs-lookup"><span data-stu-id="23ac5-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="23ac5-105">カスタム関数は XHR を通してデータを要求し、同時に要求を `fetch` したりデータをストリーミングしたりすることができます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="23ac5-106">次のドキュメンテーションはweb 要求のいくつかの例を説明していますが、ストリーミング機能を構築するには、[カスタム関数 チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="23ac5-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="23ac5-107">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="23ac5-107">Functions that return data from external sources</span></span>

<span data-ttu-id="23ac5-108">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="23ac5-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="23ac5-109">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="23ac5-110">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="23ac5-111">[`Fetch`](https://developer.mozilla.org/ja-JP/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ja-JP/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/ja-JP/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ja-JP/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="23ac5-112">カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/ja-JP/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/ja-JP/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="23ac5-113">単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="23ac5-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="23ac5-114">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="23ac5-115">コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="23ac5-116">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="23ac5-116">XHR example</span></span>

<span data-ttu-id="23ac5-117">以下のコード サンプルでは、**getTemperature**関数が sendWebRequest 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="23ac5-118">sendWebRequest 関数は XHR を使用して、データを提供するエンドポイントを要求する GET リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

<span data-ttu-id="23ac5-119">コンテキストを使った XHR リクエストのその他のサンプルについては、[Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload)の Github リポジトリの、`getFile` 関数範囲内で[このファイル ](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) を参照ください。</span><span class="sxs-lookup"><span data-stu-id="23ac5-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="23ac5-120">Fetch の使用例</span><span class="sxs-lookup"><span data-stu-id="23ac5-120">Fetch example</span></span>

<span data-ttu-id="23ac5-121">以下のコード サンプルでは、`stockPriceStream` 関数がストック ティッカー シンボルを使い、1000 ミリ秒ごとに株価を取得します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="23ac5-122">このサンプルに関する詳細については、[カスタム関数チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="23ac5-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a><span data-ttu-id="23ac5-123">WebSocket 経由のデータ受信</span><span class="sxs-lookup"><span data-stu-id="23ac5-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="23ac5-124">カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="23ac5-125">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="23ac5-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="23ac5-126">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="23ac5-126">WebSockets example</span></span>

<span data-ttu-id="23ac5-127">以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a><span data-ttu-id="23ac5-128">ストリーム関数とキャンセル関数</span><span class="sxs-lookup"><span data-stu-id="23ac5-128">Stream and cancel functions</span></span>

<span data-ttu-id="23ac5-129">ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新することは特に必要ありません。</span><span class="sxs-lookup"><span data-stu-id="23ac5-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span>

<span data-ttu-id="23ac5-130">キャンセル可能なカスタム関数を使用すると、帯域幅の消費量、作業メモリ、CPU への負荷を軽減するために、ストリーム カスタム関数の実行をキャンセルすることができます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-130">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span>

<span data-ttu-id="23ac5-131">関数をストリーミングまたはキャンセル可能として宣言するには、JSDOC コメント タグ `@stream` または `@cancelable` を使用します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-131">To declare a function as streaming or cancelable, use the JSDOC comment tags `@stream` or `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="23ac5-132">起動パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="23ac5-132">Using an invocation parameter</span></span>

<span data-ttu-id="23ac5-133">`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="23ac5-133">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="23ac5-134">`invocation` パラメーターは、セルに関するコンテキスト (アドレスなど) を提供し、`setResult` メソッドや `onCanceled` メソッドを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-134">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="23ac5-135">これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-135">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="23ac5-136">TypeScript を使用している場合は、呼び出しハンドラーは `CustomFunctions.StreamingInvocation` 型または `CustomFunctions.CancelableInvocation` 型である必要があります。</span><span class="sxs-lookup"><span data-stu-id="23ac5-136">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="23ac5-137">ストリーム関数とキャンセル可能な関数の例</span><span class="sxs-lookup"><span data-stu-id="23ac5-137">Streaming and cancelable function example</span></span>
<span data-ttu-id="23ac5-138">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="23ac5-138">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="23ac5-139">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="23ac5-139">Note the following about this code:</span></span>

- <span data-ttu-id="23ac5-140">Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-140">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="23ac5-141">2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="23ac5-141">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="23ac5-142">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="23ac5-142">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> <span data-ttu-id="23ac5-143">Excel では、次のような状況で関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="23ac5-143">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="23ac5-144">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="23ac5-144">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="23ac5-145">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="23ac5-145">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="23ac5-146">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-146">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="23ac5-147">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="23ac5-147">When the user triggers recalculation manually.</span></span> <span data-ttu-id="23ac5-148">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="23ac5-148">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="23ac5-149">次の手順</span><span class="sxs-lookup"><span data-stu-id="23ac5-149">Next steps</span></span>

* <span data-ttu-id="23ac5-150">[関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。</span><span class="sxs-lookup"><span data-stu-id="23ac5-150">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="23ac5-151">[複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。</span><span class="sxs-lookup"><span data-stu-id="23ac5-151">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="23ac5-152">関連項目</span><span class="sxs-lookup"><span data-stu-id="23ac5-152">See also</span></span>

* [<span data-ttu-id="23ac5-153">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="23ac5-153">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="23ac5-154">カスタム関数の JSON メタデータを作成する</span><span class="sxs-lookup"><span data-stu-id="23ac5-154">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="23ac5-155">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="23ac5-155">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="23ac5-156">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="23ac5-156">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="23ac5-157">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="23ac5-157">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="23ac5-158">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="23ac5-158">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="23ac5-159">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="23ac5-159">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
