---
ms.date: 06/21/2019
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Priority
ms.openlocfilehash: 39be2f0913e2eee4b1e5e7d5f704a47dee279cf5
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128256"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="6c7f7-103">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="6c7f7-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="6c7f7-104">カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="6c7f7-105">カスタム関数は XHR を通してデータを要求し、同時に要求を `fetch` したりデータをストリーミングしたりすることができます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-105">Custom functions can request data through XHR and `fetch` requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="6c7f7-106">次のドキュメンテーションはweb 要求のいくつかの例を説明していますが、ストリーミング機能を構築するには、[カスタム関数 チュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="6c7f7-107">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="6c7f7-107">Functions that return data from external sources</span></span>

<span data-ttu-id="6c7f7-108">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="6c7f7-109">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="6c7f7-110">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="6c7f7-111">[`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="6c7f7-112">カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="6c7f7-113">単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="6c7f7-114">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="6c7f7-115">コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="6c7f7-116">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="6c7f7-116">XHR example</span></span>

<span data-ttu-id="6c7f7-117">以下のコード サンプルでは、**getTemperature**関数が sendWebRequest 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="6c7f7-118">sendWebRequest 関数は XHR を使用して、データを提供するエンドポイントを要求する GET リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="6c7f7-119">コンテキストを使った XHR リクエストのその他のサンプルについては、[Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload)の Github リポジトリの、`getFile` 関数範囲内で[このファイル ](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) を参照ください。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="6c7f7-120">Fetch の使用例</span><span class="sxs-lookup"><span data-stu-id="6c7f7-120">Fetch example</span></span>

<span data-ttu-id="6c7f7-121">以下のコード サンプルでは、`stockPriceStream` 関数がストック ティッカー シンボルを使い、1000 ミリ秒ごとに株価を取得します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-121">In the following code sample, the `stockPriceStream` function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="6c7f7-122">このサンプルに関する詳細については、[カスタム関数チュートリアル](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-122">For more details about this sample, see the [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function).</span></span>

> [!NOTE]
> <span data-ttu-id="6c7f7-123">次のコードは、IEX Trading API を使用して株価情報を要求します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-123">The following code requests a stock quote using the IEX Trading API.</span></span> <span data-ttu-id="6c7f7-124">コードを実行する前に、API 要求で必要な API トークンを取得するため、[IEX Cloud を使用して無料アカウントを作成](https://iexcloud.io/)する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-124">Before you can run the code, you'll need to [create a free account with IEX Cloud](https://iexcloud.io/) so that you can get the API token that's required in the API request.</span></span>

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

        //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
        var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
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

## <a name="receive-data-via-websockets"></a><span data-ttu-id="6c7f7-125">WebSocket 経由のデータ受信</span><span class="sxs-lookup"><span data-stu-id="6c7f7-125">Receive data via WebSockets</span></span>

<span data-ttu-id="6c7f7-126">カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-126">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="6c7f7-127">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-127">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="6c7f7-128">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="6c7f7-128">WebSockets example</span></span>

<span data-ttu-id="6c7f7-129">以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-129">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="6c7f7-130">ストリーミング関数を作成する</span><span class="sxs-lookup"><span data-stu-id="6c7f7-130">Make a streaming function</span></span>

<span data-ttu-id="6c7f7-131">ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新する必要ありません。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-131">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="6c7f7-132">これは、[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)の関数のように、サービス オンラインのライブ データを確認する際に便利です。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-132">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="6c7f7-133">ストリーミング関数を宣言するには、JSDoc コメント タグ `@stream` を使用します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-133">To declare a streaming function, use the JSDoc comment tag `@stream`.</span></span> <span data-ttu-id="6c7f7-134">新しい情報に基づいて関数が再評価する可能性があることをユーザーに警告するには、関数の名前または説明にこれを示すことができるストリームまたはその他の文言を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-134">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="6c7f7-135">次の例では、指定した量だけ毎秒指定した数値を増加させるストリーミング関数を示しています。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-135">The following example shows a streaming function which increases a given number every second by an amount you specify.</span></span>

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> <span data-ttu-id="6c7f7-136">また、ストリーミング関数と関連の*ない*、キャンセル可能な関数と呼ばれる関数のカテゴリもあります。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-136">Note that there are also a category of functions called cancelable functions, which are *not* related to streaming functions.</span></span> <span data-ttu-id="6c7f7-137">以前のバージョンのカスタム関数は、手動で記述された JSON で `"cancelable": true` と `"streaming": true` を宣言する必要がありました。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-137">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="6c7f7-138">自動生成されたメタデータの導入以来、1 つの値を返す非同期のカスタム関数のみがキャンセル可能です。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-138">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="6c7f7-139">キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-139">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="6c7f7-140">タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-140">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="6c7f7-141">起動パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="6c7f7-141">Using an invocation parameter</span></span>

<span data-ttu-id="6c7f7-142">`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-142">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="6c7f7-143">`invocation` パラメーターは、セルに関するコンテキスト (アドレスなど) を提供し、`setResult` メソッドや `onCanceled` メソッドを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-143">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="6c7f7-144">これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-144">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="6c7f7-145">TypeScript を使用している場合は、呼び出しハンドラーは `CustomFunctions.StreamingInvocation` 型または `CustomFunctions.CancelableInvocation` 型である必要があります。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-145">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="6c7f7-146">ストリーム関数とキャンセル可能な関数の例</span><span class="sxs-lookup"><span data-stu-id="6c7f7-146">Streaming and cancelable function example</span></span>
<span data-ttu-id="6c7f7-147">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-147">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="6c7f7-148">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-148">Note the following about this code:</span></span>

- <span data-ttu-id="6c7f7-149">Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-149">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="6c7f7-150">2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-150">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="6c7f7-151">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-151">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

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
> <span data-ttu-id="6c7f7-152">Excel では、次のような状況で関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-152">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="6c7f7-153">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-153">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="6c7f7-154">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-154">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="6c7f7-155">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-155">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="6c7f7-156">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-156">When the user triggers recalculation manually.</span></span> <span data-ttu-id="6c7f7-157">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-157">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="6c7f7-158">次の手順</span><span class="sxs-lookup"><span data-stu-id="6c7f7-158">Next steps</span></span>

* <span data-ttu-id="6c7f7-159">[関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="6c7f7-160">[複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。</span><span class="sxs-lookup"><span data-stu-id="6c7f7-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6c7f7-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="6c7f7-161">See also</span></span>

* [<span data-ttu-id="6c7f7-162">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="6c7f7-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="6c7f7-163">カスタム関数の JSON メタデータを作成する</span><span class="sxs-lookup"><span data-stu-id="6c7f7-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="6c7f7-164">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="6c7f7-164">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="6c7f7-165">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="6c7f7-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="6c7f7-166">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="6c7f7-166">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="6c7f7-167">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="6c7f7-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="6c7f7-168">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="6c7f7-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
