---
ms.date: 03/21/2019
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: Web 要求とその他のデータがカスタム関数(プレビュー)を処理します
localization_priority: Priority
ms.openlocfilehash: 9256e2aa87ec6d7b314314a1e4bc2b3793f1df5c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926668"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="d82f1-103">カスタム関数によるデータの受信と処理</span><span class="sxs-lookup"><span data-stu-id="d82f1-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="d82f1-104">カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) など workbook以外からのデータの受信です。</span><span class="sxs-lookup"><span data-stu-id="d82f1-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="d82f1-105">カスタム関数はXHRを通してデータを要求し、同時に要求を fetch したりデータをストリーミングする事ができます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="d82f1-106">次のドキュメンテーションはweb 要求のいくつかの例を説明していますが、ストリーミング機能を構築するには、[カスタム関数 チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="d82f1-107">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="d82f1-107">Functions that return data from external sources</span></span>

<span data-ttu-id="d82f1-108">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d82f1-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="d82f1-109">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="d82f1-110">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="d82f1-111">[`Fetch`](https://developer.mozilla.org/ja-JP/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ja-JP/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-111">Within a custom function, you can request external data by using an API like Fetch or by using XmlHttpRequest (XHR), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="d82f1-112">カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/ja-JP/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-112">Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/ja-JP/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="d82f1-113">単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="d82f1-114">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="d82f1-115">コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="d82f1-116">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="d82f1-116">XHR example</span></span>

<span data-ttu-id="d82f1-117">以下のコード サンプルでは、**getTemperature**関数が sendWebRequest 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-117">In the following code sample, the  function calls the  function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="d82f1-118">sendWebRequest 関数は XHR を使用して、データを提供するエンドポイントを要求する GET リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-118">The  function uses XHR to issue a  request to an endpoint that can provide the data.</span></span>

```JavaScript
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
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

<span data-ttu-id="d82f1-119">コンテキストを使った XHR リクエストのその他のサンプルについては、[Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload)の Github リポジトリの、`getFile` 関数範囲内で[このファイル ](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) を参照ください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="d82f1-120">Fetch の使用例</span><span class="sxs-lookup"><span data-stu-id="d82f1-120">Fetch example</span></span>

<span data-ttu-id="d82f1-121">以下のコードサンプルでは、stockPriceStream 関数が ストック ティッカー シンボル を使い、1000 ミリ秒ごとに株価を取得します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="d82f1-122">このサンプルに関する詳細および JSON については、[カスタム関数 チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)を参照ください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="d82f1-123">WebSocket 経由のデータ受信</span><span class="sxs-lookup"><span data-stu-id="d82f1-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="d82f1-124">カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="d82f1-125">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="d82f1-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="d82f1-126">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="d82f1-126">WebSockets example</span></span>

<span data-ttu-id="d82f1-127">以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-127">The following code sample establishes a  connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="d82f1-128">ストリーミング関数</span><span class="sxs-lookup"><span data-stu-id="d82f1-128">Streaming functions</span></span>

<span data-ttu-id="d82f1-129">ストリーム カスタム関数を使用すると、セルに繰り返しデータを長期的に出力でき、ユーザーが再計算を明示的に要求することは特に必要ありません。</span><span class="sxs-lookup"><span data-stu-id="d82f1-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="d82f1-130">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="d82f1-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="d82f1-131">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-131">Note the following about this code:</span></span>

- <span data-ttu-id="d82f1-132">Excel は、setResult コールバックを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-132">Excel displays each new value automatically using the  callback.</span></span>
- <span data-ttu-id="d82f1-133">2 番目の入力パラメーター、ハンドラーは、オートコンプリート メニューから変数を選択する場合には Excel のエンドユーザーには表示されません。</span><span class="sxs-lookup"><span data-stu-id="d82f1-133">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="d82f1-134">onCanceled コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-134">The  callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="d82f1-135">すべてのストリーム関数には、このようなキャンセル ハンドラーの実装が必要です。</span><span class="sxs-lookup"><span data-stu-id="d82f1-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="d82f1-136">詳細については、「[関数をキャンセルする](#canceling-a-function)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d82f1-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

<span data-ttu-id="d82f1-137">JSON メタデータ ファイルでストリーミング関数のメタデータを指定する場合は、オプション オブジェクト内のプロパティ "cancelable": true と "stream": true を以下の例のように設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d82f1-137">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties  and  within the  object, as shown in the following example.</span></span>

```JSON
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="d82f1-138">関数をキャンセルする</span><span class="sxs-lookup"><span data-stu-id="d82f1-138">Canceling a function</span></span>

<span data-ttu-id="d82f1-139">状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を軽減するために、ストリーム カスタム関数の実行をキャンセルする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d82f1-139">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="d82f1-140">Excel では、次のような状況で関数の実行をキャンセルします。</span><span class="sxs-lookup"><span data-stu-id="d82f1-140">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="d82f1-141">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="d82f1-141">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="d82f1-142">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="d82f1-142">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="d82f1-143">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-143">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="d82f1-144">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="d82f1-144">When the user triggers recalculation manually.</span></span> <span data-ttu-id="d82f1-145">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="d82f1-145">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="d82f1-146">関数をキャンセル可能にするには、関数コードのハンドラーを実装し、キャンセルされたときの対応を指示します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-146">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="d82f1-147">さらに、関数を表す JavaScript Object Notation メタデータのオプション オブジェクト内のプロパティ`"cancelable": true` を指定します。</span><span class="sxs-lookup"><span data-stu-id="d82f1-147">Additionally, specify specify the property `"cancelable": true` within the options object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="d82f1-148">この記事の前のセクションのコード サンプルで、これらの手法の例が示されています。</span><span class="sxs-lookup"><span data-stu-id="d82f1-148">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="see-also"></a><span data-ttu-id="d82f1-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="d82f1-149">See also</span></span>

* [<span data-ttu-id="d82f1-150">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="d82f1-150">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="d82f1-151">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="d82f1-151">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d82f1-152">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="d82f1-152">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="d82f1-153">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="d82f1-153">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d82f1-154">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="d82f1-154">Custom functions changelog</span></span>](custom-functions-changelog.md)
