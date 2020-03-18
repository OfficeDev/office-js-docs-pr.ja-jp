---
ms.date: 01/14/2020
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Normal
ms.openlocfilehash: 418c8124f8ed99b5ef1321c66f31ee0483da667b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719596"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="4fb65-103">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="4fb65-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="4fb65-104">カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。</span><span class="sxs-lookup"><span data-stu-id="4fb65-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="4fb65-105">[`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![API から時刻をストリームしているカスタム関数の GIF](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="4fb65-107">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="4fb65-107">Functions that return data from external sources</span></span>

<span data-ttu-id="4fb65-108">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="4fb65-109">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="4fb65-110">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="4fb65-111">Fetch の使用例</span><span class="sxs-lookup"><span data-stu-id="4fb65-111">Fetch example</span></span>

<span data-ttu-id="4fb65-112">次のコードサンプルでは、 `webRequest`関数は、"スペースがあります" という名前の "API" に到達します。これは、現在、国際宇宙ステーションにいるユーザーの数を追跡します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-112">In the following code sample, the `webRequest` function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="4fb65-113">この関数は JavaScript Promise を返し、fetchを使って API から情報を要求します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="4fb65-114">結果のデータは JSON に変換され、`names`プロパティは Promise を解決するために使用される文字列に変換されます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="4fb65-115">独自の機能を開発するときに、Web 要求が時間内に完了しない場合は、アクションを実行するか、[複数の API 要求をバッチ処理すること](./custom-functions-batching.md)を検討してください。</span><span class="sxs-lookup"><span data-stu-id="4fb65-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).</span></span>

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

>[!NOTE]
><span data-ttu-id="4fb65-116">`Fetch`を使用すると、コールバックのネストが回避され、場合によっては XHR に適している場合があります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="4fb65-117">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="4fb65-117">XHR example</span></span>

<span data-ttu-id="4fb65-118">カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-118">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="4fb65-119">単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="4fb65-119">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="4fb65-120">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-120">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="4fb65-121">コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-121">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

<span data-ttu-id="4fb65-122">次のコードサンプルでは、 `getStarCount`この関数は Github API を呼び出して、特定のユーザーのリポジトリに与えられた星の量を検出します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-122">In the following code sample, the `getStarCount` function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="4fb65-123">これは JavaScript Promise を返す非同期関数です。</span><span class="sxs-lookup"><span data-stu-id="4fb65-123">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="4fb65-124">データが web 呼び出しから取得されると、Promise が解決され、データがセルに返されます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-124">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="4fb65-125">ストリーミング関数を作成する</span><span class="sxs-lookup"><span data-stu-id="4fb65-125">Make a streaming function</span></span>

<span data-ttu-id="4fb65-126">ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新する必要ありません。</span><span class="sxs-lookup"><span data-stu-id="4fb65-126">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="4fb65-127">これは、[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)の関数のように、サービス オンラインのライブ データを確認する際に便利です。</span><span class="sxs-lookup"><span data-stu-id="4fb65-127">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="4fb65-128">ストリーミング関数を宣言するには、タグ`@streaming`を使用するか、`CustomFunctions.StreamingInvocation`呼び出しパラメーターを使用します。これは、関数がストリーミング中であることを示します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-128">To declare a streaming function, either use the tag `@streaming` or make use of the `CustomFunctions.StreamingInvocation` invocation parameter, which will indicate that your function is streaming.</span></span> <span data-ttu-id="4fb65-129">新しい情報に基づいて関数が再評価する可能性があることをユーザーに警告するには、関数の名前または説明にこれを示すことができるストリームまたはその他の文言を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="4fb65-129">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="4fb65-130">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="4fb65-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="4fb65-131">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="4fb65-131">Note the following about this code:</span></span>

- <span data-ttu-id="4fb65-132">Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-132">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="4fb65-133">2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="4fb65-133">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="4fb65-134">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-134">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="4fb65-135">ストリーミングは必ずしもWeb 要求の実行に結び付けられているわけではありません。この場合、関数は Web 要求を行うのではなく、設定された間隔でデータを取得しているため、ストリーミング `invocation` パラメータを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-135">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

```js
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
```

<span data-ttu-id="4fb65-136">`onCanceled`コールバックについて理解するだけでなく、Excel が次のような場合に関数の実行をキャンセルすることも理解しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-136">In addition to knowing about the `onCanceled` callback, you should also know that Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="4fb65-137">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="4fb65-137">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="4fb65-138">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="4fb65-138">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="4fb65-139">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-139">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="4fb65-140">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="4fb65-140">When the user triggers recalculation manually.</span></span> <span data-ttu-id="4fb65-141">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-141">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="4fb65-142">また、要求が発生したときに、オフラインの場合でも、ケースを処理する既定のストリーミング値を設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-142">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

> [!NOTE]
> <span data-ttu-id="4fb65-143">また、ストリーミング関数と関連の_ない_、キャンセル可能な関数と呼ばれる関数のカテゴリもあります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-143">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="4fb65-144">以前のバージョンのカスタム関数は、手動で記述された JSON で `"cancelable": true` と `"streaming": true` を宣言する必要がありました。</span><span class="sxs-lookup"><span data-stu-id="4fb65-144">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="4fb65-145">自動生成されたメタデータの導入以来、1 つの値を返す非同期のカスタム関数のみがキャンセル可能です。</span><span class="sxs-lookup"><span data-stu-id="4fb65-145">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="4fb65-146">キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-146">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="4fb65-147">タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-147">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="4fb65-148">起動パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="4fb65-148">Using an invocation parameter</span></span>

<span data-ttu-id="4fb65-149">`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="4fb65-149">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="4fb65-150">`invocation` パラメーターは、セルに関するコンテキスト (アドレスやコンテンツなど) を提供し、`setResult` メソッドや `onCanceled` メソッドを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-150">The `invocation` parameter gives context about the cell (such as its address and contents) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="4fb65-151">これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-151">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="4fb65-152">TypeScript を使用している場合は、呼び出しハンドラーは `CustomFunctions.StreamingInvocation` 型または `CustomFunctions.CancelableInvocation` 型である必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fb65-152">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

## <a name="receive-data-via-websockets"></a><span data-ttu-id="4fb65-153">WebSocket 経由のデータ受信</span><span class="sxs-lookup"><span data-stu-id="4fb65-153">Receive data via WebSockets</span></span>

<span data-ttu-id="4fb65-154">カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="4fb65-154">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="4fb65-155">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="4fb65-155">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="4fb65-156">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="4fb65-156">WebSockets example</span></span>

<span data-ttu-id="4fb65-157">以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="4fb65-157">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="4fb65-158">次の手順</span><span class="sxs-lookup"><span data-stu-id="4fb65-158">Next steps</span></span>

- <span data-ttu-id="4fb65-159">[関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。</span><span class="sxs-lookup"><span data-stu-id="4fb65-159">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="4fb65-160">[複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。</span><span class="sxs-lookup"><span data-stu-id="4fb65-160">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4fb65-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="4fb65-161">See also</span></span>

- [<span data-ttu-id="4fb65-162">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="4fb65-162">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="4fb65-163">カスタム関数の JSON メタデータを作成する</span><span class="sxs-lookup"><span data-stu-id="4fb65-163">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="4fb65-164">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="4fb65-164">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="4fb65-165">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="4fb65-165">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
- [<span data-ttu-id="4fb65-166">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="4fb65-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="4fb65-167">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="4fb65-167">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
