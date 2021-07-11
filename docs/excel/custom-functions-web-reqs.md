---
ms.date: 03/15/2021
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Normal
ms.openlocfilehash: 60f09b791b13d34a4a7f307bb9677c9fcc72ee97
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349601"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="a98c1-103">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="a98c1-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="a98c1-104">カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。</span><span class="sxs-lookup"><span data-stu-id="a98c1-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="a98c1-105">[`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![API から時刻をストリームするカスタム関数の GIF。](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="a98c1-107">外部ソースからデータを返す関数</span><span class="sxs-lookup"><span data-stu-id="a98c1-107">Functions that return data from external sources</span></span>

<span data-ttu-id="a98c1-108">カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a98c1-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="a98c1-109">JavaScript Promise を Excel に返します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="a98c1-110">コールバック関数を使用して Promise を最終値で解決します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="a98c1-111">Fetch の使用例</span><span class="sxs-lookup"><span data-stu-id="a98c1-111">Fetch example</span></span>

<span data-ttu-id="a98c1-112">次のコード サンプルでは、この関数は、国際宇宙ステーションの現在の人数を追跡する架空の Contoso "スペース内のユーザー数" API に `webRequest` 到達します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-112">In the following code sample, the `webRequest` function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="a98c1-113">この関数は JavaScript Promise を返し、fetchを使って API から情報を要求します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="a98c1-114">結果のデータは JSON に変換され、`names`プロパティは Promise を解決するために使用される文字列に変換されます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="a98c1-115">独自の機能を開発するときに、Web 要求が時間内に完了しない場合は、アクションを実行するか、[複数の API 要求をバッチ処理すること](custom-functions-batching.md)を検討してください。</span><span class="sxs-lookup"><span data-stu-id="a98c1-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](custom-functions-batching.md).</span></span>

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
><span data-ttu-id="a98c1-116">`Fetch`を使用すると、コールバックのネストが回避され、場合によっては XHR に適している場合があります。</span><span class="sxs-lookup"><span data-stu-id="a98c1-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="a98c1-117">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="a98c1-117">XHR example</span></span>

<span data-ttu-id="a98c1-118">次のコード サンプルでは、この関数は Github API を呼び出して、特定のユーザーのリポジトリに与えられた星の量 `getStarCount` を検出します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-118">In the following code sample, the `getStarCount` function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="a98c1-119">これは JavaScript Promise を返す非同期関数です。</span><span class="sxs-lookup"><span data-stu-id="a98c1-119">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="a98c1-120">データが web 呼び出しから取得されると、Promise が解決され、データがセルに返されます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-120">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

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

## <a name="make-a-streaming-function"></a><span data-ttu-id="a98c1-121">ストリーミング関数を作成する</span><span class="sxs-lookup"><span data-stu-id="a98c1-121">Make a streaming function</span></span>

<span data-ttu-id="a98c1-122">ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新する必要ありません。</span><span class="sxs-lookup"><span data-stu-id="a98c1-122">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="a98c1-123">これは、[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)の関数のように、サービス オンラインのライブ データを確認する際に便利です。</span><span class="sxs-lookup"><span data-stu-id="a98c1-123">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="a98c1-124">ストリーミング関数を宣言するには、次のいずれかを使用できます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-124">To declare a streaming function, you can use either:</span></span>

- <span data-ttu-id="a98c1-125">タグ `@streaming` 。</span><span class="sxs-lookup"><span data-stu-id="a98c1-125">The `@streaming` tag.</span></span>
- <span data-ttu-id="a98c1-126">呼 `CustomFunctions.StreamingInvocation` び出しパラメーター。</span><span class="sxs-lookup"><span data-stu-id="a98c1-126">The `CustomFunctions.StreamingInvocation` invocation parameter.</span></span>

<span data-ttu-id="a98c1-127">以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。</span><span class="sxs-lookup"><span data-stu-id="a98c1-127">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="a98c1-128">このコードについては以下の点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="a98c1-128">Note the following about this code.</span></span>

- <span data-ttu-id="a98c1-129">Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-129">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="a98c1-130">2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。</span><span class="sxs-lookup"><span data-stu-id="a98c1-130">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="a98c1-131">`onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-131">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="a98c1-132">ストリーミングは必ずしもWeb 要求の実行に結び付けられているわけではありません。この場合、関数は Web 要求を行うのではなく、設定された間隔でデータを取得しているため、ストリーミング `invocation` パラメータを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a98c1-132">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

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
```

## <a name="canceling-a-function"></a><span data-ttu-id="a98c1-133">関数をキャンセルする</span><span class="sxs-lookup"><span data-stu-id="a98c1-133">Canceling a function</span></span>

<span data-ttu-id="a98c1-134">Excel場合は、関数の実行を取り消します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-134">Excel cancels the execution of a function in the following situations.</span></span>

- <span data-ttu-id="a98c1-135">ユーザーが、関数を参照するセルを編集または削除した場合。</span><span class="sxs-lookup"><span data-stu-id="a98c1-135">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="a98c1-136">関数の引数 (入力) の 1 つが変更されたとき。</span><span class="sxs-lookup"><span data-stu-id="a98c1-136">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="a98c1-137">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-137">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="a98c1-138">ユーザーが手動で再計算をトリガーしたとき。</span><span class="sxs-lookup"><span data-stu-id="a98c1-138">When the user triggers recalculation manually.</span></span> <span data-ttu-id="a98c1-139">この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-139">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="a98c1-140">また、要求が発生したときに、オフラインの場合でも、ケースを処理する既定のストリーミング値を設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-140">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

<span data-ttu-id="a98c1-141">また、ストリーミング関数と関連の _ない_、キャンセル可能な関数と呼ばれる関数のカテゴリもあります。</span><span class="sxs-lookup"><span data-stu-id="a98c1-141">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="a98c1-142">1 つの値を返す非同期のカスタム関数だけが取り消し可能です。</span><span class="sxs-lookup"><span data-stu-id="a98c1-142">Only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="a98c1-143">キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-143">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="a98c1-144">タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-144">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="a98c1-145">起動パラメーターの使用</span><span class="sxs-lookup"><span data-stu-id="a98c1-145">Using an invocation parameter</span></span>

<span data-ttu-id="a98c1-146">`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="a98c1-146">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="a98c1-147">この `invocation` パラメーターは、セルに関するコンテキスト (アドレスやコンテンツなど) を提供し、使用およびメソッド `setResult` を `onCanceled` 使用できます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-147">The `invocation` parameter gives context about the cell (such as its address and contents) and allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="a98c1-148">これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-148">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="a98c1-149">TypeScript を使用している場合、呼び出しハンドラーは型または [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) .</span><span class="sxs-lookup"><span data-stu-id="a98c1-149">If you're using TypeScript, the invocation handler needs to be of type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) or [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).</span></span>

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="a98c1-150">WebSocket 経由のデータ受信</span><span class="sxs-lookup"><span data-stu-id="a98c1-150">Receiving data via WebSockets</span></span>

<span data-ttu-id="a98c1-151">カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="a98c1-151">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="a98c1-152">WebSockets を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベントが発生した場合にサーバーからメッセージを自動的に受信できます。データをサーバーに明示的にポーリングする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="a98c1-152">Using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="a98c1-153">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="a98c1-153">WebSockets example</span></span>

<span data-ttu-id="a98c1-154">以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="a98c1-154">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="a98c1-155">次の手順</span><span class="sxs-lookup"><span data-stu-id="a98c1-155">Next steps</span></span>

- <span data-ttu-id="a98c1-156">[関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。</span><span class="sxs-lookup"><span data-stu-id="a98c1-156">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="a98c1-157">[複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。</span><span class="sxs-lookup"><span data-stu-id="a98c1-157">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a98c1-158">関連項目</span><span class="sxs-lookup"><span data-stu-id="a98c1-158">See also</span></span>

- [<span data-ttu-id="a98c1-159">関数の揮発性の値</span><span class="sxs-lookup"><span data-stu-id="a98c1-159">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="a98c1-160">カスタム関数の JSON メタデータを作成する</span><span class="sxs-lookup"><span data-stu-id="a98c1-160">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="a98c1-161">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="a98c1-161">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
- [<span data-ttu-id="a98c1-162">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="a98c1-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="a98c1-163">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="a98c1-163">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
