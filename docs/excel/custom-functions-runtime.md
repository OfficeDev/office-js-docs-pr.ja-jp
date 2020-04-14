---
ms.date: 04/13/2020
description: 新しい JavaScript ランタイムを使用する Excel カスタム関数を開発する場合の重要なシナリオについて、理解します。
title: Excel カスタム関数のランタイム
localization_priority: Normal
ms.openlocfilehash: dc049aa681ae4f7664d5bd92f925e7566c0d7103
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241043"
---
# <a name="runtime-for-excel-custom-functions"></a><span data-ttu-id="55c3e-103">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="55c3e-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="55c3e-104">カスタム関数は、作業ウィンドウやその他の UI 要素など、アドインの他の部分で使用されるランタイムとは異なる新しい JavaScript ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="55c3e-105">この JavaScript ランタイムは、カスタム関数での計算のパフォーマンスを最適化するよう設計されており、外部データの要求やサーバーとの固定接続によるデータ交換など、カスタム関数内で一般的な Web ベース アクションを実行する際に使用可能な新しい API を公開します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="55c3e-106">JavaScript ランタイムは、カスタム関数内またはアドインの他の部分で使用してデータを格納、または、ダイアログボックスを表示するために使用できる、`OfficeRuntime` 名前空間内の新しい API へのアクセスも提供します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="55c3e-107">この記事では、カスタム関数内でこれらの API を使用する方法について説明し、カスタム関数を開発する際に留意する事項についても説明します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="55c3e-108">外部データの要求</span><span class="sxs-lookup"><span data-stu-id="55c3e-108">Requesting external data</span></span>

<span data-ttu-id="55c3e-109">カスタム関数内では、[Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="55c3e-110">カスタム関数によって使用される JavaScript ランタイムでは、XHR は[同じ送信元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)と単純な[CORS](https://www.w3.org/TR/cors/)を要求することによって、追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="55c3e-111">単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="55c3e-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="55c3e-112">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="55c3e-113">コンテンツ`Content-Type`タイプが`application/x-www-form-urlencoded` `text/plain`、、またはの場合は`multipart/form-data`、単純な CORS のヘッダーを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="55c3e-114">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="55c3e-114">XHR example</span></span>

<span data-ttu-id="55c3e-115">以下のコード サンプルでは、`getTemperature` 関数が `sendWebRequest` 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="55c3e-116">`sendWebRequest` 関数は、XHR を使用して、データを提供するエンドポイントを要求する `GET` リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="55c3e-117">fetch または XHR を使用すると、新しい JavaScript `Promise` が返されます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="55c3e-118">2018 年 9 月より前は、Office JavaScript API 内で Promise を使用するには `OfficeExtension.Promise` を指定する必要がありましたが、現在は JavaScript `Promise` を使用できます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

```js
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
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="55c3e-119">WebSocket を使用したデータ受信</span><span class="sxs-lookup"><span data-stu-id="55c3e-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="55c3e-120">カスタム関数内で、[WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) を使用して、サーバーとの固定接続経由でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="55c3e-121">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="55c3e-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="55c3e-122">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="55c3e-122">WebSockets example</span></span>

<span data-ttu-id="55c3e-123">以下のコード サンプルでは、`WebSocket` 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span>

```js
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="55c3e-124">データの格納およびアクセス</span><span class="sxs-lookup"><span data-stu-id="55c3e-124">Storing and accessing data</span></span>

<span data-ttu-id="55c3e-125">カスタム関数 (またはアドインの他の部分) 内で、`OfficeRuntime.storage` オブジェクトを使用して、データの格納とデータへのアクセスを実行することができます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="55c3e-126">`Storage` は、カスタム関数内では使用できない [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) の代わりとして使用できる、暗号化されていない永続的キー値ストレージ システムです。</span><span class="sxs-lookup"><span data-stu-id="55c3e-126">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="55c3e-127">`Storage`ドメインごとに 10 MB のデータを提供します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-127">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="55c3e-128">ドメインは複数のアドインで共有できます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-128">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="55c3e-129">`Storage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="55c3e-129">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="55c3e-130">たとえば、ユーザー認証用のトークンを `storage` に保存し、カスタム関数と、作業ウィンドウなどのアドイン UI 要素の両方が、そのトークンにアクセスできるようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-130">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="55c3e-131">同様に、2つのアドインが同じドメイン (たとえば`www.contoso.com/addin1`、など`www.contoso.com/addin2`) を共有している場合は、情報を相互間`storage`で共有することもできます。</span><span class="sxs-lookup"><span data-stu-id="55c3e-131">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="55c3e-132">サブドメインが異なるアドインは、の`storage`インスタンスが異なることに注意してください ( `subdomain.contoso.com/addin1`例`differentsubdomain.contoso.com/addin2`:)。</span><span class="sxs-lookup"><span data-stu-id="55c3e-132">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="55c3e-133">`storage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="55c3e-133">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="55c3e-134">`storage` オブジェクトでは、以下の方法が利用可能です。</span><span class="sxs-lookup"><span data-stu-id="55c3e-134">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="55c3e-135">.</span><span class="sxs-lookup"><span data-stu-id="55c3e-135">.</span></span>[!NOTE]
> <span data-ttu-id="55c3e-136">すべての情報 (など`clear`) を消去する方法はありません。</span><span class="sxs-lookup"><span data-stu-id="55c3e-136">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="55c3e-137">代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。</span><span class="sxs-lookup"><span data-stu-id="55c3e-137">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="55c3e-138">一例</span><span class="sxs-lookup"><span data-stu-id="55c3e-138">OfficeRuntime.storage example</span></span>

<span data-ttu-id="55c3e-139">次のコードサンプルでは`OfficeRuntime.storage.setItem` 、関数を呼び出してキーと値`storage`をに設定します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-139">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="55c3e-140">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="55c3e-140">Additional considerations</span></span>

<span data-ttu-id="55c3e-141">複数のプラットフォーム (Office アドインのキー テナントの 1 つ) で実行するアドインを作成するには、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用したりしないでください。</span><span class="sxs-lookup"><span data-stu-id="55c3e-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="55c3e-142">カスタム関数が JavaScript ランタイムを使用する Windows 上の Excel では、カスタム関数は DOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="55c3e-142">In Excel on Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="55c3e-143">次の手順</span><span class="sxs-lookup"><span data-stu-id="55c3e-143">Next steps</span></span>
<span data-ttu-id="55c3e-144">[カスタム関数を使用して web 要求を実行](custom-functions-web-reqs.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="55c3e-144">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="55c3e-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="55c3e-145">See also</span></span>

* [<span data-ttu-id="55c3e-146">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="55c3e-146">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="55c3e-147">カスタム関数のアーキテクチャ</span><span class="sxs-lookup"><span data-stu-id="55c3e-147">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="55c3e-148">カスタム関数にダイアログを表示する</span><span class="sxs-lookup"><span data-stu-id="55c3e-148">Display a dialog in custom functions</span></span>](custom-functions-dialog.md)
* [<span data-ttu-id="55c3e-149">カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="55c3e-149">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
