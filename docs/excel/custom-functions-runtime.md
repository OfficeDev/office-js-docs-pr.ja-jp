---
ms.date: 02/06/2019
description: 新しい JavaScript ランタイムを使用する Excel カスタム関数を開発する場合の重要なシナリオについて、理解します。
title: Excel カスタム関数のランタイム (プレビュー)
localization_priority: Normal
ms.openlocfilehash: d891a41dc9e142ef3cfaa00c8b54d8d27913c57d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982042"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="0d550-103">Excel カスタム関数のランタイム (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="0d550-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="0d550-104">カスタム関数は、作業ウィンドウやその他の UI 要素など、アドインの他の部分で使用されるランタイムとは異なる新しい JavaScript ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="0d550-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="0d550-105">この JavaScript ランタイムは、カスタム関数での計算のパフォーマンスを最適化するよう設計されており、外部データの要求やサーバーとの固定接続によるデータ交換など、カスタム関数内で一般的な Web ベース アクションを実行する際に使用可能な新しい API を公開します。</span><span class="sxs-lookup"><span data-stu-id="0d550-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="0d550-106">JavaScript ランタイムは、カスタム関数内またはアドインの他の部分で使用してデータを格納、または、ダイアログボックスを表示するために使用できる、`OfficeRuntime` 名前空間内の新しい API へのアクセスも提供します。</span><span class="sxs-lookup"><span data-stu-id="0d550-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="0d550-107">この記事では、カスタム関数内でこれらの API を使用する方法について説明し、カスタム関数を開発する際に留意する事項についても説明します。</span><span class="sxs-lookup"><span data-stu-id="0d550-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="0d550-108">外部データの要求</span><span class="sxs-lookup"><span data-stu-id="0d550-108">Requesting external data</span></span>

<span data-ttu-id="0d550-109">カスタム関数内では、[Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます。</span><span class="sxs-lookup"><span data-stu-id="0d550-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="0d550-110">カスタム関数で使用される JavaScript ランタイム内では、XHR は、[同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな[CORS](https://www.w3.org/TR/cors/)を要求することによって追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="0d550-110">Within the JavaScript runtime used by custom functions, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="0d550-111">CORS の単純な実装は cookie を使用することはできません、単純なメソッド (GET、HEAD、POST) のみをサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="0d550-111">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="0d550-112">フィールド名を持つ単純なヘッダーを受け入れる単純な CORS `Accept`、 `Accept-Language`、 `Content-Language`。</span><span class="sxs-lookup"><span data-stu-id="0d550-112">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="0d550-113">使用することも、`Content-Type`で単純な CORS は、ヘッダーのコンテンツ型があるが提供される`application/x-www-form-urlencoded`、 `text/plain`、または`multipart/form-data`。</span><span class="sxs-lookup"><span data-stu-id="0d550-113">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="0d550-114">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="0d550-114">XHR example</span></span>

<span data-ttu-id="0d550-115">以下のコード サンプルでは、`getTemperature` 関数が `sendWebRequest` 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。</span><span class="sxs-lookup"><span data-stu-id="0d550-115">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="0d550-116">`sendWebRequest` 関数は、XHR を使用して、データを提供するエンドポイントを要求する `GET` リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="0d550-116">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>

> [!NOTE] 
> <span data-ttu-id="0d550-117">fetch または XHR を使用すると、新しい JavaScript `Promise` が返されます。</span><span class="sxs-lookup"><span data-stu-id="0d550-117">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="0d550-118">2018 年 9 月より前は、Office JavaScript API 内で Promise を使用するには `OfficeExtension.Promise` を指定する必要がありましたが、現在は JavaScript `Promise` を使用できます。</span><span class="sxs-lookup"><span data-stu-id="0d550-118">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="0d550-119">WebSocket を使用したデータ受信</span><span class="sxs-lookup"><span data-stu-id="0d550-119">Receiving data via WebSockets</span></span>

<span data-ttu-id="0d550-120">カスタム関数内で、[WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) を使用して、サーバーとの固定接続経由でデータを交換することができます。</span><span class="sxs-lookup"><span data-stu-id="0d550-120">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="0d550-121">WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。</span><span class="sxs-lookup"><span data-stu-id="0d550-121">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="0d550-122">WebSocket の使用例</span><span class="sxs-lookup"><span data-stu-id="0d550-122">WebSockets example</span></span>

<span data-ttu-id="0d550-123">以下のコード サンプルでは、`WebSocket` 接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="0d550-123">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="0d550-124">データの格納およびアクセス</span><span class="sxs-lookup"><span data-stu-id="0d550-124">Storing and accessing data</span></span>

<span data-ttu-id="0d550-125">カスタム関数 (またはアドインの他の部分) 内で、`OfficeRuntime.AsyncStorage` オブジェクトを使用して、データの格納とデータへのアクセスを実行することができます。</span><span class="sxs-lookup"><span data-stu-id="0d550-125">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="0d550-126">`AsyncStorage` は、カスタム関数内では使用できない [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) の代わりとして使用できる、暗号化されていない永続的キー値ストレージ システムです。</span><span class="sxs-lookup"><span data-stu-id="0d550-126">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="0d550-127">アドインは `AsyncStorage` を使用すると、最大 10 MB のデータを格納できます。</span><span class="sxs-lookup"><span data-stu-id="0d550-127">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="0d550-128">`AsyncStorage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="0d550-128">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="0d550-129">たとえば、ユーザー認証用のトークンを `AsyncStorage` に保存し、カスタム関数と、作業ウィンドウなどのアドイン UI 要素の両方が、そのトークンにアクセスできるようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="0d550-129">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="0d550-130">同様に、2 つのアドインが同じドメインを共有している場合 (例: www.contoso.com/addin1、www.contoso.com/addin2)、アドイン間で `AsyncStorage` を介して情報を共有できるようにすることができます。</span><span class="sxs-lookup"><span data-stu-id="0d550-130">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="0d550-131">サブドメインが異なるアドインについては (例: subdomain.contoso.com/addin1、differentsubdomain.contoso.com/addin2)、`AsyncStorage` インスタンスも別々となることに留意してください。</span><span class="sxs-lookup"><span data-stu-id="0d550-131">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="0d550-132">`AsyncStorage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="0d550-132">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="0d550-133">`AsyncStorage` オブジェクトでは、以下の方法が利用可能です。</span><span class="sxs-lookup"><span data-stu-id="0d550-133">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="0d550-134">`multiRemove`: すべての情報をクリアする方法 (`clear` など) は実装されていません。</span><span class="sxs-lookup"><span data-stu-id="0d550-134">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="0d550-135">代わりに、一度に複数のエントリを削除できる `multiRemove` を使用してください。</span><span class="sxs-lookup"><span data-stu-id="0d550-135">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="0d550-136">AsyncStorage の使用例</span><span class="sxs-lookup"><span data-stu-id="0d550-136">AsyncStorage example</span></span> 

<span data-ttu-id="0d550-137">以下のコード サンプルでは、`AsyncStorage.getItem` 関数を呼び出してストレージから値を取得します。</span><span class="sxs-lookup"><span data-stu-id="0d550-137">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```typescript
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
        }
    } catch (error) {
        //handle errors here
    }
}
```

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="0d550-138">ダイアログ ボックスの表示</span><span class="sxs-lookup"><span data-stu-id="0d550-138">Displaying a dialog box</span></span>

<span data-ttu-id="0d550-139">カスタム関数 (またはアドインの他の部分) 内で、`OfficeRuntime.displayWebDialog` API を使用してダイアログ ボックスを表示することができます。</span><span class="sxs-lookup"><span data-stu-id="0d550-139">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialog` API to display a dialog box.</span></span> <span data-ttu-id="0d550-140">このダイアログ API は、作業ウィンドウとアドイン コマンド内では使用可能であるが、カスタム関数内では使用できない[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) の代わりに、使用できます。</span><span class="sxs-lookup"><span data-stu-id="0d550-140">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="0d550-141">ダイアログ API の使用例</span><span class="sxs-lookup"><span data-stu-id="0d550-141">Dialog API example</span></span>

<span data-ttu-id="0d550-142">以下のコード サンプルでは、関数 `getTokenViaDialog` がダイアログ API の `displayWebDialog` 関数を使用して、ダイアログ ボックスを表示します。</span><span class="sxs-lookup"><span data-stu-id="0d550-142">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
        let timeout = 5;
        let count = 0;
        var intervalId = setInterval(function () {
          count++;
          if(_cachedToken) {
            resolve(_cachedToken);
            clearInterval(intervalId);
          }
          if(count >= timeout) {
            reject("Timeout while waiting for token");
            clearInterval(intervalId);
          }
        }, 1000);
      } else {
        _dialogOpen = true;
        OfficeRuntime.displayWebDialog(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.close();
            return;
          },
          onRuntimeError: function(error, dialog) {
            reject(error);
          },
        }).catch(function (e) {
          reject(e);
        });
      }
    });
  }
}
```

## <a name="additional-considerations"></a><span data-ttu-id="0d550-143">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="0d550-143">Additional considerations</span></span>

<span data-ttu-id="0d550-144">複数のプラットフォーム (Office アドインのキー テナントの 1 つ) で実行するアドインを作成するには、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用したりしないでください。</span><span class="sxs-lookup"><span data-stu-id="0d550-144">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="0d550-145">カスタム関数が JavaScript ランタイムを使用する Excel for Windows では、カスタム関数は DOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="0d550-145">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="0d550-146">関連項目</span><span class="sxs-lookup"><span data-stu-id="0d550-146">See also</span></span>

* [<span data-ttu-id="0d550-147">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0d550-147">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0d550-148">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0d550-148">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0d550-149">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0d550-149">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0d550-150">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="0d550-150">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="0d550-151">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0d550-151">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
