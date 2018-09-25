---
ms.date: 09/20/2018
description: Excel のカスタム関数は、標準のアドインの WebView コントロールのランタイムと異なる、新しい JavaScript ランタイムを使用します。
title: Excel のカスタム関数のランタイム
ms.openlocfilehash: fa2b2030259e05f64b8b4660ded8b80c6af1eb5a
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985796"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="41be8-103">Excel カスタム関数のランタイム (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="41be8-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="41be8-104">カスタム関数は、web ブラウザーではなく、サンドボックス JavaScript エンジンを使用する新しい JavaScript ランタイムを使用して、Excel の機能を拡張します。</span><span class="sxs-lookup"><span data-stu-id="41be8-104">Custom functions extend Excel’s capabilities by using a new JavaScript runtime that uses a sandboxed JavaScript engine rather than a web browser.</span></span> <span data-ttu-id="41be8-105">カスタム関数は UI 要素をレンダリングする必要がなく、新しい JavaScript のランタイムは計算に最適化されているため、何千ものカスタム関数を同時に実行できます。</span><span class="sxs-lookup"><span data-stu-id="41be8-105">Because custom functions do not need to render UI elements, the new JavaScript runtime is optimized for performing calculations, enabling you to run thousands of custom functions simultaneously.</span></span>

## <a name="key-facts-about-the-new-javascript-runtime"></a><span data-ttu-id="41be8-106">新しい JavaScript ランタイムに関する重要な事実</span><span class="sxs-lookup"><span data-stu-id="41be8-106">Key facts about the new JavaScript runtime</span></span> 

<span data-ttu-id="41be8-107">アドイン内のカスタム関数だけが、この記事で説明する新しい JavaScript ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="41be8-107">Only custom functions within an add-in will use the new JavaScript runtime that's described in this article.</span></span> <span data-ttu-id="41be8-108">カスタム関数に加え、作業ウィンドウや他の UI 要素など他のコンポーネントがアドインに含まれる場合、アドインのこれら他のコンポーネントは、ブラウザーのような WebView ランタイムで引き続き実行されます。</span><span class="sxs-lookup"><span data-stu-id="41be8-108">If an add-in includes other components such as task panes and other UI elements, in addition to custom functions, these other components of the add-in will continue to run in the browser-like WebView runtime.</span></span>  <span data-ttu-id="41be8-109">さらに:</span><span class="sxs-lookup"><span data-stu-id="41be8-109">Additionally:</span></span> 

- <span data-ttu-id="41be8-110">JavaScript ランタイムは、ドキュメント オブジェクト モデル (DOM)、または DOM に依存している jQuery のようなサポート ライブラリ へのアクセスを行いません。</span><span class="sxs-lookup"><span data-stu-id="41be8-110">The JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM.</span></span>

- <span data-ttu-id="41be8-111">アドインの JavaScript ファイルで定義されているカスタム関数は、`Promise` を返す代わりに `OfficeExtension.Promise`通常の JavaScript を返すことができます。</span><span class="sxs-lookup"><span data-stu-id="41be8-111">A custom function that's defined in an add-in's JavaScript file can return a regular JavaScript `Promise` instead of returning `OfficeExtension.Promise`.</span></span>  

- <span data-ttu-id="41be8-112">カスタム関数メタデータを指定する JSON ファイルは、**オプション** 内で **同期**または**非同期**を指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="41be8-112">The JSON file that specifies custom function metatdata does not need to specify **sync** or **async** within **options**.</span></span>

## <a name="new-apis"></a><span data-ttu-id="41be8-113">新しい API</span><span class="sxs-lookup"><span data-stu-id="41be8-113">New and updated APIs</span></span> 

<span data-ttu-id="41be8-114">カスタム関数で使用されている JavaScript のランタイムには、次の API があります。</span><span class="sxs-lookup"><span data-stu-id="41be8-114">The JavaScript runtime that's used by custom functions has the following APIs:</span></span>

- [<span data-ttu-id="41be8-115">XHR</span><span class="sxs-lookup"><span data-stu-id="41be8-115">XHR</span></span>](#xhr)
- [<span data-ttu-id="41be8-116">WebSocket</span><span class="sxs-lookup"><span data-stu-id="41be8-116">WebSockets</span></span>](#websockets)
- [<span data-ttu-id="41be8-117">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="41be8-117">AsyncStorage</span></span>](#asyncstorage)
- [<span data-ttu-id="41be8-118">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="41be8-118">Dialog API requirement sets</span></span>](#dialog-api)

### <a name="xhr"></a><span data-ttu-id="41be8-119">XHR</span><span class="sxs-lookup"><span data-stu-id="41be8-119">XHR</span></span>

<span data-ttu-id="41be8-120">XHR は [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を表し、これはサーバーと対話する HTTP 要求を発行する標準的な web API です。</span><span class="sxs-lookup"><span data-stu-id="41be8-120">XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="41be8-121">新しい JavaScript ランタイムでは、XHR は[同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな[CORS](https://www.w3.org/TR/cors/)を要求することによって追加のセキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="41be8-121">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

<span data-ttu-id="41be8-122">次のコード例で、 `getTemperature()` 関数は、温度計の ID に基づいて、特定の領域の温度を取得する web 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="41be8-122">In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="41be8-123">関数は、XHR を使用して、データを提供するエンドポイントへの`GET`要求を発行します。`sendWebRequest()`</span><span class="sxs-lookup"><span data-stu-id="41be8-123">The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
          };
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

```

### <a name="websockets"></a><span data-ttu-id="41be8-124">WebSocket</span><span class="sxs-lookup"><span data-stu-id="41be8-124">WebSockets</span></span>

<span data-ttu-id="41be8-125">[Websocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) は、サーバーと 1 つ以上のクライアント間でリアルタイムのコミュニケーションを作成するネットワーク プロトコルです。</span><span class="sxs-lookup"><span data-stu-id="41be8-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients.</span></span> <span data-ttu-id="41be8-126">テキストを同時に読み書きすることができるので、多くの場合チャット アプリケーションに使用します。</span><span class="sxs-lookup"><span data-stu-id="41be8-126">It is often used for chat applications because it allows you to read and write text simultaneously.</span></span>  

<span data-ttu-id="41be8-127">次のコード サンプルに示すように、カスタム関数は Websocket を使用できます。</span><span class="sxs-lookup"><span data-stu-id="41be8-127">As shown in the following code sample, custom functions can use WebSockets.</span></span> <span data-ttu-id="41be8-128">この例では、WebSocket は、受信した各メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="41be8-128">In this example, the WebSocket logs each message that it receives.</span></span>

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a><span data-ttu-id="41be8-129">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="41be8-129">AsyncStorage</span></span>

<span data-ttu-id="41be8-130">AsyncStorage は、認証トークンを格納するために使用するキーと値のストレージ システムです。</span><span class="sxs-lookup"><span data-stu-id="41be8-130">AsyncStorage is a key-value storage system that can be used to store authentication tokens.</span></span> <span data-ttu-id="41be8-131">たとえば、</span><span class="sxs-lookup"><span data-stu-id="41be8-131">It is framework-agnostic.</span></span>

- <span data-ttu-id="41be8-132">持続性</span><span class="sxs-lookup"><span data-stu-id="41be8-132">persistent</span></span>
- <span data-ttu-id="41be8-133">暗号化なし</span><span class="sxs-lookup"><span data-stu-id="41be8-133">Unencrypted</span></span>
- <span data-ttu-id="41be8-134">非同期</span><span class="sxs-lookup"><span data-stu-id="41be8-134">Asynchronous calls</span></span>

<span data-ttu-id="41be8-135">AsyncStorage は、アドイン内のすべての部分にグローバルに利用できます。</span><span class="sxs-lookup"><span data-stu-id="41be8-135">AsyncStorage is globally available to all parts of your add-in.</span></span> <span data-ttu-id="41be8-136">カスタム関数では、 `AsyncStorage` は、グローバル オブジェクトとして公開されます。</span><span class="sxs-lookup"><span data-stu-id="41be8-136">For custom functions, `AsyncStorage` is exposed as a global object.</span></span> <span data-ttu-id="41be8-137">(WebView ランタイムを使用する作業ウィンドウおよびその他の要素などのアドインの他の部分では、`OfficeRuntime` を通じて AsyncStorage が公開されます。) 各アドインは、既定サイズが 5 MB の独自のストレージ パーティションを持ちます。</span><span class="sxs-lookup"><span data-stu-id="41be8-137">(For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB.</span></span> 

<span data-ttu-id="41be8-138">オブジェクトでは、以下の方法が利用可能です。`AsyncStorage`</span><span class="sxs-lookup"><span data-stu-id="41be8-138">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
<span data-ttu-id="41be8-139">この時点で、 `mergeItem` と `multiMerge` のメソッドはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="41be8-139">At this time, the `mergeItem` and `multiMerge` methods are not supported.</span></span>

<span data-ttu-id="41be8-140">次のコード サンプルは、ストレージから値を取得するために `AsyncStorage.getItem` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="41be8-140">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```js
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
}
```

### <a name="dialog-api"></a><span data-ttu-id="41be8-141">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="41be8-141">Dialog API scenarios</span></span>

<span data-ttu-id="41be8-142">ダイアログ API を使用すると、ユーザーのサインインを求めるダイアログ ボックスを開くことができます。</span><span class="sxs-lookup"><span data-stu-id="41be8-142">The Dialog API enables you to open a dialog box that prompts user sign-in.</span></span> <span data-ttu-id="41be8-143">ユーザーが関数を使用する前に、Google や Facebook などの外部のリソースを通じ、ダイアログ API を使用してユーザー認証を要求します。</span><span class="sxs-lookup"><span data-stu-id="41be8-143">You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.</span></span>   

<span data-ttu-id="41be8-144">次のコード サンプルで、 `getTokenViaDialog()` メソッドは、ダイアログ API の `displayWebDialog()` メソッドを使用してダイアログ ボックスを開きます。</span><span class="sxs-lookup"><span data-stu-id="41be8-144">In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"
 
function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://myauthurl")
    .then(function (token) {
      
      // Use token to get stock price
      fetch("https://myservice.com/?token=token&ticker= + ticker")
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
            dialog.closeDialog();
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

> [!NOTE]
> <span data-ttu-id="41be8-145">このセクションで説明しているダイアログ API は、カスタム関数の新しい JavaScript ランタイムの一部であり、カスタム関数内でのみ使用することができます。</span><span class="sxs-lookup"><span data-stu-id="41be8-145">The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions.</span></span> <span data-ttu-id="41be8-146">この API は、作業ウィンドウおよびアドイン コマンド内で使用できる [ダイアログ API](../develop/dialog-api-in-office-add-ins.md) とは異なります。</span><span class="sxs-lookup"><span data-stu-id="41be8-146">This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.</span></span>

## <a name="see-also"></a><span data-ttu-id="41be8-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="41be8-147">See also</span></span>

* [<span data-ttu-id="41be8-148">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="41be8-148">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="41be8-149">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="41be8-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="41be8-150">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="41be8-150">Custom functions best practices</span></span>](custom-functions-best-practices.md)
