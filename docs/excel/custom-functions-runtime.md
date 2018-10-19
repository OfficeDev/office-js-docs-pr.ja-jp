---
ms.date: 10/17/2018
description: 新しい JavaScript ランタイムを使用する Excel のカスタム機能開発の主要なシナリオを理解しましょう。
title: Excel カスタム関数のランタイム
ms.openlocfilehash: 333816c3916af1490d14b8344c4bb49094f9a7f9
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640016"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="8f051-103">Excel カスタム関数のランタイム (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8f051-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="8f051-p101">カスタム関数は、作業ウィンドウやその他の UI 要素など、アドインの他の部分で使用されるランタイムとは異なる新しい JavaScript ランタイムを使用します。この JavaScript ランタイムは、カスタム関数での計算のパフォーマンスを最適化するよう設計されており、外部データの要求やサーバーとの固定接続によるデータ交換など、カスタム関数内で一般的な Web ベース アクションを実行する際に使用可能な新しい API を公開します。JavaScript ランタイムは、カスタム関数内またはアドインの他の部分で使用してデータを格納、または、ダイアログボックスを表示するために使用できる、`OfficeRuntime` 名前空間内の新しい API へのアクセスも提供します。この項目では、カスタム関数内でこれらの API を使用する方法について説明し、カスタム関数を開発する際に留意する事項についても説明します。</span><span class="sxs-lookup"><span data-stu-id="8f051-p101">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server. The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box. This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="8f051-108">外部データの要求</span><span class="sxs-lookup"><span data-stu-id="8f051-108">Requesting external data</span></span>

<span data-ttu-id="8f051-p102">カスタム関数内では、[ Fetch ](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である[  XmlHttpRequest (XHR) ](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます 。JavaScript ランタイムでは、 XHR は[ 同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな[  CORS ](https://www.w3.org/TR/cors/)を要求することにより、追加セキュリティ対策を実装します。</span><span class="sxs-lookup"><span data-stu-id="8f051-p102">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="8f051-111">XHR の使用例</span><span class="sxs-lookup"><span data-stu-id="8f051-111">XHR example</span></span>

<span data-ttu-id="8f051-p103">以下のコード サンプルでは、 `getTemperature` 関数は `sendWebRequest` 関数を呼び出して温度計 ID に基づく特定の領域の温度を取得します。`sendWebRequest` 関数は、XHR を使用して、データを提供するエンドポイントを要求する `GET` リクエストを発行します。</span><span class="sxs-lookup"><span data-stu-id="8f051-p103">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID. The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="8f051-p104">fetch または XHR を使用すると、新しい JavaScript  `Promise`が返されます。2018年9月より前は、Office JavaScript API 内で約束を使用するには`OfficeExtension.Promise` を指定する必要がありましたが、今は JavaScript  `Promise` を使用するだけです</span><span class="sxs-lookup"><span data-stu-id="8f051-p104">When using fetch or XHR, a new JavaScript `Promise` is returned. Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="8f051-116">Websocket を使用したデータ受信</span><span class="sxs-lookup"><span data-stu-id="8f051-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="8f051-p105">カスタム関数内部サーバーとの固定接続を介してのデータ交換には、 [Websocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) を使用できます。WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータをポーリングする必要がありません。</span><span class="sxs-lookup"><span data-stu-id="8f051-p105">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="8f051-119">Websocket の使用例</span><span class="sxs-lookup"><span data-stu-id="8f051-119">WebSockets example</span></span>

<span data-ttu-id="8f051-120">以下のコードサンプルは`WebSocket`接続を確立し、サーバーからの各受信メッセージを記録します。</span><span class="sxs-lookup"><span data-stu-id="8f051-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="8f051-121">データの格納およびアクセス</span><span class="sxs-lookup"><span data-stu-id="8f051-121">Storing and accessing data</span></span>

<span data-ttu-id="8f051-p106">カスタム関数 (またはアドインの他の部分) 内では、`OfficeRuntime.AsyncStorage` オブジェクトを使用してデータを格納およびアクセできます。 `AsyncStorage` は、[localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage)の代替機能を提供する、暗号化されていない永続的キー値ストレージ システムであり、カスタム関数内では使用できません。`AsyncStorage` を使用すると、アドインには最大 10 MB のデータを格納できます 。</span><span class="sxs-lookup"><span data-stu-id="8f051-p106">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object. `AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions. An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="8f051-125">`AsyncStorage` オブジェクトでは、以下の方法が利用可能です。</span><span class="sxs-lookup"><span data-stu-id="8f051-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="8f051-126">AsyncStorage の使用例</span><span class="sxs-lookup"><span data-stu-id="8f051-126">AsyncStorage example</span></span> 

<span data-ttu-id="8f051-127">以下のコードサンプルは、`AsyncStorage.getItem`関数を呼び出してストレージから値を取得します。</span><span class="sxs-lookup"><span data-stu-id="8f051-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="8f051-128">ダイアログボックスの表示</span><span class="sxs-lookup"><span data-stu-id="8f051-128">Open a dialog box</span></span>

<span data-ttu-id="8f051-p107">カスタム関数（またはアドインの他の部分）内では、`OfficeRuntime.displayWebDialogOptions`  API を使用してダイアログボックスを表示できます。このダイアログボックス API は、[ Dialog API ](../develop/dialog-api-in-office-add-ins.md) の代わりに作業ウィンドウやアドインコマンドで使用できますが、カスタム関数では使用できません。</span><span class="sxs-lookup"><span data-stu-id="8f051-p107">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box. This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="8f051-131">ダイアログ API の使用例</span><span class="sxs-lookup"><span data-stu-id="8f051-131">Dialog API example</span></span> 

<span data-ttu-id="8f051-132">以下のコードサンプルでは、関数`getTokenViaDialog`は Dialog API の`displayWebDialogOptions`関数を使用してダイアログボックスを表示しています。</span><span class="sxs-lookup"><span data-stu-id="8f051-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

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
        OfficeRuntime.displayWebDialogOptions(url, {
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

## <a name="additional-considerations"></a><span data-ttu-id="8f051-133">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="8f051-133">Additional considerations</span></span>

<span data-ttu-id="8f051-p108">複数のプラットフォーム（Officeアドインのキー テナントの一つ）で実行するアドインを生成するには、カスタム関数でドキュメント オブジェクト モデル(DOM)にアクセスしたり、jQueryのようなDOMに依存するライブラリーを使用してはいけません。カスタム関数が JavaScript ランタイムを使用する Excel for Windows では、カスタム関数は DOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="8f051-p108">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="8f051-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="8f051-136">See also</span></span>

* [<span data-ttu-id="8f051-137">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="8f051-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="8f051-138">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="8f051-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8f051-139">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8f051-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8f051-140">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="8f051-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
