---
ms.date: 09/20/2018
description: Excel のカスタム関数は、標準のアドインの WebView コントロールのランタイムと異なる、新しい JavaScript ランタイムを使用します。
title: Excel のカスタム関数のランタイム
ms.openlocfilehash: d31002096fccd682c0f2a23a8b43249af5d4df8f
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068824"
---
# <a name="runtime-for-excel-custom-functions"></a>Excel のカスタム関数のランタイム

カスタム関数は、web ブラウザーではなく、サンドボックス JavaScript エンジンを使用する新しい JavaScript ランタイムを使用して、Excel の機能を拡張します。 カスタム関数は UI 要素をレンダリングする必要がなく、新しい JavaScript のランタイムは計算に最適化されているため、何千ものカスタム関数を同時に実行できます。

## <a name="key-facts-about-the-new-javascript-runtime"></a>新しい JavaScript ランタイムに関する重要な事実 

アドイン内のカスタム関数だけが、この記事で説明する新しい JavaScript ランタイムを使用します。 カスタム関数に加え、作業ウィンドウや他の UI 要素など他のコンポーネントがアドインに含まれる場合、アドインのこれら他のコンポーネントは、ブラウザーのような WebView ランタイムで引き続き実行されます。  さらに、次の特徴を備えています。 

- JavaScript ランタイムは、ドキュメント オブジェクト モデル (DOM)、または DOM に依存している jQuery のようなサポート ライブラリ へのアクセスを行いません。

- アドインの JavaScript ファイルで定義されているカスタム関数は、`Promise` を返す代わりに `OfficeExtension.Promise`通常の JavaScript を返すことができます。  

- カスタム関数メタデータを指定する JSON ファイルは、**オプション** 内で **同期**または**非同期**を指定する必要はありません。

## <a name="new-apis"></a>新しい API 

カスタム関数で使用されている JavaScript のランタイムには、次の API があります。

- [XHR](#xhr)
- [WebSocket](#websockets)
- [AsyncStorage](#asyncstorage)
- [ダイアログ API](#dialog-api)

### <a name="xhr"></a>XHR

XHR は [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を表し、これはサーバーと対話する HTTP 要求を発行する標準的な web API です。 新しい JavaScript ランタイムでは、XHR は[同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな[CORS](https://www.w3.org/TR/cors/)を要求することによって追加のセキュリティ対策を実装します。  

次のコード例で、 `getTemperature()` 関数は、温度計の ID に基づいて、特定の領域の温度を取得する web 要求を送信します。  `sendWebRequest()`関数は、XHR を使用して、データを提供するエンドポイントへの`GET`要求を発行します。  

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

### <a name="websockets"></a>WebSocket

[Websocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) は、サーバーと 1 つ以上のクライアント間でリアルタイムのコミュニケーションを作成するネットワーク プロトコルです。 テキストを同時に読み書きすることができるので、多くの場合チャット アプリケーションに使用します。  

次のコード サンプルに示すように、カスタム関数は Websocket を使用できます。 この例では、WebSocket は、受信した各メッセージを記録します。

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a>AsyncStorage

AsyncStorage は、認証トークンを格納するために使用するキーと値のストレージ システムです。 たとえば、

- 持続性
- 暗号化なし
- 非同期

AsyncStorage は、アドイン内のすべての部分にグローバルに利用できます。 カスタム関数では、 `AsyncStorage` は、グローバル オブジェクトとして公開されます。 (WebView ランタイムを使用する作業ウィンドウおよびその他の要素などのアドインの他の部分では、`OfficeRuntime` を通じて AsyncStorage が公開されます。) 各アドインは、既定サイズが 5 MB の独自のストレージ パーティションを持ちます。 

 `AsyncStorage` オブジェクトでは、以下の方法が利用可能です。
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
この時点で、 `mergeItem` と `multiMerge` のメソッドはサポートされていません。

次のコード サンプルは、ストレージから値を取得するために `AsyncStorage.getItem` を呼び出します。

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

### <a name="dialog-api"></a>ダイアログ API

ダイアログ API を使用すると、ユーザーのサインインを求めるダイアログ ボックスを開くことができます。 ユーザーが関数を使用する前に、Google や Facebook などの外部のリソースを通じ、ダイアログ API を使用してユーザー認証を要求します。   

次のコード サンプルで、 `getTokenViaDialog()` メソッドは、ダイアログ API の `displayWebDialog()` メソッドを使用してダイアログ ボックスを開きます。

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
> このセクションで説明しているダイアログ API は、カスタム関数の新しい JavaScript ランタイムの一部であり、カスタム関数内でのみ使用することができます。 この API は、作業ウィンドウおよびアドイン コマンド内で使用できる [ダイアログ API](../develop/dialog-api-in-office-add-ins.md) とは異なります。

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)