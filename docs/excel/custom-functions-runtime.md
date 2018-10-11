---
ms.date: 10/03/2018
description: 新しい JavaScript ランタイムを使用する Excel のカスタム機能開発の主要なシナリオを理解しましょう。
title: Excel カスタム関数のランタイム
ms.openlocfilehash: a48b02a8ca404b51740d9052d199da934eb9312e
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459106"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Excel カスタム関数のランタイム (プレビュー)

カスタム関数は、作業ウィンドウやその他の UI 要素など、アドインの他の部分で用いられるランタイムとは異なる、新しい JavaScript ランタイムを使用します。 この JavaScript ランタイムは、カスタム関数での計算のパフォーマンスを最適化するよう設計されており、外部データの要求やサーバーとの固定接続によるデータ交換など、カスタム関数内で一般的な Web ベースアクションを実行する際に使用可能な、新しい API を公開します。 JavaScript ランタイムは、カスタム関数内またはアドインの他の部分で使用してデータを格納、または、ダイアログボックスを表示するために使用できる、`OfficeRuntime` 名前空間内の新しい API へのアクセスも提供します。 この記事では、これらのAPIをカスタム関数内で使用する方法と、カスタム関数を展開する際に留意すべき追加の考慮事項について説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>外部データの要求

カスタム関数内では、[ Fetch ](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である[   XmlHttpRequest (XHR) ](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます 。 JavaScript ランタイムでは、 XHR は[同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな[ CORS ](https://www.w3.org/TR/cors/)を要求することにより、追加セキュリティ対策を実装します。  

### <a name="xhr-example"></a>XHR の使用例

以下のコードサンプルでは、`getTemperature`関数は`sendWebRequest`関数を呼び出して温度計IDに基づく特定の領域の温度を取得します。 `sendWebRequest` 関数は、XHR を使用してデータを提供するエンドポイントへの`GET`要求を発行します。 

> [!NOTE] 
> fetch または XHR を使用すると、新しい JavaScript  `Promise`が返されます。 2018年9月より前は、Office JavaScript API 内で約束を使用するには`OfficeExtension.Promise`を指定する必要がありましたが、今は JavaScript  `Promise`を使用するだけです。

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

## <a name="receiving-data-via-websockets"></a>Websocket を使用したデータ受信

カスタム関数内部サーバーとの固定接続を介してのデータ交換には、 [ Websocket ](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) を使用できます。 WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信しますので、サーバーに明示的にデータをポーリングする必要がありません。

### <a name="websockets-example"></a>Websocket の使用例

以下のコードサンプルは`WebSocket`接続を確立し、サーバーからの各受信メッセージを記録します。 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>データの格納およびアクセス

カスタム関数（またはアドインの他の部分）内では、`OfficeRuntime.AsyncStorage`オブジェクトを使用してデータを格納およびアクセスできます。 `AsyncStorage` [X]は、[  localStorage ](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) の代替機能を提供する、暗号化されていない永続的キー値ストレージシステムであり、カスタム関数内では使用できません。 アドインは、`AsyncStorage`を使用して最大 10 MB のデータを格納できます。

`AsyncStorage`オブジェクトでは、以下のメソッドを使用できます。
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a>AsyncStorage の使用例 

以下のコードサンプルは、`AsyncStorage.getItem`関数を呼び出してストレージから値を取得します。

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

## <a name="displaying-a-dialog-box"></a>ダイアログボックスの表示

カスタム関数（またはアドインの他の部分）内では、`OfficeRuntime.displayWebDialogOptions`  API を使用してダイアログボックスを表示できます。 このダイアログボックス API は、[ Dialog API ](../develop/dialog-api-in-office-add-ins.md) の代わりに作業ウィンドウやアドインコマンドで使用できますが、カスタム関数では使用できません。

### <a name="dialog-api-example"></a>ダイアログ API の使用例 

以下のコードサンプルでは、関数`getTokenViaDialog`は Dialog API の`displayWebDialogOptions`関数を使用してダイアログボックスを表示しています。

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

## <a name="additional-considerations"></a>その他の考慮事項

複数のプラットフォーム（Officeアドインの主要テナントの一つ）で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQueryのようなDOMに依存するライブラリーを使用してはいけません。 カスタム関数が JavaScript ランタイムを使用する Excel for Windows では、カスタム関数は DOM にアクセスできません。

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
