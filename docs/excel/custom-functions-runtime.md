---
ms.date: 12/5/2018
description: 新しい JavaScript ランタイムを使用する Excel カスタム関数を開発する場合の重要なシナリオについて、理解します。
title: Excel カスタム関数のランタイム
ms.openlocfilehash: 715690c5cba2466e4a50ba2a33d2324a1abe02f5
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270832"
---
# <a name="runtime-for-excel-custom-functions-preview"></a>Excel カスタム関数のランタイム (プレビュー)

カスタム関数は、作業ウィンドウやその他の UI 要素など、アドインの他の部分で使用されるランタイムとは異なる新しい JavaScript ランタイムを使用します。 この JavaScript ランタイムは、カスタム関数での計算のパフォーマンスを最適化するよう設計されており、外部データの要求やサーバーとの固定接続によるデータ交換など、カスタム関数内で一般的な Web ベース アクションを実行する際に使用可能な新しい API を公開します。 JavaScript ランタイムは、カスタム関数内またはアドインの他の部分で使用してデータを格納、または、ダイアログボックスを表示するために使用できる、`OfficeRuntime` 名前空間内の新しい API へのアクセスも提供します。 この記事では、カスタム関数内でこれらの API を使用する方法について説明し、カスタム関数を開発する際に留意する事項についても説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>外部データの要求

カスタム関数内では、[Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) などの API や、サーバーとやり取りする HTTP 要求を発行する標準 Web API である [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) を使用して、外部データを要求できます。 JavaScript ランタイムでは、XHR は[同一生成元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、追加セキュリティ対策を実装します。  

### <a name="xhr-example"></a>XHR の使用例

以下のコード サンプルでは、`getTemperature` 関数が `sendWebRequest` 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。 `sendWebRequest` 関数は、XHR を使用して、データを提供するエンドポイントを要求する `GET` リクエストを発行します。

> [!NOTE] 
> fetch または XHR を使用すると、新しい JavaScript `Promise` が返されます。 2018 年 9 月より前は、Office JavaScript API 内で Promise を使用するには `OfficeExtension.Promise` を指定する必要がありましたが、現在は JavaScript `Promise` を使用できます。

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

## <a name="receiving-data-via-websockets"></a>WebSocket を使用したデータ受信

カスタム関数内で、[WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) を使用して、サーバーとの固定接続経由でデータを交換することができます。 WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。

### <a name="websockets-example"></a>WebSocket の使用例

以下のコード サンプルでは、`WebSocket` 接続を確立し、サーバーからの各受信メッセージを記録します。 

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

カスタム関数 (またはアドインの他の部分) 内で、`OfficeRuntime.AsyncStorage` オブジェクトを使用して、データの格納とデータへのアクセスを実行することができます。 `AsyncStorage` は、カスタム関数内では使用できない [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) の代わりとして使用できる、暗号化されていない永続的キー値ストレージ システムです。 アドインは `AsyncStorage` を使用すると、最大 10 MB のデータを格納できます。

`AsyncStorage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。 たとえば、ユーザー認証用のトークンを `AsyncStorage` に保存し、カスタム関数と、作業ウィンドウなどのアドイン UI 要素の両方が、そのトークンにアクセスできるようにすることができます。 同様に、2 つのアドインが同じドメインを共有している場合 (例: www.contoso.com/addin1、www.contoso.com/addin2)、アドイン間で `AsyncStorage` を介して情報を共有できるようにすることができます。 サブドメインが異なるアドインについては (例: subdomain.contoso.com/addin1、differentsubdomain.contoso.com/addin2)、`AsyncStorage` インスタンスも別々となることに留意してください。 

`AsyncStorage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。

`AsyncStorage` オブジェクトでは、以下の方法が利用可能です。
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`: すべての情報をクリアする方法 (`clear` など) は実装されていません。 代わりに、一度に複数のエントリを削除できる `multiRemove` を使用してください。

### <a name="asyncstorage-example"></a>AsyncStorage の使用例 

以下のコード サンプルでは、`AsyncStorage.getItem` 関数を呼び出してストレージから値を取得します。

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

## <a name="displaying-a-dialog-box"></a>ダイアログ ボックスの表示

カスタム関数 (またはアドインの他の部分) 内で、`OfficeRuntime.displayWebDialogOptions` API を使用してダイアログ ボックスを表示することができます。 このダイアログ API は、作業ウィンドウとアドイン コマンド内では使用可能であるが、カスタム関数内では使用できない[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) の代わりに、使用できます。

### <a name="dialog-api-example"></a>ダイアログ API の使用例

以下のコード サンプルでは、関数 `getTokenViaDialog` がダイアログ API の `displayWebDialogOptions` 関数を使用して、ダイアログ ボックスを表示します。

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

複数のプラットフォーム (Office アドインのキー テナントの 1 つ) で実行するアドインを作成するには、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用したりしないでください。 カスタム関数が JavaScript ランタイムを使用する Excel for Windows では、カスタム関数は DOM にアクセスできません。

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](excel-tutorial-custom-functions.md)
