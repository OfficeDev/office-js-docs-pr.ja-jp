---
ms.date: 05/30/2019
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Priority
ms.openlocfilehash: 22f79c8b4e7e39569d3b955477e9397a053e1a8f
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910337"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>カスタム関数でデータを受信して​​処理する

カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。 カスタム関数は XHR を通してデータを要求し、同時に要求を `fetch` したりデータをストリーミングしたりすることができます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

次のドキュメンテーションはweb 要求のいくつかの例を説明していますが、ストリーミング機能を構築するには、[カスタム関数 チュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)を参照してください。

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返します。
2. コールバック関数を使用して Promise を最終値で解決します。

[`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。

カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。

単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。

### <a name="xhr-example"></a>XHR の使用例

以下のコード サンプルでは、**getTemperature**関数が sendWebRequest 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。 sendWebRequest 関数は XHR を使用して、データを提供するエンドポイントを要求する GET リクエストを発行します。

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
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

コンテキストを使った XHR リクエストのその他のサンプルについては、[Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload)の Github リポジトリの、`getFile` 関数範囲内で[このファイル ](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) を参照ください。

### <a name="fetch-example"></a>Fetch の使用例

以下のコード サンプルでは、`stockPriceStream` 関数がストック ティッカー シンボルを使い、1000 ミリ秒ごとに株価を取得します。 このサンプルに関する詳細については、[カスタム関数チュートリアル](../tutorials/excel-tutorial-create-custom-functions.md#create-a-streaming-asynchronous-custom-function)を参照してください。

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
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
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a>WebSocket 経由のデータ受信

カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。 WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。

### <a name="websockets-example"></a>WebSocket の使用例

以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a>ストリーミング関数を作成する

ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新する必要ありません。 これは、[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)の関数のように、サービス オンラインのライブ データを確認する際に便利です。

ストリーミング関数を宣言するには、JSDoc コメント タグ `@stream` を使用します。 新しい情報に基づいて関数が再評価する可能性があることをユーザーに警告するには、関数の名前または説明にこれを示すことができるストリームまたはその他の文言を使用することをお勧めします。

次の例では、指定した量だけ毎秒指定した数値を増加させるストリーミング関数を示しています。

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
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> また、ストリーミング関数と関連の*ない*、キャンセル可能な関数と呼ばれる関数のカテゴリもあります。 以前のバージョンのカスタム関数は、手動で記述された JSON で `"cancelable": true` と `"streaming": true` を宣言する必要がありました。 自動生成されたメタデータの導入以来、1 つの値を返す非同期のカスタム関数のみがキャンセル可能です。 キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。 タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。

### <a name="using-an-invocation-parameter"></a>起動パラメーターの使用

`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。 `invocation` パラメーターは、セルに関するコンテキスト (アドレスなど) を提供し、`setResult` メソッドや `onCanceled` メソッドを使用することもできます。 これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。

TypeScript を使用している場合は、呼び出しハンドラーは `CustomFunctions.StreamingInvocation` 型または `CustomFunctions.CancelableInvocation` 型である必要があります。

### <a name="streaming-and-cancelable-function-example"></a>ストリーム関数とキャンセル可能な関数の例
以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。
- 2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。
- `onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> Excel では、次のような状況で関数の実行をキャンセルします。
>
> - ユーザーが、関数を参照するセルを編集または削除した場合。
> - 関数の引数 (入力) の 1 つが変更されたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。
> - ユーザーが手動で再計算をトリガーしたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

## <a name="next-steps"></a>次の手順

* [関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。
* [複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。

## <a name="see-also"></a>関連項目

* [関数の揮発性の値](custom-functions-volatile.md)
* [カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
