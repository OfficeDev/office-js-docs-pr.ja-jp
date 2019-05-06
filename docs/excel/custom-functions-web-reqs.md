---
ms.date: 04/20/2019
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: Web 要求とその他のデータがカスタム関数(プレビュー)を処理します
localization_priority: Priority
ms.openlocfilehash: 2942ec56e46d6eb586b516eedab17c1eeb98d9c8
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353266"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a>カスタム関数によるデータの受信と処理

カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) など workbook以外からのデータの受信です。 カスタム関数はXHRを通してデータを要求し、同時に要求を fetch したりデータをストリーミングする事ができます。

次のドキュメンテーションはweb 要求のいくつかの例を説明していますが、ストリーミング機能を構築するには、[カスタム関数 チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)を参照してください。

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返します。
2. コールバック関数を使用して Promise を最終値で解決します。

[`Fetch`](https://developer.mozilla.org/ja-JP/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/ja-JP/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。

カスタム関数のランタイムは、[同送信元ポリシー](https://developer.mozilla.org/ja-JP/docs/Web/Security/Same-origin_policy)とシンプルな [CORS](https://www.w3.org/TR/cors/) を要求することにより、XHR が追加のセキュリティ対策を実装します。

単純な CORS 実装は cookies を使用できず、簡単なメソッド(GET、 HEAD、 POST) のみをサポートすることに注意してください。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 コンテンツ タイプが、 `application/x-www-form-urlencoded`、 `text/plain`、または `multipart/form-data`の単純な CORS のコンテンツ タイプ ヘッダーを使う事もできます。

### <a name="xhr-example"></a>XHR の使用例

以下のコード サンプルでは、**getTemperature**関数が sendWebRequest 関数を呼び出して、温度計 ID に基づく特定の領域の温度を取得します。 sendWebRequest 関数は XHR を使用して、データを提供するエンドポイントを要求する GET リクエストを発行します。

```JavaScript
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

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

コンテキストを使った XHR リクエストのその他のサンプルについては、[Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload)の Github リポジトリの、`getFile` 関数範囲内で[このファイル ](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) を参照ください。

### <a name="fetch-example"></a>Fetch の使用例

以下のコードサンプルでは、stockPriceStream 関数が ストック ティッカー シンボル を使い、1000 ミリ秒ごとに株価を取得します。 このサンプルに関する詳細および JSON については、[カスタム関数 チュートリアル](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)を参照ください。 

```JavaScript
function stockPriceStream(ticker, handler) {
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
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a>WebSocket 経由のデータ受信

カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。 WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベント発生時にサーバーからメッセージを自動的に受信するので、サーバーに明示的にデータ用のポーリングを行う必要がありません。

### <a name="websockets-example"></a>WebSocket の使用例

以下のコード サンプルは、WebSocket 接続を確立し、サーバーからの各受信メッセージを記録します。

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a>ストリーミング関数

ストリーム カスタム関数を使用すると、セルに繰り返しデータを長期的に出力でき、ユーザーが再計算を明示的に要求することは特に必要ありません。 以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、setResult コールバックを使用して自動的に新しい値を表示します。
- 2 番目の入力パラメーター、ハンドラーは、オートコンプリート メニューから変数を選択する場合には Excel のエンドユーザーには表示されません。
- onCanceled コールバックは、関数がキャンセルされた場合に実行される関数を定義します。 すべてのストリーム関数には、このようなキャンセル ハンドラーの実装が必要です。 詳細については、「[関数をキャンセルする](#canceling-a-function)」を参照してください。

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

JSON メタデータ ファイルでストリーミング関数のメタデータを指定する場合は、関数のスクリプト ファイル内の `@streaming` JSDOC コメント タグを使用してこれを自動生成できます。 詳しくは、[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)をご覧ください。

## <a name="canceling-a-function"></a>関数をキャンセルする

状況によっては、帯域幅の消費量、作業メモリ、UPC への負荷を軽減するために、ストリーム カスタム関数の実行をキャンセルする必要があります。 Excel では、次のような状況で関数の実行をキャンセルします。

- ユーザーが、関数を参照するセルを編集または削除した場合。
- 関数の引数 (入力) の 1 つが変更されたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。
- ユーザーが手動で再計算をトリガーしたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

関数をキャンセル可能にするには、関数コードのハンドラーを実装し、キャンセルされたときの対応を指示します。 または、関数のスクリプト ファイル内の `@cancelable` JSDOC コメント タグを使用します。 詳しくは、[カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)をご覧ください。

## <a name="see-also"></a>関連項目

* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
