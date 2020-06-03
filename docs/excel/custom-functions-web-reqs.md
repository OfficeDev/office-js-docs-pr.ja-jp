---
ms.date: 04/29/2020
description: Excel でのカスタム関数を使って外部データを workbook にストリーミング要求したりキャンセルしたりします
title: カスタム関数でデータを受信して​​処理する
localization_priority: Normal
ms.openlocfilehash: 1ae1baa912c914c3a508f1bbf6bd5d9fa6044f7b
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44275747"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>カスタム関数でデータを受信して​​処理する

カスタム関数によって Excel の機能を強化する方法の一つは、ウェブやサーバー (WebSockets 経由) などブック以外からのデータの受信です。 [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![API から時刻をストリームしているカスタム関数の GIF](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. JavaScript Promise を Excel に返します。
2. コールバック関数を使用して Promise を最終値で解決します。

### <a name="fetch-example"></a>Fetch の使用例

次のコードサンプルでは、関数は、" `webRequest` スペースがあります" という名前の "API" に到達します。これは、現在、国際宇宙ステーションにいるユーザーの数を追跡します。 この関数は JavaScript Promise を返し、fetchを使って API から情報を要求します。 結果のデータは JSON に変換され、`names`プロパティは Promise を解決するために使用される文字列に変換されます。

独自の機能を開発するときに、Web 要求が時間内に完了しない場合は、アクションを実行するか、[複数の API 要求をバッチ処理すること](./custom-functions-batching.md)を検討してください。

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
>`Fetch`を使用すると、コールバックのネストが回避され、場合によっては XHR に適している場合があります。

### <a name="xhr-example"></a>XHR の使用例

次のコードサンプルでは、この `getStarCount` 関数は GITHUB API を呼び出して、特定のユーザーのリポジトリに与えられた星の量を検出します。 これは JavaScript Promise を返す非同期関数です。 データが web 呼び出しから取得されると、Promise が解決され、データがセルに返されます。

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

## <a name="make-a-streaming-function"></a>ストリーミング関数を作成する

ストリーム カスタム関数を使用すると、繰り返し更新されるセルにデータを出力でき、ユーザーが明示的に何かを更新する必要ありません。 これは、[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)の関数のように、サービス オンラインのライブ データを確認する際に便利です。

ストリーミング関数を宣言するには、次のいずれかを使用できます。

- `@streaming`タグ。
- `CustomFunctions.StreamingInvocation`呼び出しパラメーター。

以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。
- 2 番目の入力パラメーター、起動は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。
- `onCanceled` コールバックは、関数がキャンセルされた場合に実行される関数を定義します。
- ストリーミングは必ずしもWeb 要求の実行に結び付けられているわけではありません。この場合、関数は Web 要求を行うのではなく、設定された間隔でデータを取得しているため、ストリーミング `invocation` パラメータを使用する必要があります。

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

## <a name="canceling-a-function"></a>関数をキャンセルする

Excel では、次のような状況で関数の実行をキャンセルします。

- ユーザーが、関数を参照するセルを編集または削除した場合。
- 関数の引数 (入力) の 1 つが変更されたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。
- ユーザーが手動で再計算をトリガーしたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

また、要求が発生したときに、オフラインの場合でも、ケースを処理する既定のストリーミング値を設定することもできます。

また、ストリーミング関数と関連の_ない_、キャンセル可能な関数と呼ばれる関数のカテゴリもあります。 1つの値を返す非同期のカスタム関数のみが取り消し可能です。 キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。 タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。

### <a name="using-an-invocation-parameter"></a>起動パラメーターの使用

`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。 パラメーターには、 `invocation` セルに関するコンテキスト (アドレスやコンテンツなど) が指定され、メソッドを使用できるように `setResult` `onCanceled` なります。 これらのメソッドでは、関数がストリーミング (`setResult`) またはキャンセルされた (`onCanceled`) 場合に、関数が何を実行するかを定義します。

TypeScript を使用している場合は、呼び出しハンドラーの型がまたはである必要があり [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) ます。

## <a name="receiving-data-via-websockets"></a>WebSocket 経由のデータ受信

カスタム関数内で、WebSocket を使用してサーバーとの固定接続でデータを交換することができます。 Websocket を使用すると、カスタム関数はサーバーとの接続を開いて、特定のイベントが発生したときに、データに対して明示的にサーバーをポーリングすることなく、サーバーからのメッセージを自動的に受信することができます。

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

## <a name="next-steps"></a>次の手順

- [関数で使用できるさまざまなパラメーターのタイプ](custom-functions-parameter-options.md)についての詳細。
- [複数の API の呼び出しをバッチする](custom-functions-batching.md)方法を探す。

## <a name="see-also"></a>関連項目

- [関数の揮発性の値](custom-functions-volatile.md)
- [カスタム関数の JSON メタデータを作成する](custom-functions-json-autogeneration.md)
- [カスタム関数のメタデータ](custom-functions-json.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
