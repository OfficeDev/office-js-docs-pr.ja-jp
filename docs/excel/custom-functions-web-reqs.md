---
ms.date: 05/02/2022
description: Excel のカスタム関数を使用して、ブックへの外部データのストリーミングを要求、ストリーミング、キャンセルします。
title: カスタム関数でデータを受信して​​処理する
ms.localizationpriority: medium
ms.openlocfilehash: fbe319e79d4cded5fe4b37ce5a654e633996f22a
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958546"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>カスタム関数でデータを受信して​​処理する

カスタム関数が Excel の機能を向上させる方法の 1 つは、Web やサーバーなどのブック以外の場所 ( [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) を介して) からデータを受信することです。 [`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API)などの API や、サーバーとの情報のやりとりを要求する HTTP を発行する標準 ウェブ API である `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)を使って外部データを要求することができます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![API から時刻をストリーミングするカスタム関数の GIF。](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>外部ソースからデータを返す関数

カスタム関数が外部ソースからデータを取得する場合には、以下のことを実行する必要があります。

1. Excel に [JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) を返します。
2. コールバック関数を `Promise` 使用して、最終的な値を使用して解決します。

### <a name="fetch-example"></a>Fetch の使用例

次のコード サンプルでは、この関数は `webRequest` 、国際宇宙ステーションに現在いるユーザーの数を追跡する架空の外部 API に到達します。 この関数は JavaScript `Promise` を返し、仮想 API から情報を要求するために使用 `fetch` します。 結果のデータは JSON に変換され、 `names` プロパティは文字列に変換されます。これは、Promise を解決するために使用されます。

独自の機能を開発するときに、Web 要求が時間内に完了しない場合は、アクションを実行するか、[複数の API 要求をバッチ処理すること](custom-functions-batching.md)を検討してください。

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
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

> [!NOTE]
> `fetch`を使用すると、コールバックのネストが回避され、場合によっては XHR に適している場合があります。

### <a name="xhr-example"></a>XHR の使用例

次のコード サンプルでは、この関数は `getStarCount` Github API を呼び出して、特定のユーザーのリポジトリに与えられた星の量を検出します。 これは、JavaScript `Promise`を返す非同期関数です。 Web 呼び出しからデータが取得されると、データをセルに返す Promise が解決されます。

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

ストリーミング関数を宣言するには、次の 2 つのオプションのいずれかを使用します。

- `@streaming`タグ。
- `CustomFunctions.StreamingInvocation`呼び出しパラメーター。

以下のコード サンプルは、毎秒ごとに結果に数値を追加するカスタム関数です。 このコードについては、次の点に注意してください。

- Excel は、`setResult` メソッドを使用して自動的に新しい値を表示します。
- 2 番目の入力パラメーターの `invocation` は、[オートコンプリート] メニューから関数が選択された場合、Excel のエンドユーザーに表示されません。
- コールバックは `onCanceled` 、関数が取り消されたときに実行される関数を定義します。
- ストリーミングは必ずしも Web 要求の作成に関連しているとは限りません。 この場合、関数は Web 要求を行っていませんが、設定された間隔でデータを取得しているため、ストリーミング `invocation` パラメーターを使用する必要があります。

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

## <a name="cancel-a-function"></a>関数を取り消す

Excel は、次の状況で関数の実行を取り消します。

- ユーザーが、関数を参照するセルを編集または削除した場合。
- 関数の引数 (入力) の 1 つが変更されたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。
- ユーザーが手動で再計算をトリガーしたとき。 この場合、キャンセルに続いて、関数の新しい呼び出しがトリガーされます。

また、要求が発生したときに、オフラインの場合でも、ケースを処理する既定のストリーミング値を設定することもできます。

> [!NOTE]
> 取り消し可能な関数と呼ばれる関数のカテゴリもあります。これらはストリーミング関数とは関係 _ありません_ 。 1 つの値を返す非同期カスタム関数のみが取り消し可能です。 キャンセル可能な関数を使用すると、Web 要求を要求中に終了させることができます。キャンセルするときの処理を決定するには、[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)を使用します。 タグ `@cancelable` を使用して、キャンセル可能な関数を宣言します。

### <a name="use-an-invocation-parameter"></a>呼び出しパラメーターを使用する

`invocation` パラメーターは、既定ではカスタム関数の最後のパラメーターです。 このパラメーターは`invocation`、セルに関するコンテキスト (アドレスや内容など) を提供し、メソッドと`onCanceled`イベントを`setResult`使用して、関数がストリーム (`setResult`) または取り消された`onCanceled` () ときに行う処理を定義できます。

TypeScript を使用している場合、呼び出しハンドラーは型 [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) または [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).

## <a name="receiving-data-via-websockets"></a>WebSocket を使用したデータ受信

カスタム関数内で、[WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) を使用して、サーバーとの固定接続経由でデータを交換することができます。 WebSocket を使用すると、カスタム関数はサーバーとの接続を開き、特定のイベントが発生したときにサーバーからメッセージを自動的に受信できます。データをサーバーに明示的にポーリングする必要はありません。

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
- [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
