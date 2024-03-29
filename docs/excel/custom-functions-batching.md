---
ms.date: 09/09/2022
description: バッチ処理カスタム関数を組み合わせてリモート サービスへのネットワーク呼び出しを減らします。
title: リモート サービスのためのバッチ処理カスタム関数の呼び出し
ms.localizationpriority: medium
ms.openlocfilehash: f779351789350bbc591b1b5d7a975ff9f70cda26
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234922"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>リモート サービスの Batch カスタム関数呼び出し

カスタム関数がリモート サービスを呼び出す場合は、リモート サービスへのネットワークの呼び出し数を減らすバッチ処理のパターンを使用できます。 バッチ処理をしたネットワーク ラウンド トリップのウェブ サービスへのすべての呼び出しを、1 回に減らします。 これは、ワークシートが再計算するときに最適な方法です。

たとえば、別のユーザーがスプレッドシートの 100 セル内でカスタム関数を使用し、スプレッドシートを再計算した場合、カスタム関数は 100 回実行され、100 回ネットワークの呼び出しを行います。 バッチ処理のパターンを使用すると、1 つのネットワークの呼び出しで 100 の計算すべてを結合することができます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>完成したサンプルを表示する

完成したサンプルを表示するには、この記事に従って、独自のプロジェクトにコード例を貼り付けます。 たとえば、TypeScript 用の新しいカスタム関数プロジェクトを作成するには、 [Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用し、この記事のすべてのコードをプロジェクトに追加します。 コードを実行して試してください。

または、 [カスタム関数のバッチ処理パターン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)で完全なサンプル プロジェクトをダウンロードまたは表示します。 読み進める前に全体のコードを表示したい場合、 [スクリプト ファイル](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js)をご覧ください。

## <a name="create-the-batching-pattern-in-this-article"></a>この記事内でバッチ処理パターンを作成する

カスタム関数にバッチ処理を設定するには、次の 3 つの主要なセクションのコードを記述する必要があります。

1. Excel がカスタム関数を呼び出すたびに、呼び出しのバッチに新しい操作を追加する [プッシュ操作](#add-the-_pushoperation-function) 。
2. バッチの準備ができたら [リモート要求を行う関数](#make-the-remote-request) 。
3. [バッチ要求に応答し](#process-the-batch-call-on-the-remote-service)、すべての操作結果を計算し、値を返すサーバー コード。

次のセクションでは、コードを一度に 1 つの例で作成する方法について説明します。 [Office アドイン ジェネレーターの Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用して、まったく新しいカスタム関数プロジェクトを作成することをお勧めします。 新しいプロジェクトを作成するには、「 [Excel カスタム関数の開発を開始](../quickstarts/excel-custom-functions-quickstart.md)する」を参照してください。 TypeScript または JavaScript を使用できます。

## <a name="batch-each-call-to-your-custom-function"></a>カスタム関数の各呼び出しにバッチ処理をする

操作を実行するリモート サービスの呼び出し機能を使ってカスタム関数の演算を実行し、必要な結果を計算します。 要求された各操作をバッチ内に保存する方法を提供します。 後で、その操作にバッチ処理をする `_pushOperation`関数を作成する方法が表示されます。 最初に、カスタム関数から`_pushOperation`を呼び出す方法については、次のコード例をみてください。

次のコードでは、カスタム関数は除算を実行しますが、実際の計算を実行するにはリモート サービスに依存しています。 リモート サービスにその操作と別の操作を一緒にバッチ処理し、`_pushOperation`を呼び出します。 その名称は **div2** 操作といいます。 リモート サービスが同じスキーム (詳細については、この後のリモート サービスで) を使用する限り、任意の名前付けスキームを操作に使用することができます。 また、操作を実行する必要があるリモートサービスの引数が渡されます。

### <a name="add-the-div2-custom-function"></a>div2 カスタム関数を追加する

**functions.js** または **functions.ts** ファイルに次のコードを追加します (JavaScript または TypeScript を使用したかどうかによって異なります)。

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### <a name="add-global-variables-for-tracking-batch-requests"></a>バッチ要求を追跡するためのグローバル変数を追加する

次に、 **functions.js** または **functions.ts** ファイルに 2 つのグローバル変数を追加します。 `_isBatchedRequestScheduled` は、後でリモート サービスへのバッチ呼び出しのタイミングを設定するために重要です。

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### <a name="add-the-_pushoperation-function"></a>関数を追加する`_pushOperation`

Excel がカスタム関数を呼び出すときは、バッチ配列に操作をプッシュする必要があります。 次の **_pushOperation** 関数コードは、カスタム関数から新しい操作を追加する方法を示しています。 新しいバッチ エントリを作成し、処理を解決または拒否するための新しい promise を作成し、そしてバッチ配列にエントリをプッシュします。

このコードは、バッチがスケジュールされているかどうかも確認します。 この例では、それぞれのバッチはすべて100 ミリ秒ごとに実行するようスケジュールされています。 必要に応じて、この値を調整することができます。 高い値は、リモート サービスに送信される大きなバッチで発生し、ユーザーが結果を確認するまでの応答時間が長くなります。 小さい値は、より多くのバッチがリモート サービスに送信されますが、ユーザーの応答時間は短くなる傾向があります。

この関数は、実行する操作の文字列名を含む **invocationEntry** オブジェクトを作成します。 たとえば、 `multiply` と `divide`という名前の 2 つのカスタム関数がある場合、バッチのエントリ内で操作名として再利用できます。 `args` は、Excel からカスタム関数に渡された引数を保持します。 最後に、または`reject`メソッドは、`resolve`リモート サービスが返す情報を保持する Promise を格納します。

**functions.js** または **functions.ts** ファイルに次のコードを追加します。

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>リモートの要求を行う

`_makeRemoteRequest`関数の目的は、操作のバッチをリモート サービスに渡し、それから各カスタム関数に結果を返します。 まず、バッチ配列のコピーを作成します。 これにより、concurrent カスタム関数は、Excel からすぐに新しい配列にバッチ処理を呼び出すことができます。 そのコピーは、それから promise 情報が含まれていない単純な配列になります。 機能しない場合は、リモート サービスにその promise を渡しても意味をなしません。 リモート サービスが何を返すかによって、`_makeRemoteRequest` は拒否するか、またはそれぞれの promise を解決します。

**functions.js** または **functions.ts** ファイルに次のコードを追加します。

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>独自のソリューションに`_makeRemoteRequest`を変更します。

`_makeRemoteRequest`関数は、あとで表示されますが、リモート サービスを表すモックの`_fetchFromRemoteService`を呼び出します。 これにより、簡単に学習でき、この記事でコードを実行することができます。 ただし、実際のリモート サービスにこのコードを使用する場合は、次の変更を行う必要があります。

- ネットワーク経由でバッチ処理をシリアル化する方法を決定します。 たとえば、JSON の本文に、配列を配置することがあります。
- `_fetchFromRemoteService`を呼び出す代わりに、バッチ処理を渡すリモート サービスに実際にネットワークの呼び出しをする必要があります。

## <a name="process-the-batch-call-on-the-remote-service"></a>リモート サービスでバッチの呼び出しを処理します。

最後の手順では、リモート サービスでバッチの呼び出しを処理をします。 つぎのコード サンプルは、`_fetchFromRemoteService`関数を表しています。 この関数は、それぞれの操作を展開せずに指定した操作を実行し、それから結果を返します。 この記事の学習の目的は、 `_fetchFromRemoteService`関数がリモート サービスを web アドインで実行し、リモート サービスをモックするように設計されています。 このコードを **functions.js** または **functions.ts ファイルに** 追加すると、実際のリモート サービスを設定しなくても、この記事のすべてのコードを調査して実行できます。

**functions.js** または **functions.ts** ファイルに次のコードを追加します。

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>`_fetchFromRemoteService`をライブ リモート サービスに変更する

ライブ リモート サービスで実行するように関数を変更 `_fetchFromRemoteService` するには、次の変更を行います。

- サーバー プラットフォーム (Node.js またはその他) のマップによっては、クライアント ネットワークがこの関数を呼び出します。 
- モックの一部としてネットワークの遅延をシミュレートする`pause`関数を削除する。
- パラメーターがネットワーク用に変更された場合、渡されたパラメーターで動作する関数の宣言を変更します。 たとえば、配列の代わりに、JSON 本体のバッチ処理で処理をします。
- 操作を実行する関数を変更する (または、操作を実行する関数を呼び出す)。
- 適切な認証機構を適用する。 適切な呼び出し元のみが関数にアクセスできることを確認します。
- リモート サービスで、コードを配置します。

## <a name="next-steps"></a>次の手順

カスタム関数で使用できる[さまざまなパラメーター](custom-functions-parameter-options.md)について確認してください。 または、[カスタム関数で Web 通話](custom-functions-web-reqs.md)を発信する際の基本事項を確認してください。

## <a name="see-also"></a>関連項目

- [関数の揮発性の値](custom-functions-volatile.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
