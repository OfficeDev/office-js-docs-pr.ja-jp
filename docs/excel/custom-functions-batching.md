---
ms.date: 07/10/2019
description: バッチ処理カスタム関数を組み合わせてリモート サービスへのネットワーク呼び出しを減らします。
title: リモート サービスのためのバッチ処理カスタム関数の呼び出し
localization_priority: Normal
ms.openlocfilehash: 5e48488b323e53e35698b2f64724b78da6abc599
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718746"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a>リモート サービスのためのバッチ処理カスタム関数の呼び出し

カスタム関数がリモート サービスを呼び出す場合は、リモート サービスへのネットワークの呼び出し数を減らすバッチ処理のパターンを使用できます。 バッチ処理をしたネットワーク ラウンド トリップのウェブ サービスへのすべての呼び出しを、1 回に減らします。 これは、ワークシートが再計算するときに最適な方法です。

たとえば、別のユーザーがスプレッドシートの 100 セル内でカスタム関数を使用し、スプレッドシートを再計算した場合、カスタム関数は 100 回実行され、100 回ネットワークの呼び出しを行います。 バッチ処理のパターンを使用すると、1 つのネットワークの呼び出しで 100 の計算すべてを結合することができます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>完成したサンプルを表示する

この記事を参考にして、自分のプロジェクトにコードの例を貼り付けることができます。 たとえば、[Yo Office ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して TypeScript 用の新しいカスタム関数プロジェクトを作成し、この記事のすべてのコードをそのプロジェクトに追加することができます。 その後、コードを実行して試してください。

[カスタム関数のバッチ処理パターン](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching)で完全なサンプル プロジェクトをダウンロードまたは表示することができます。 読み進める前に全体のコードを表示したい場合、 [スクリプト ファイル](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts)をご覧ください。

## <a name="create-the-batching-pattern-in-this-article"></a>この記事内でバッチ処理パターンを作成する

カスタム関数にバッチ処理を設定するには、次の 3 つの主要なセクションのコードを記述する必要があります。

1. バッチに新しい操作を追加するプッシュ操作の呼び出しのたびに、Excel はカスタム関数を呼び出します。
2. バッチの準備ができたときのリモート要求を行う関数です。
3. バッチ要求に応答するサーバー コードは、すべての操作の結果を計算して値を返します。

次のセクションでは、一度に 1 つのコード例を構築する方法が表示されます。 **functions.ts** ファイルにそれぞれのコード例を追加します。 Yo Office ジェネレーター 使用して、新しいカスタム関数のプロジェクトを作成することをお勧めします。 新しいプロジェクトを作成するには、 [Excel のカスタム関数の開発を開始する](../quickstarts/excel-custom-functions-quickstart.md)を参照し、JavaScript ではなく TypeScript を使用してください。

## <a name="batch-each-call-to-your-custom-function"></a>カスタム関数の各呼び出しにバッチ処理をする

操作を実行するリモート サービスの呼び出し機能を使ってカスタム関数の演算を実行し、必要な結果を計算します。 要求された各操作をバッチ内に保存する方法を提供します。 後で、その操作にバッチ処理をする `_pushOperation`関数を作成する方法が表示されます。 最初に、カスタム関数から`_pushOperation`を呼び出す方法については、次のコード例をみてください。

次のコードでは、カスタム関数は除算を実行しますが、実際の計算を実行するにはリモート サービスに依存しています。 リモート サービスにその操作と別の操作を一緒にバッチ処理し、`_pushOperation`を呼び出します。 その名称は**div2**操作といいます。 リモート サービスが同じスキーム (詳細については、この後のリモート サービスで) を使用する限り、任意の名前付けスキームを操作に使用することができます。 また、操作を実行する必要があるリモートサービスの引数が渡されます。

### <a name="add-the-div2-custom-function-to-functionsts"></a>functions.ts に div2 カスタム関数を追加する

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}
```

次に、1 つのネットワークの呼び出しに渡されるすべての操作が格納されるバッチの配列を定義します。 次のコードでは、配列内で各バッチのエントリを記述するインターフェイスを定義する方法を表示します。 どの文字列名のどの操作を実行するのか、インターフェイスが操作を定義します。 たとえば、 `multiply` と `divide`という名前の 2 つのカスタム関数がある場合、バッチのエントリ内で操作名として再利用できます。 `args` は、Excel からカスタム関数に渡された引数が保持されます。 最後に、`resolve` または `reject`はリモート サービスが返した情報を保持している promise を格納します。

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

次に、前のインターフェイスを使用するバッチの配列を作成します。 バッチが予定されているかどうかを追跡するため、`_isBatchedRequestSchedule` 変数を作成します。 リモート サービスへのバッチの呼び出しのタイミングは、後で重要になります。

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

最後に、Excel がカスタム関数を呼び出すと、バッチ配列への操作をプッシュする必要があります。 次のコードでは、カスタム関数から新しい操作を追加する方法を示します。 新しいバッチ エントリを作成し、処理を解決または拒否するための新しい promise を作成し、そしてバッチ配列にエントリをプッシュします。

このコードは、バッチがスケジュールされているかどうかも確認します。 この例では、それぞれのバッチはすべて100 ミリ秒ごとに実行するようスケジュールされています。 必要に応じて、この値を調整することができます。 高い値は、リモート サービスに送信される大きなバッチで発生し、ユーザーが結果を確認するまでの応答時間が長くなります。 小さい値は、より多くのバッチがリモート サービスに送信されますが、ユーザーの応答時間は短くなる傾向があります。

### <a name="add-the-_pushoperation-function-to-functionsts"></a>functions.ts に `_pushOperation` 関数を追加する

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
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
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>リモートの要求を行う

`_makeRemoteRequest`関数の目的は、操作のバッチをリモート サービスに渡し、それから各カスタム関数に結果を返します。 まず、バッチ配列のコピーを作成します。 これにより、concurrent カスタム関数は、Excel からすぐに新しい配列にバッチ処理を呼び出すことができます。 そのコピーは、それから promise 情報が含まれていない単純な配列になります。 機能しない場合は、リモート サービスにその promise を渡しても意味をなしません。 リモート サービスが何を返すかによって、`_makeRemoteRequest` は拒否するか、またはそれぞれの promise を解決します。

### <a name="add-the-following-_makeremoterequest-method-to-functionsts"></a>次の`_makeRemoteRequest`メソッドを functions.ts に追加します。

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>独自のソリューションに`_makeRemoteRequest`を変更します。

`_makeRemoteRequest`関数は、あとで表示されますが、リモート サービスを表すモックの`_fetchFromRemoteService`を呼び出します。 これにより、簡単に学習でき、この記事でコードを実行することができます。 ただし、実際のリモート サービスでこのコードを使用するときは、次の変更を行う必要があります:

- ネットワーク経由でバッチ処理をシリアル化する方法を決定します。 たとえば、JSON の本文に、配列を配置することがあります。
- `_fetchFromRemoteService`を呼び出す代わりに、バッチ処理を渡すリモート サービスに実際にネットワークの呼び出しをする必要があります。

## <a name="process-the-batch-call-on-the-remote-service"></a>リモート サービスでバッチの呼び出しを処理します。

最後の手順では、リモート サービスでバッチの呼び出しを処理をします。 つぎのコード サンプルは、`_fetchFromRemoteService`関数を表しています。 この関数は、それぞれの操作を展開せずに指定した操作を実行し、それから結果を返します。 この記事の学習の目的は、 `_fetchFromRemoteService`関数がリモート サービスを web アドインで実行し、リモート サービスをモックするように設計されています。 **functions.ts** ファイルにこのコードを追加することができ、実際のリモート サービスを設定しなくても、この記事内のすべてのコードを学習し実行することができます。

### <a name="add-the-following-_fetchfromremoteservice-function-to-functionsts"></a>次の `_fetchFromRemoteService` 関数を functions.ts に追加します。

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
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

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>`_fetchFromRemoteService`をライブ リモート サービスに変更する

ライブ リモート サービスで実行する`_fetchFromRemoteService`関数を変更するには、次の変更を行います:

- サーバー プラットフォーム (Node.js またはその他) のマップによっては、クライアント ネットワークがこの関数を呼び出します。 
- モックの一部としてネットワークの遅延をシミュレートする`pause`関数を削除する。
- パラメーターがネットワーク用に変更された場合、渡されたパラメーターで動作する関数の宣言を変更します。 たとえば、配列の代わりに、JSON 本体のバッチ処理で処理をします。
- 操作を実行する関数を変更する (または、操作を実行する関数を呼び出す)。
- 適切な認証機構を適用する。 適切な呼び出し元のみが関数にアクセスできることを確認します。
- リモート サービスで、コードを配置します。

## <a name="next-steps"></a>次の手順
カスタム関数で使用できる[さまざまなパラメーター](custom-functions-parameter-options.md)について確認してください。 または、[カスタム関数で Web 通話](custom-functions-web-reqs.md)を発信する際の基本事項を確認してください。

## <a name="see-also"></a>関連項目

* [関数の揮発性の値](custom-functions-volatile.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
