---
title: Office アドインを初期化する
description: Office アドインを初期化する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 8310c5efb803391f7f0d4b01fda70dc0df537b21
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608140"
---
# <a name="initialize-your-office-add-in"></a>Office アドインを初期化する

Office アドインには、次のような処理を行うスタートアップ ロジックがよくあります。

- ユーザーのバージョンの Office で、コードが呼び出すすべての Office Api をサポートしていることを確認してください。

- 特定の名前のワークシートなど、特定の成果物が存在することを確認します。

- Excel でセルを選択するようにユーザーに求め、選択した値で初期化されたグラフを挿入します。

- バインディングを確立します。

- Office ダイアログ API を使用して、既定のアドイン設定値をユーザーに確認します。

ただし、Office アドインは、ライブラリが読み込まれるまでは、Office JavaScript Api を正常に呼び出せません。 この記事では、ライブラリが読み込まれていることをコードが確認する2つの方法について説明します。

- を使用して初期化 `Office.onReady()` します。
- を使用して初期化 `Office.initialize` します。

> [!TIP]
> `Office.initialize` の代わりに `Office.onReady()` を使用することをお勧めします。 `Office.initialize`はまだサポートされていますが、 `Office.onReady()` より柔軟な機能を提供します。 割り当てることができるハンドラーは1つだけ `Office.initialize` で、Office のインフラストラクチャによって一度だけ呼び出されます。 `Office.onReady()`コード内の別の場所で呼び出し、さまざまなコールバックを使用できます。
> 
> これらの手法の違いの詳細については、「[Office.initialize と Office.onReady の間の主な相違点](#major-differences-between-officeinitialize-and-officeonready)」を参照してください。

アドインの初期化時のイベントのシーケンスの詳細については、「[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)」を参照してください。

## <a name="initialize-with-officeonready"></a>Office.onReady() を使用した初期化

`Office.onReady()`は、Office .js ライブラリが読み込まれているかどうかを確認するときに、 [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)オブジェクトを返す非同期メソッドです。 ライブラリが読み込まれるとき (に限り)、Office ホスト アプリケーションを `Office.HostType` 列挙値 (`Excel`、`Word` など)、およびプラットフォームを `Office.PlatformType` 列挙値 (`PC`、`Mac`、`OfficeOnline` など) で指定するオブジェクトとして Promise を解決します。 `Office.onReady()` を呼び出すときにライブラリが既に読み込まれている場合、Promise をすぐに解決します。

`Office.onReady()` を呼び出す方法の 1 つは、コールバック メソッドを渡すことです。 次に例を示します。

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

また、コールバックを渡す代わりに、`then()` メソッドを `Office.onReady()` の呼び出しにチェーン接続することもできます。 たとえば、次のコードで、ユーザーのバージョンの Excel が、アドインで呼び出す可能性があるすべての API をサポートしているかを確認します。

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

`async` と `await` キーワードを TypeScript で使用する同じ例を次に示します。

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、*通常*、そのようなフレームワークは `Office.onReady()` への応答内に配置される必要があります。 たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

ただし、この実習には例外があります。 たとえば、ブラウザーのツールを使用してご使用の UI をデバッグするため、(Office ホスト内にサイドロードする代わりに) ブラウザーでご利用のアドインを開く必要があるとします。 Office.js がブラウザーに読み込まれないため、`onReady` は実行できず、Office `onReady` 内に呼び出される場合は、`$(document).ready` は実行されません。 

アドインの読み込み中に作業ウィンドウに進行状況のインジケーターが表示されるようにする場合は、別の例外があります。 このシナリオでは、コードで jQuery を呼び出し、コールバックを使用して進行状況インジケーターをレンダリングする必要があり `ready` ます。 その後、Office `onReady` のコールバックで、進行状況のインジケーターを最終的な UI に置き換えることができます。 

## <a name="initialize-with-officeinitialize"></a>Office.initialize を使用した初期化

Office.js ライブラリが読み込まれ、ユーザーとの対話の準備が完了すると、初期化イベントが発生します。 初期化ロジックを実装する `Office.initialize` にハンドラーを割り当てることができます。 ユーザーのバージョンの Excel が、アドインで呼び出す可能性があるすべての API をサポートしているかを確認する例は、次のとおりです。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合は、*通常*、これらはイベント内に配置する必要があり `Office.initialize` ます (前の手順では、「 **Office. onready ()** セクションでの初期化」で説明されている例外)。 たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

作業ウィンドウ アドインとコンテンツ アドインの場合、`Office.initialize` で追加の _reason_ パラメーターが提供されます。 このパラメーターでは、アドインがどのように現在のドキュメントに追加されたかが示されます。 これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに、別のロジックを提供するために使用できます。

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

詳細については、[Office.initialize イベント](/javascript/api/office)に関するページ、および [InitializationReason 列挙型](/javascript/api/office/office.initializationreason)に関するページを参照してください。

## <a name="major-differences-between-officeinitialize-and-officeonready"></a>Office.initialize と Office.onReady の間の主な相違点

- `Office.initialize` にハンドラーは 1 つだけ割り当てることができ、1 回だけは、Office のインフラストラクチャで呼び出されますが、`Office.onReady()` の呼び出しはコードと異なる場所にして、異なるコールバックを使用します。 たとえば、ご利用のコードでは、カスタム スクリプトが初期化ロジックを実行するコールバックを読み込むとすぐに `Office.onReady()` を呼び出しますが、ご利用のコードには、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出す、ボタンを作業ウィンドウに含めることもできます。 その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。

- `Office.initialize` イベントは、Office.js 自体が初期化される内部プロセスの最後に発生します。 内部のプロセスが終了した後、*すぐに*発生します。 イベントにハンドラーを割り当てるコードが、イベント発生後に長時間実行される場合、ハンドラーは実行されません。 たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読み込まれた後で、カスタム JavaScript を読み込む前に、ポリフィルのファイルを読み込むためのアドインのホーム ページを構成する場合があります。 ご使用のスクリプトでハンドラーの読み込みと割り当てが行われる時点で、初期化イベントは既に発生しています。 ですが、`Office.onReady()` を呼び出すのに "遅すぎる" ことは決してありません。 初期化イベントが既に発生している場合、コールバックがすぐに実行されます。

> [!NOTE]
> スタートアップ ロジックがない場合でも、アドイン JavaScript を読み込むときには、`Office.onReady()` を呼び出すか、または空の関数を `Office.initialize` に割り当てる必要があります。 Office ホストとプラットフォームの組み合わせによっては、これらのいずれかが発生するまでは作業ウィンドウが読み込まれないことがあります。 次の例はこの 2 つの方法を示しています。
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)