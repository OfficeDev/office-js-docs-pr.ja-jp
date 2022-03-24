---
title: Office アドインを初期化する
description: アドインを初期化するOfficeする方法について学習します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: afc8b5fa497c517c943708837829636e2867c9f7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743432"
---
# <a name="initialize-your-office-add-in"></a>Office アドインを初期化する

Office アドインには、次のような処理を行うスタートアップ ロジックがよくあります。

- ユーザーのバージョンの API が、コードOffice呼び出Office API をサポートしている必要があります。

- 特定の名前を持つワークシートなど、特定の成果物が存在する必要があります。

- ユーザーにセルを選択するように求めるメッセージが表示Excel、選択した値で初期化されたグラフを挿入します。

- バインディングを確立します。

- ユーザーに既定Office設定の値を求めるプロンプトを表示するには、ダイアログ API を使用します。

ただし、Officeが読み込まれるまで、JavaScript API Office呼び出しに成功することはできません。 この記事では、ライブラリが読み込まれた場合にコードが確実に読み込まれる 2 つの方法について説明します。

- で初期化します `Office.onReady()`。
- で初期化します `Office.initialize`。

> [!TIP]
> `Office.initialize` の代わりに `Office.onReady()` を使用することをお勧めします。 まだ `Office.initialize` サポートされているが、柔軟性 `Office.onReady()` が高い。 ハンドラーは 1 つのみ割り当て`Office.initialize`、そのハンドラーは 1 回だけ呼び出され、Officeできます。 コード内の異 `Office.onReady()` なる場所で呼び出し、異なるコールバックを使用できます。
> 
> これらの手法の違いの詳細については、「[Office.initialize と Office.onReady の間の主な相違点](#major-differences-between-officeinitialize-and-officeonready)」を参照してください。

アドインの初期化時のイベントのシーケンスの詳細については、「[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)」を参照してください。

## <a name="initialize-with-officeonready"></a>Office.onReady() を使用した初期化

`Office.onReady()`は、Promise オブジェクトを返す非同期[](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)メソッドで、ライブラリが読み込まれているOffice.js確認します。 ライブラリが読み込まれると、Promise は、列挙値 (など) を持つ Office `Office.HostType` クライアント アプリケーションと、列挙値 (`Mac``PC``Word``Excel``OfficeOnline`、、など) `Office.PlatformType` を持つプラットフォームを指定するオブジェクトとして解決します。 `Office.onReady()` を呼び出すときにライブラリが既に読み込まれている場合、Promise をすぐに解決します。

呼び出す方法の `Office.onReady()` 1 つは、コールバック メソッドを渡す方法です。例を次に示します。

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

TypeScript で and キーワードを使用する `async` のと `await` 同じ例を次に示します。

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

独自の初期化ハンドラーまたはテストを含む追加の JavaScript フレームワークを使用している場合は、通常、 への応答内に配置する必要があります`Office.onReady()`。たとえば、[JQuery の関数は](https://jquery.com)`$(document).ready()`次のように参照されます。

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

ただし、この実習には例外があります。 たとえば、ブラウザー ツールを使用して UI をデバッグするために、ブラウザーでアドインを (Office アプリケーションでサイドロードする代わりに) ブラウザーで開きたいとします。 Office.js がブラウザーに読み込まれないため、`onReady` は実行できず、Office `onReady` 内に呼び出される場合は、`$(document).ready` は実行されません。 

もう 1 つの例外は、アドインの読み込み中に作業ウィンドウに進行状況インジケーターを表示する場合です。 このシナリオでは、コードは jQuery を呼び出し `ready` 、そのコールバックを使用して進行状況インジケーターをレンダリングする必要があります。 その後、Office `onReady` のコールバックで、進行状況のインジケーターを最終的な UI に置き換えることができます。 

## <a name="initialize-with-officeinitialize"></a>Office.initialize を使用した初期化

Office.js ライブラリが読み込まれ、ユーザーとの対話の準備が完了すると、初期化イベントが発生します。 初期化ロジックを実装する `Office.initialize` にハンドラーを割り当てることができます。 ユーザーのバージョンの Excel が、アドインで呼び出す可能性があるすべての API をサポートしているかを確認する例は、次のとおりです。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

独自の初期化ハンドラーまたはテストを含む追加の JavaScript  `Office.initialize` フレームワークを使用している場合は、通常、イベント内に配置する必要があります (この場合は、前述の 「**Office.onReady()** で初期化する」セクションで説明されている例外も適用されます)。 たとえば、[JQuery](https://jquery.com) の `$(document).ready()` 関数は次のように参照します。

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

- `Office.initialize` イベントは、Office.js 自体が初期化される内部プロセスの最後に発生します。 内部のプロセスが終了した後、*すぐに* 発生します。 イベントにハンドラーを割り当てるコードが、イベント発生後に長時間実行される場合、ハンドラーは実行されません。 たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読み込まれた後で、カスタム JavaScript を読み込む前に、ポリフィルのファイルを読み込むためのアドインのホーム ページを構成する場合があります。 ご使用のスクリプトでハンドラーの読み込みと割り当てが行われる時点で、初期化イベントは既に発生しています。 ですが、`Office.onReady()` を呼び出すのに "遅すぎる" ことは決してありません。 初期化イベントが既に発生している場合、コールバックがすぐに実行されます。

> [!NOTE]
> スタートアップ ロジックがない場合でも、アドイン JavaScript を読み込むときには、`Office.onReady()` を呼び出すか、または空の関数を `Office.initialize` に割り当てる必要があります。 アプリケーションOfficeプラットフォームの組み合わせの中には、作業ウィンドウが読み込み込みできない場合があります。 次の例はこの 2 つの方法を示しています。
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