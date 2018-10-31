---
title: JavaScript API for Office について
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e9d9efdda5e237ab076d22d50b1f7ded5e075845
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505951"
---
# <a name="understanding-the-javascript-api-for-office"></a>JavaScript API for Office について

この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。

> [!NOTE]
> AppSource にアドインを [ [公開](../publish/publish.md) ]し、Office エクスペリエンスで利用できるようにする予定がある場合は、[ [AppSource の検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)]に準拠していることを確認してください。たとえば、検証に合格するためには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[ [セクション 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) ] と [ [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md) ]のページを参照してください)。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>アドインで JavaScript API for Office ライブラリを参照する

[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。

バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、[「 JavaScript API for Office ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する」を参照してください。](referencing-the-javascript-api-for-office-library-from-its-cdn.md)

## <a name="initializing-your-add-in"></a>アドインを初期化しています

**適用対象:** すべてのアドインの種類

Office アドインでは、次のように処理を実行するスタートアップ ロジックが多くある場合があります。

- ユーザーの Office のバージョンがコードを呼び出すすべての Office APIをサポートするかを確認します。

- 特定の名前を含むワークシートなどの特定の成果物の有無を確認します。

- Excel では、いくつかのセルを選択するプロンプトを表示し、選択した値で初期化されたグラフを挿入することです。

- バインディングを確立します。

- Office ダイアログ ボックス API を使用して、アドインの設定の既定値をユーザーに確認します。

ライブラリが完全にロードされるまで、スタートアップ コードは Office.js Api を呼び出すしない必要があります。コードがライブラリがロードされていることを確認する 2 つの方法があります。それらについては、以下のセクションで説明します。新しいより柔軟性が高いこの手法を使用することをお勧めします。 呼び出し `Office.onReady()`。ハンドラーを割り当て、古いテクニック `Office.initialize`、まだサポートされています。 [Office.initialize と Office.onReady() の間の主な相違点](#major-differences-between-office-initialize-and-office-onready)を参照してください。

アドインの初期化時のイベントのシーケンスの詳細については、[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)を参照してください。

### <a name="initialize-with-officeonready"></a>Office.onReady() を使用して初期化します。

`Office.onReady()` は、Office.js ライブラリが完全に読み込まれているかどうかをチェックインするときに、Promise オブジェクトを返す非同期メソッドです。ライブラリが読み込まれるときのみ、 `Office.HostType` 列挙型の値 (`Excel`、 `Word`など) およびプラットフォーム `Office.PlatformType` 列挙型の値 (`PC`、 `Mac`、 `OfficeOnline`、など)を持つ Office ホスト アプリケーションを指定するオブジェクトとして、約束を解決します。ライブラリが既に読み込まれている場合に `Office.onReady()` を呼び出すと、約束をすぐに解決します。

呼び出す方法の 1 つ `Office.onReady()` コールバック メソッドを渡すことです。例を以下に示します。

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

また、繰り返すことができます、 `then()` メソッドの呼び出しを `Office.onReady()`、コールバックを渡す代わりにします。たとえば、次のコードは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているを確認します。

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

これの同じ例では、 `async` と `await` キーワードを TypeScript で使用します。

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、そのようなフレームワークは `Office.onReady()` 応答の内側に配置する* 必要 * があります。たとえば、[ JQuery](https://jquery.com) の `$(document).ready()`  関数は次のように参照します。

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

ただし、この実習には例外があります。たとえば、ブラウザーで、アドインを開く (sideload の代わりに、Office ホストで) ブラウザーのツールを使用して UI をデバッグするためにします。Office.js がブラウザーに読み込まれないので `onReady` を実行できないと、 `$(document).ready` 、Office の中に呼び出されます場合は実行されません `onReady`。別の例外: アドインの読み込み中に、作業ウィンドウに表示する進行状況のインジケーターを選択します。このシナリオでは、コードは、jQuery を呼び出す必要があります `ready` のコールバックを使用して、進行状況インジケーターを表示するとします。Office では、 `onReady`のコールバックは、進行状況インジケーターを最終的な UI に置き換えることができます。 

### <a name="initialize-with-officeinitialize"></a>Office.initialize を使用した初期化

Office.js ライブラリは、完全に読み込まれ、ユーザーとの対話の準備が完了すると、initialize イベントが発生します。ハンドラーを割り当てることができます `Office.initialize` 、初期化ロジックを実装します。次に、ユーザーのバージョンの Excel がアドインを呼び出すすべての APIをサポートしているかを確認する例を示します。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

 *通常* これらは、独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合に、 `Office.initialize` イベントです。(しかし、以前にこの例で適用された **Office.onReady() を使用して初期化** のセクションで説明した例外もあります)。 [JQuery](https://jquery.com) の例では、 `$(document).ready()` 関数は次のように参照されるでしょう。

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

作業ウィンドウ アドインとコンテンツ アドインについては、`Office.initialize` は、追加の _reason_ パラメーター提供します。このパラメーターは、アドインがどのように現在のドキュメントに追加されたかを判断するために使用できます。これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに別のロジックを提供するために使用できます。

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

詳細については、「[Office.initialize イベント](https://docs.microsoft.com/javascript/api/office?view=office-js)」および「[InitializationReason 列挙型](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)」を参照してください。

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Office.initialize と Office.onReady の間の主な相違点

- ハンドラーを 1 つだけを`Office.initialize`に割り当てることができ、1 回だけ、Office のインフラストラクチャで呼び出すことができます。しかしコードの異なる場所で `Office.onReady()` を呼び出すことができますが、異なるコールバックを使用してください。  初期化ロジックを実行するコールバックをカスタム スクリプトが読み込まれるとすぐにコードは`Office.onReady()`を呼び出すかもしれません。コードは、作業ウィンドウに、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出すボタンをもっているかもしれません。その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。

-  `Office.initialize` イベントが、 Office.js 自身の初期化の内部プロセスの最後に発生します。内部のプロセスが終了した後に *すぐ* に発生します。イベント発生後、イベントにハンドラーを割り当てるコードが長時間実行した場合、ハンドラーは実行されません。たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読みこんで、しかしカスタムの JavaScriptを読み込む前に、polyfillファイルをロードするようにアドインのホーム ページを構成する場合があります。この時点で、スクリプトをロードし、ハンドラーを割り当てます、initialize イベントは、すでに実行されています。`Office.onReady()` を呼び出すことは決して「手遅れ」ではありません、Initialize イベントは、すでに実行されており、すぐにコールバックが実行されます。

> [!NOTE]
> スタートアップ ロジックがない場合でも、 JアドインのJavaScript を読み込む際に `Office.onReady()` を呼び出すか、または `Office.initialize` に空の関数を割り当てることは良い練習になります。これは、Office のホストとプラットフォームの組み合わせによっては、これらのいずれかが発生するまで、作業ウィンドウをロードできないためです。以下の二つの行は、これが行われる二つの方法を示しています。
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Office JavaScript オブジェクト モデル

初期化されたアドインはホスト (例 : Excel、Outlook) と連携できます。「[Office JavaScript API オブジェクトモデル](office-javascript-api-object-model.md)」のページで特定の使用パターンの詳細を見ることができます。また、[共有 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) と特定のホストに関する詳細なリファレンス ドキュメントも用意されています。

## <a name="api-support-matrix"></a>API サポート マトリックス

次の表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](update-your-javascript-api-for-office-and-manifest-schema-version.md) を使用してアドインがサポートする Office のホスト アプリケーションを指定する際に、これらの API と機能をホストする Office アプリケーションについてまとめたものです。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**ホスト名**|データベース|ブック|メールボックス|プレゼンテーション|ドキュメント|プロジェクト|
||**サポートされる****ホスト アプリケーション**|Access Web アプリ|Excel、<br/>Excel Online|Outlook、<br/>Outlook Web App、<br/>デバイス用OWA|PowerPoint、<br/>PowerPoint Online|Word|プロジェクト|
|**サポートされるアドインの種類**|コンテンツ|Y|Y||Y|||
||作業ウィンドウ||Y||Y|Y|Y|
||Outlook|||Y||||
|**サポートされている API 機能**|テキストの読み取り/書き込み||Y||Y|Y|Y<br/>(読み取り専用)|
||マトリックスの読み取り/書き込み||Y|||Y||
||テーブルの読み取り/書き込み||Y|||Y||
||HTML の読み取り/書き込み|||||Y||
||読み取り/書き込み<br/>Office Open XML|||||Y||
||タスク、リソース、ビュー、フィールド プロパティの読み取り||||||Y|
||選択変更イベント||Y|||Y||
||ドキュメント全体の取得||||Y|Y||
||バインドとイベント バインド|Y<br/>(完全および部分的なテーブル バインドのみ)|Y|||Y||
||カスタム XML パーツの読み取り/書き込み|||||Y||
||アドイン状態データの保持 (設定)|Y<br/>(ホスト アドインごと)|Y<br/>(ドキュメントごと)|Y<br/>(メールボックスごと)|Y<br/>(ドキュメントごと)|Y<br/>(ドキュメントごと)||
||設定変更イベント|Y|Y||Y|Y||
||アクティブ ビュー モードの取得<br/>およびビュー変更イベントの取得||||Y|||
||ドキュメント内の<br/>場所に移動||Y||Y|Y||
||ルールと RegEx を使用した<br/>文脈からのアクティブ化|||Y||||
||アイテム プロパティの読み取り|||Y||||
||ユーザー プロファイルの読み取り|||Y||||
||添付ファイルの取得|||Y||||
||ユーザー ID トークンの取得|||Y||||
||Exchange Web サービスの呼出|||Y||||
