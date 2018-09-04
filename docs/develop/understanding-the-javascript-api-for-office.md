---
title: JavaScript API for Office について
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 12e7d9030ec37746f84e3fc725cddda2a5675761
ms.sourcegitcommit: 5bef9828f047da03ecf2f43c6eb5b8514eff28ce
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/31/2018
ms.locfileid: "23782795"
---
# <a name="understanding-the-javascript-api-for-office"></a>JavaScript API for Office について

この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。

> [!NOTE]
> AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>アドインで JavaScript API for Office ライブラリを参照する

[JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。

バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「 JavaScript API for Office ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する」を参照してください。[ ](referencing-the-javascript-api-for-office-library-from-its-cdn.md)

## <a name="initializing-your-add-in"></a>アドインの初期化

**適用対象:** すべてのアドインの種類

Office アドインでは、次のように処理を実行するスタートアップ ロジックが多くある場合があります。

- ユーザーの Office のバージョンがコードを呼び出すすべての Office Api をサポートするかを確認します。

- 特定の名前を含むワークシートなどの特定の成果物の有無を確認します。

- Excel では、いくつかのセルを選択するプロンプトを表示し、選択した値で初期化されたグラフを挿入することです。

- バインディングを確立します。

- Office ダイアログ ボックス API を使用して、アドインの設定の既定値をユーザーに確認します。

しかし、ライブラリが完全にロードされるまで、スタートアップ コードは Office.js Api を呼び出してはいけません。 コードがライブラリがロードされていることを確認する 2 つの方法があります。 それらについては、以下のセクションで説明します。 新しく、より柔軟性が高い手法、呼び出し `Office.onReady()`の使用をお勧めします。 ハンドラーを割り当てる古いテクニック `Office.initialize`、はまだサポートされています。  [Office.initialize と Office.onReady() の間の主な相違点](#major-differences-between-office-initialize-and-office-onready)を参照してください。

アドインの初期化時のイベントのシーケンスの詳細については、[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)を参照してください。

### <a name="initialize-with-officeonready"></a>Office.onReady() を使用して初期化します。

`Office.onReady()` Office.js ライブラリが完全に読み込まれているかどうかをチェックするときに、Promise オブジェクトを返す非同期メソッドは、です。 ライブラリが読み込まれるときのみ、 Office ホスト アプリケーションを指定するオブジェクトとして、 `Office.HostType` 列挙型の値 (`Excel`、 `Word`など) および `Office.PlatformType` 列挙型の値 (`PC`、 `Mac`、 `OfficeOnline`、など)プラットフォームでPromiseを解決します。 ライブラリが既に読み込まれている場合に `Office.onReady()` を呼び出すと、Promiseをすぐに解決します。

`Office.onReady()`を呼び出す方法の 1 つは、 コールバック メソッドを渡すことです。 次に例を示します:

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

また、 `then()` メソッドの呼び出し `Office.onReady()`を、コールバックを渡す代わりにすることもできます。 たとえば、次のコードは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているかどうかを確認します。

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

これの同じ例では、 `async` と `await` キーワードをTypeScriptで使用します。

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合、これらは *通常*  `Office.onReady()`への応答内に設置されます。 たとえば、[ JQuery の  ](https://jquery.com) `$(document).ready()` 関数は次のように参照されます。

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

ただし、この実習には例外があります。 たとえば、ブラウザーのツールを使用して UI をデバッグするため、ブラウザーでアドインを開く (Office ホスト内にsideload する代わりに）ことを考えます。 Office.js がブラウザーに読み込まれないので `onReady` は実行できず、 Office の中に呼び出される場合は`$(document).ready` は実行されません `onReady`。 別の例外: アドインの読み込み中に、作業ウィンドウに表示する進行状況のインジケーターを表示するようにします。 このシナリオでは、コードは、jQuery  `ready` を呼び出す必要があり、コールバックを使用して、進行状況のインジケーターを表示します。 Office では、 `onReady`のコールバックは、進行状況のインジケーターを最終的な UI に置き換えることができます。 

### <a name="initialize-with-officeinitialize"></a>Office.initialize を使用した初期化

Office.js ライブラリが完全に読み込まれ、ユーザーとの対話の準備が完了すると、initialize イベントが発生します。 初期化ロジックを実装する `Office.initialize` にハンドラーを割り当てることができます。 次に示すのは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているかを確認する例です。

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合、これらは *通常*  `Office.initialize`への応答内に設置されます。 (しかし、前に **Office.onReady() を使用した初期化** のセクションで説明した例外がこの場合も適用されます)。 [JQuery](https://jquery.com) の例では、 `$(document).ready()` 関数は次のように参照されるでしょう。

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

作業ウィンドウおよびコンテンツのアドインには、 `Office.initialize` が追加の _理由_ のパラメーターを提供します。 このパラメーターは、どのようにアドインが現在のドキュメントに追加されたかを指定します。 アドインが最初の挿入される場合と、文書内に既に存在していた場合に異なるロジックを提供するには、これを使用します。

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

詳細については、「[Office.initialize イベント](https://dev.office.com/reference/add-ins/shared/office.initialize)」および「[InitializationReason 列挙型](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)」を参照してください。

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Office.initialize と Office.onReadyの間の主な相違点

- `Office.initialize`にハンドラーを 1 つだけを割り当てることができ、1 回だけは、Office のインフラストラクチャで呼び出すことができますが、 `Office.onReady()`の呼び出しは コードと異なる場所にして、異なるコールバックを使用します。 例えば、コードは、カスタム スクリプトが初期化ロジックを実行するコールバックを読み込むとすぐに `Office.onReady()` をコールしますが、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出す作業ウィンドウにボタンを置くことができます。 その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。

-  `Office.initialize` イベントは、 Office.js 自分自身の初期化の内部プロセスの最後に発生します。 内部のプロセスが終了した後 *すぐ* に発生します。 イベントにハンドラーを割り当てるコードが、イベント発生後長時間実行した場合、ハンドラーは実行されません。 たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読み込まれた後で、カスタムjavascript を読み込む前に、ポリフィルのファイルをロードためのアドインのホーム ページを構成する場合があります。 この時点では、スクリプトはハンドラーをロードし、割り当てます。初期化 イベントは、すでに実行されています。 `Office.onReady()`呼び出すことは決して「手遅れ」ではありません。 Initialize イベントがすでに実行されている場合には、すぐにコールバックが実行されます。

> [!NOTE]
> スタートアップ ロジックがない場合でも、 JアドインのJavaScript を読み込む際に `Office.onReady()` を呼び出すか、または `Office.initialize` に空の関数を割り当てることは良い練習になります。これは、Office のホストとプラットフォームの組み合わせによっては、これらのいずれかが発生するまで、作業ウィンドウをロードできないためです。 次の行では、これを行う 2 つの方法を示しています。
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Office JavaScript オブジェクト モデル

初期化されると、アドインはホスト (たとえば Excel、Outlook など) を操作できるようになります。 特定の使用パターンに関する詳細については、[Office JavaScript API オブジェクトモデル](office-javascript-api-object-model.md)ページを参照してください。 [共有 API](https://dev.office.com/reference/add-ins/javascript-api-for-office) および特定のホスト両方についても、詳細な参照ドキュメントがあります。

## <a name="api-support-matrix"></a>API サポート マトリックス

次の表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](update-your-javascript-api-for-office-and-manifest-schema-version.md) を使用してアドインがサポートする Office のホスト アプリケーションを指定する際に、これらの API と機能をホストする Office アプリケーションについてまとめたものです。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**ホスト名**|データベース|ブック|メールボックス|プレゼンテーション|ドキュメント|プロジェクト|
||**サポートされる****ホスト アプリケーション**|Access Web アプリ|Excel、<br/>Excel Online|Outlook、<br/>Outlook Web App、<br/>デバイス用OWA|PowerPointでは、<br/>PowerPoint Online|Word|プロジェクト|
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
||アクティブ ビュー モード<br/>およびビュー変更イベントの取得||||Y|||
||ドキュメント内の<br/>場所に移動||Y||Y|Y||
||ルールと RegEx を使用した<br/>文脈からのアクティブ化|||Y||||
||アイテム プロパティの読み取り|||Y||||
||ユーザー プロファイルの読み取り|||Y||||
||添付ファイルの取得|||Y||||
||ユーザー ID トークンの取得|||Y||||
||Exchange Web サービスの呼出|||Y||||
