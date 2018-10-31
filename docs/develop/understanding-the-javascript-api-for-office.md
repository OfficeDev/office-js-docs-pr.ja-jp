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
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="6910c-102">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="6910c-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="6910c-p101">この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6910c-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="6910c-p102">AppSource にアドインを [ [公開](../publish/publish.md) ]し、Office エクスペリエンスで利用できるようにする予定がある場合は、[ [AppSource の検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)]に準拠していることを確認してください。たとえば、検証に合格するためには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[ [セクション 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) ] と [ [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md) ]のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="6910c-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="6910c-108">アドインで JavaScript API for Office ライブラリを参照する</span><span class="sxs-lookup"><span data-stu-id="6910c-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="6910c-p103">[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。</span><span class="sxs-lookup"><span data-stu-id="6910c-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="6910c-111">これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。</span><span class="sxs-lookup"><span data-stu-id="6910c-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="6910c-112">バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、[「 JavaScript API for Office ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する」を参照してください。](referencing-the-javascript-api-for-office-library-from-its-cdn.md)</span><span class="sxs-lookup"><span data-stu-id="6910c-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="6910c-113">アドインを初期化しています</span><span class="sxs-lookup"><span data-stu-id="6910c-113">Initializing your add-in</span></span>

<span data-ttu-id="6910c-114">**適用対象:** すべてのアドインの種類</span><span class="sxs-lookup"><span data-stu-id="6910c-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="6910c-115">Office アドインでは、次のように処理を実行するスタートアップ ロジックが多くある場合があります。</span><span class="sxs-lookup"><span data-stu-id="6910c-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="6910c-116">ユーザーの Office のバージョンがコードを呼び出すすべての Office APIをサポートするかを確認します。</span><span class="sxs-lookup"><span data-stu-id="6910c-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="6910c-117">特定の名前を含むワークシートなどの特定の成果物の有無を確認します。</span><span class="sxs-lookup"><span data-stu-id="6910c-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="6910c-118">Excel では、いくつかのセルを選択するプロンプトを表示し、選択した値で初期化されたグラフを挿入することです。</span><span class="sxs-lookup"><span data-stu-id="6910c-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="6910c-119">バインディングを確立します。</span><span class="sxs-lookup"><span data-stu-id="6910c-119">Establish bindings.</span></span>

- <span data-ttu-id="6910c-120">Office ダイアログ ボックス API を使用して、アドインの設定の既定値をユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="6910c-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="6910c-p104">ライブラリが完全にロードされるまで、スタートアップ コードは Office.js Api を呼び出すしない必要があります。コードがライブラリがロードされていることを確認する 2 つの方法があります。それらについては、以下のセクションで説明します。新しいより柔軟性が高いこの手法を使用することをお勧めします。 呼び出し `Office.onReady()`。ハンドラーを割り当て、古いテクニック `Office.initialize`、まだサポートされています。 [Office.initialize と Office.onReady() の間の主な相違点](#major-differences-between-office-initialize-and-office-onready)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6910c-p104">But your start-up code must not call any Office.js APIs until the library is fully loaded. There are two ways that your code can ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, technique, calling `Office.onReady()`. The older technique, assigning a handler to `Office.initialize`, is still supported. See also [Major differences between Office.initialize and Office.onReady()](#major-differences-between-office-initialize-and-office-onready).</span></span>

<span data-ttu-id="6910c-127">アドインの初期化時のイベントのシーケンスの詳細については、[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6910c-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="6910c-128">Office.onReady() を使用して初期化します。</span><span class="sxs-lookup"><span data-stu-id="6910c-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="6910c-p105">`Office.onReady()` は、Office.js ライブラリが完全に読み込まれているかどうかをチェックインするときに、Promise オブジェクトを返す非同期メソッドです。ライブラリが読み込まれるときのみ、 `Office.HostType` 列挙型の値 (`Excel`、 `Word`など) およびプラットフォーム `Office.PlatformType` 列挙型の値 (`PC`、 `Mac`、 `OfficeOnline`、など)を持つ Office ホスト アプリケーションを指定するオブジェクトとして、約束を解決します。ライブラリが既に読み込まれている場合に `Office.onReady()` を呼び出すと、約束をすぐに解決します。</span><span class="sxs-lookup"><span data-stu-id="6910c-p105">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded. When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="6910c-p106">呼び出す方法の 1 つ `Office.onReady()` コールバック メソッドを渡すことです。例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="6910c-p106">One way to call `Office.onReady()` is to pass it a callback method. Here's an example:</span></span>

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

<span data-ttu-id="6910c-p107">また、繰り返すことができます、 `then()` メソッドの呼び出しを `Office.onReady()`、コールバックを渡す代わりにします。たとえば、次のコードは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているを確認します。</span><span class="sxs-lookup"><span data-stu-id="6910c-p107">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback. For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="6910c-136">これの同じ例では、 `async` と `await` キーワードを TypeScript で使用します。</span><span class="sxs-lookup"><span data-stu-id="6910c-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="6910c-p108">独自の初期化ハンドラーやテストを含む追加の JavaScript フレームワークを使用している場合、そのようなフレームワークは `Office.onReady()` 応答の内側に配置する\* 必要 \* があります。たとえば、[ JQuery](https://jquery.com) の `$(document).ready()`  関数は次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="6910c-p108">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="6910c-p109">ただし、この実習には例外があります。たとえば、ブラウザーで、アドインを開く (sideload の代わりに、Office ホストで) ブラウザーのツールを使用して UI をデバッグするためにします。Office.js がブラウザーに読み込まれないので `onReady` を実行できないと、 `$(document).ready` 、Office の中に呼び出されます場合は実行されません `onReady`。別の例外: アドインの読み込み中に、作業ウィンドウに表示する進行状況のインジケーターを選択します。このシナリオでは、コードは、jQuery を呼び出す必要があります `ready` のコールバックを使用して、進行状況インジケーターを表示するとします。Office では、 `onReady`のコールバックは、進行状況インジケーターを最終的な UI に置き換えることができます。</span><span class="sxs-lookup"><span data-stu-id="6910c-p109">However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools. Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`. Another exception: you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator. Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="6910c-145">Office.initialize を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="6910c-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="6910c-p110">Office.js ライブラリは、完全に読み込まれ、ユーザーとの対話の準備が完了すると、initialize イベントが発生します。ハンドラーを割り当てることができます `Office.initialize` 、初期化ロジックを実装します。次に、ユーザーのバージョンの Excel がアドインを呼び出すすべての APIをサポートしているかを確認する例を示します。</span><span class="sxs-lookup"><span data-stu-id="6910c-p110">An initialize event fires when the Office.js library is fully loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="6910c-p111"> \*通常* これらは、独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合に、 `Office.initialize` イベントです。(しかし、以前にこの例で適用された *\*Office.onReady() を使用して初期化** のセクションで説明した例外もあります)。 [JQuery](https://jquery.com) の例では、 `$(document).ready()` 関数は次のように参照されるでしょう。</span><span class="sxs-lookup"><span data-stu-id="6910c-p111">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event. (But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="6910c-p112">作業ウィンドウ アドインとコンテンツ アドインについては、`Office.initialize` は、追加の _reason_ パラメーター提供します。このパラメーターは、アドインがどのように現在のドキュメントに追加されたかを判断するために使用できます。これは、最初にアドインが挿入されたときと、既にアドインがドキュメント内に存在しているときに別のロジックを提供するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="6910c-p112">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="6910c-154">詳細については、「[Office.initialize イベント](https://docs.microsoft.com/javascript/api/office?view=office-js)」および「[InitializationReason 列挙型](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6910c-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="6910c-155">Office.initialize と Office.onReady の間の主な相違点</span><span class="sxs-lookup"><span data-stu-id="6910c-155">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="6910c-p113">ハンドラーを 1 つだけを`Office.initialize`に割り当てることができ、1 回だけ、Office のインフラストラクチャで呼び出すことができます。しかしコードの異なる場所で `Office.onReady()` を呼び出すことができますが、異なるコールバックを使用してください。  初期化ロジックを実行するコールバックをカスタム スクリプトが読み込まれるとすぐにコードは`Office.onReady()`を呼び出すかもしれません。コードは、作業ウィンドウに、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出すボタンをもっているかもしれません。その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。</span><span class="sxs-lookup"><span data-stu-id="6910c-p113">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="6910c-p114"> `Office.initialize` イベントが、 Office.js 自身の初期化の内部プロセスの最後に発生します。内部のプロセスが終了した後に \*すぐ* に発生します。イベント発生後、イベントにハンドラーを割り当てるコードが長時間実行した場合、ハンドラーは実行されません。たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読みこんで、しかしカスタムの JavaScriptを読み込む前に、polyfillファイルをロードするようにアドインのホーム ページを構成する場合があります。この時点で、スクリプトをロードし、ハンドラーを割り当てます、initialize イベントは、すでに実行されています。`Office.onReady()` を呼び出すことは決して「手遅れ」ではありません、Initialize イベントは、すでに実行されており、すぐにコールバックが実行されます。</span><span class="sxs-lookup"><span data-stu-id="6910c-p114">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it is never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="6910c-p115">スタートアップ ロジックがない場合でも、 JアドインのJavaScript を読み込む際に `Office.onReady()` を呼び出すか、または `Office.initialize` に空の関数を割り当てることは良い練習になります。これは、Office のホストとプラットフォームの組み合わせによっては、これらのいずれかが発生するまで、作業ウィンドウをロードできないためです。以下の二つの行は、これが行われる二つの方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="6910c-p115">Even if you have no start-up logic, it is a good practice to either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads, because some Office host and platform combinations won't load the task pane until one of these happens. The following lines show the two ways this can be done:</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="6910c-168">Office JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="6910c-168">Office JavaScript API object model</span></span>

<span data-ttu-id="6910c-p116">初期化されたアドインはホスト (例 : Excel、Outlook) と連携できます。「[Office JavaScript API オブジェクトモデル](office-javascript-api-object-model.md)」のページで特定の使用パターンの詳細を見ることができます。また、[共有 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) と特定のホストに関する詳細なリファレンス ドキュメントも用意されています。</span><span class="sxs-lookup"><span data-stu-id="6910c-p116">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns. There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="6910c-172">API サポート マトリックス</span><span class="sxs-lookup"><span data-stu-id="6910c-172">API support matrix</span></span>

<span data-ttu-id="6910c-173">次の表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](update-your-javascript-api-for-office-and-manifest-schema-version.md) を使用してアドインがサポートする Office のホスト アプリケーションを指定する際に、これらの API と機能をホストする Office アプリケーションについてまとめたものです。</span><span class="sxs-lookup"><span data-stu-id="6910c-173">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="6910c-174">**ホスト名**</span><span class="sxs-lookup"><span data-stu-id="6910c-174">**Host name**</span></span>|<span data-ttu-id="6910c-175">データベース</span><span class="sxs-lookup"><span data-stu-id="6910c-175">Database</span></span>|<span data-ttu-id="6910c-176">ブック</span><span class="sxs-lookup"><span data-stu-id="6910c-176">Workbook</span></span>|<span data-ttu-id="6910c-177">メールボックス</span><span class="sxs-lookup"><span data-stu-id="6910c-177">Mailbox</span></span>|<span data-ttu-id="6910c-178">プレゼンテーション</span><span class="sxs-lookup"><span data-stu-id="6910c-178">Presentation</span></span>|<span data-ttu-id="6910c-179">ドキュメント</span><span class="sxs-lookup"><span data-stu-id="6910c-179">Document</span></span>|<span data-ttu-id="6910c-180">プロジェクト</span><span class="sxs-lookup"><span data-stu-id="6910c-180">Project</span></span>|
||<span data-ttu-id="6910c-181">**サポートされる\*\*\*\*ホスト アプリケーション**</span><span class="sxs-lookup"><span data-stu-id="6910c-181">**Supported** **Host applications**</span></span>|<span data-ttu-id="6910c-182">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="6910c-182">Access web apps</span></span>|<span data-ttu-id="6910c-183">Excel、</span><span class="sxs-lookup"><span data-stu-id="6910c-183">Excel,</span></span><br/><span data-ttu-id="6910c-184">Excel Online</span><span class="sxs-lookup"><span data-stu-id="6910c-184">Excel Online</span></span>|<span data-ttu-id="6910c-185">Outlook、</span><span class="sxs-lookup"><span data-stu-id="6910c-185">Outlook,</span></span><br/><span data-ttu-id="6910c-186">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="6910c-186">Outlook Web App,</span></span><br/><span data-ttu-id="6910c-187">デバイス用OWA</span><span class="sxs-lookup"><span data-stu-id="6910c-187">OWA for Devices</span></span>|<span data-ttu-id="6910c-188">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="6910c-188">PowerPoint,</span></span><br/><span data-ttu-id="6910c-189">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="6910c-189">PowerPoint Online</span></span>|<span data-ttu-id="6910c-190">Word</span><span class="sxs-lookup"><span data-stu-id="6910c-190">Word</span></span>|<span data-ttu-id="6910c-191">プロジェクト</span><span class="sxs-lookup"><span data-stu-id="6910c-191">Project</span></span>|
|<span data-ttu-id="6910c-192">**サポートされるアドインの種類**</span><span class="sxs-lookup"><span data-stu-id="6910c-192">**Supported add-in types**</span></span>|<span data-ttu-id="6910c-193">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="6910c-193">Content</span></span>|<span data-ttu-id="6910c-194">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-194">Y</span></span>|<span data-ttu-id="6910c-195">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-195">Y</span></span>||<span data-ttu-id="6910c-196">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-196">Y</span></span>|||
||<span data-ttu-id="6910c-197">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="6910c-197">Task pane</span></span>||<span data-ttu-id="6910c-198">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-198">Y</span></span>||<span data-ttu-id="6910c-199">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-199">Y</span></span>|<span data-ttu-id="6910c-200">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-200">Y</span></span>|<span data-ttu-id="6910c-201">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-201">Y</span></span>|
||<span data-ttu-id="6910c-202">Outlook</span><span class="sxs-lookup"><span data-stu-id="6910c-202">Outlook</span></span>|||<span data-ttu-id="6910c-203">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-203">Y</span></span>||||
|<span data-ttu-id="6910c-204">**サポートされている API 機能**</span><span class="sxs-lookup"><span data-stu-id="6910c-204">**Supported API features**</span></span>|<span data-ttu-id="6910c-205">テキストの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-205">Read/Write Text</span></span>||<span data-ttu-id="6910c-206">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-206">Y</span></span>||<span data-ttu-id="6910c-207">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-207">Y</span></span>|<span data-ttu-id="6910c-208">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-208">Y</span></span>|<span data-ttu-id="6910c-209">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-209">Y</span></span><br/><span data-ttu-id="6910c-210">(読み取り専用)</span><span class="sxs-lookup"><span data-stu-id="6910c-210">(Read only)</span></span>|
||<span data-ttu-id="6910c-211">マトリックスの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-211">Read/Write Matrix</span></span>||<span data-ttu-id="6910c-212">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-212">Y</span></span>|||<span data-ttu-id="6910c-213">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-213">Y</span></span>||
||<span data-ttu-id="6910c-214">テーブルの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-214">Read/Write Table</span></span>||<span data-ttu-id="6910c-215">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-215">Y</span></span>|||<span data-ttu-id="6910c-216">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-216">Y</span></span>||
||<span data-ttu-id="6910c-217">HTML の読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-217">Read/Write HTML</span></span>|||||<span data-ttu-id="6910c-218">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-218">Y</span></span>||
||<span data-ttu-id="6910c-219">読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-219">Read/Write</span></span><br/><span data-ttu-id="6910c-220">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="6910c-220">Office Open XML</span></span>|||||<span data-ttu-id="6910c-221">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-221">Y</span></span>||
||<span data-ttu-id="6910c-222">タスク、リソース、ビュー、フィールド プロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="6910c-222">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="6910c-223">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-223">Y</span></span>|
||<span data-ttu-id="6910c-224">選択変更イベント</span><span class="sxs-lookup"><span data-stu-id="6910c-224">Selection changed events</span></span>||<span data-ttu-id="6910c-225">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-225">Y</span></span>|||<span data-ttu-id="6910c-226">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-226">Y</span></span>||
||<span data-ttu-id="6910c-227">ドキュメント全体の取得</span><span class="sxs-lookup"><span data-stu-id="6910c-227">Get whole document</span></span>||||<span data-ttu-id="6910c-228">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-228">Y</span></span>|<span data-ttu-id="6910c-229">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-229">Y</span></span>||
||<span data-ttu-id="6910c-230">バインドとイベント バインド</span><span class="sxs-lookup"><span data-stu-id="6910c-230">Bindings and binding events</span></span>|<span data-ttu-id="6910c-231">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-231">Y</span></span><br/><span data-ttu-id="6910c-232">(完全および部分的なテーブル バインドのみ)</span><span class="sxs-lookup"><span data-stu-id="6910c-232">(Only full and partial table bindings)</span></span>|<span data-ttu-id="6910c-233">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-233">Y</span></span>|||<span data-ttu-id="6910c-234">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-234">Y</span></span>||
||<span data-ttu-id="6910c-235">カスタム XML パーツの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="6910c-235">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="6910c-236">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-236">Y</span></span>||
||<span data-ttu-id="6910c-237">アドイン状態データの保持 (設定)</span><span class="sxs-lookup"><span data-stu-id="6910c-237">Persist add-in state data (settings)</span></span>|<span data-ttu-id="6910c-238">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-238">Y</span></span><br/><span data-ttu-id="6910c-239">(ホスト アドインごと)</span><span class="sxs-lookup"><span data-stu-id="6910c-239">(Per host add-in)</span></span>|<span data-ttu-id="6910c-240">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-240">Y</span></span><br/><span data-ttu-id="6910c-241">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="6910c-241">(Per document)</span></span>|<span data-ttu-id="6910c-242">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-242">Y</span></span><br/><span data-ttu-id="6910c-243">(メールボックスごと)</span><span class="sxs-lookup"><span data-stu-id="6910c-243">(Per mailbox)</span></span>|<span data-ttu-id="6910c-244">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-244">Y</span></span><br/><span data-ttu-id="6910c-245">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="6910c-245">(Per document)</span></span>|<span data-ttu-id="6910c-246">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-246">Y</span></span><br/><span data-ttu-id="6910c-247">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="6910c-247">(Per document)</span></span>||
||<span data-ttu-id="6910c-248">設定変更イベント</span><span class="sxs-lookup"><span data-stu-id="6910c-248">Settings changed events</span></span>|<span data-ttu-id="6910c-249">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-249">Y</span></span>|<span data-ttu-id="6910c-250">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-250">Y</span></span>||<span data-ttu-id="6910c-251">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-251">Y</span></span>|<span data-ttu-id="6910c-252">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-252">Y</span></span>||
||<span data-ttu-id="6910c-253">アクティブ ビュー モードの取得</span><span class="sxs-lookup"><span data-stu-id="6910c-253">Get active view mode</span></span><br/><span data-ttu-id="6910c-254">およびビュー変更イベントの取得</span><span class="sxs-lookup"><span data-stu-id="6910c-254">and view changed events</span></span>||||<span data-ttu-id="6910c-255">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-255">Y</span></span>|||
||<span data-ttu-id="6910c-256">ドキュメント内の</span><span class="sxs-lookup"><span data-stu-id="6910c-256">Navigate to locations</span></span><br/><span data-ttu-id="6910c-257">場所に移動</span><span class="sxs-lookup"><span data-stu-id="6910c-257">in the document</span></span>||<span data-ttu-id="6910c-258">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-258">Y</span></span>||<span data-ttu-id="6910c-259">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-259">Y</span></span>|<span data-ttu-id="6910c-260">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-260">Y</span></span>||
||<span data-ttu-id="6910c-261">ルールと RegEx を使用した</span><span class="sxs-lookup"><span data-stu-id="6910c-261">Activate contextually</span></span><br/><span data-ttu-id="6910c-262">文脈からのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="6910c-262">using rules and RegEx</span></span>|||<span data-ttu-id="6910c-263">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-263">Y</span></span>||||
||<span data-ttu-id="6910c-264">アイテム プロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="6910c-264">Read Item properties</span></span>|||<span data-ttu-id="6910c-265">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-265">Y</span></span>||||
||<span data-ttu-id="6910c-266">ユーザー プロファイルの読み取り</span><span class="sxs-lookup"><span data-stu-id="6910c-266">Read User profile</span></span>|||<span data-ttu-id="6910c-267">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-267">Y</span></span>||||
||<span data-ttu-id="6910c-268">添付ファイルの取得</span><span class="sxs-lookup"><span data-stu-id="6910c-268">Get attachments</span></span>|||<span data-ttu-id="6910c-269">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-269">Y</span></span>||||
||<span data-ttu-id="6910c-270">ユーザー ID トークンの取得</span><span class="sxs-lookup"><span data-stu-id="6910c-270">Get User identity token</span></span>|||<span data-ttu-id="6910c-271">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-271">Y</span></span>||||
||<span data-ttu-id="6910c-272">Exchange Web サービスの呼出</span><span class="sxs-lookup"><span data-stu-id="6910c-272">Call Exchange Web Services</span></span>|||<span data-ttu-id="6910c-273">Y</span><span class="sxs-lookup"><span data-stu-id="6910c-273">Y</span></span>||||
