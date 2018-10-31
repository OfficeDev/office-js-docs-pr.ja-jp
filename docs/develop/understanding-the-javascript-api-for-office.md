---
title: JavaScript API for Office について
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 58829c623c06225bcc7d15925fb02a082df039c6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640093"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="39ded-102">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="39ded-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="39ded-p101">この記事では、JavaScript API for Office とその使用方法に関する情報を提供します。参照情報については、「[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)」を参照してください。Visual Studio プロジェクト ファイルを JavaScript API for Office の最新バージョンに更新する方法については、「[JavaScript API for Office およびマニフェスト スキーマ ファイルのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ded-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="39ded-p102">AppSource にアドインを [ [公開](../publish/publish.md) ]し、Office エクスペリエンスで利用できるようにする予定がある場合は、[ [AppSource の検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)]に準拠していることを確認してください。たとえば、検証に合格するためには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[ [セクション 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) ] と [ [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md) ]のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="39ded-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="39ded-108">アドインで JavaScript API for Office ライブラリを参照する</span><span class="sxs-lookup"><span data-stu-id="39ded-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="39ded-p103">[JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) ライブラリは、Office.js ファイルと関連するホスト アプリケーション固有のファイル (Excel-15.js や Outlook-15.js など) で構成されています。最も簡単に API を参照する方法は、次に示す `<script>` をページの `<head>` タグに追加して、CDN を使用することです。</span><span class="sxs-lookup"><span data-stu-id="39ded-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="39ded-111">これにより、アドインが最初に読み込まれるときに JavaScript API for Office ファイルのダウンロードとキャッシュを実行して、アドインが確実に指定したバージョンの最新の Office.js および関連ファイルを使用するようにします。</span><span class="sxs-lookup"><span data-stu-id="39ded-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="39ded-112">バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、[「 JavaScript API for Office ライブラリをそのコンテンツ配信ネットワーク (CDN) から参照する」を参照してください。](referencing-the-javascript-api-for-office-library-from-its-cdn.md)</span><span class="sxs-lookup"><span data-stu-id="39ded-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="39ded-113">アドインを初期化しています</span><span class="sxs-lookup"><span data-stu-id="39ded-113">Initializing your add-in</span></span>

<span data-ttu-id="39ded-114">**適用対象:** すべてのアドインの種類</span><span class="sxs-lookup"><span data-stu-id="39ded-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="39ded-115">Office アドインでは、次のように処理を実行するスタートアップ ロジックが多くある場合があります。</span><span class="sxs-lookup"><span data-stu-id="39ded-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="39ded-116">ユーザーの Office のバージョンがコードを呼び出すすべての Office Api をサポートするかを確認します。</span><span class="sxs-lookup"><span data-stu-id="39ded-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="39ded-117">特定の名前を含むワークシートなどの特定の成果物の有無を確認します。</span><span class="sxs-lookup"><span data-stu-id="39ded-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="39ded-118">Excel では、いくつかのセルを選択するプロンプトを表示し、選択した値で初期化されたグラフを挿入することです。</span><span class="sxs-lookup"><span data-stu-id="39ded-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="39ded-119">バインディングを確立します。</span><span class="sxs-lookup"><span data-stu-id="39ded-119">Establish bindings.</span></span>

- <span data-ttu-id="39ded-120">Office ダイアログ API を使用して、アドインの設定の既定値をユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="39ded-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="39ded-p104">しかし、ライブラリが完全にロードされるまで、スタートアップ コードは Office.js Api を呼び出してはいけません。コードがライブラリがロードされていることを確認するには 2 つの方法があります。それらについては、以下のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="39ded-p104">But your start-up code must not call any Office.js APIs until the library is fully loaded. There are two ways that your code can ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, technique, calling . The older technique, assigning a handler to , is still supported. See also Major differences between Office.initialize and Office.onReady().</span></span> 

- [<span data-ttu-id="39ded-124">Office.onReady() を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="39ded-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="39ded-125">Office.initialize を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="39ded-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

<span data-ttu-id="39ded-126">これらの手法の違いの詳細については、 [Office.initialize と Office.onReady() の間の主な相違点](#major-differences-between-officeinitialize-and-officeonready)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ded-126">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span> <span data-ttu-id="39ded-127">アドインの初期化時のイベントのシーケンスの詳細については、[DOM とランタイム環境を読み込む](loading-the-dom-and-runtime-environment.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ded-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="39ded-128">Office.onReady() を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="39ded-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="39ded-129">`Office.onReady()` Office.js ライブラリが完全に読み込まれているかどうかをチェックするときに、Promise オブジェクトを返す非同期メソッドは、です。</span><span class="sxs-lookup"><span data-stu-id="39ded-129">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded.</span></span> <span data-ttu-id="39ded-130">ライブラリが読み込まれるときのみ、 Office ホスト アプリケーションを指定するオブジェクトとして、 `Office.HostType` 列挙型の値 (`Excel`、 `Word`など) および `Office.PlatformType` 列挙型の値 (`PC`、 `Mac`、 `OfficeOnline`、など)プラットフォームでPromiseを解決します。</span><span class="sxs-lookup"><span data-stu-id="39ded-130">When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="39ded-131">ライブラリが既に読み込まれている場合に `Office.onReady()` を呼び出すと、Promiseをすぐに解決します。</span><span class="sxs-lookup"><span data-stu-id="39ded-131">If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="39ded-132">`Office.onReady()`を呼び出す方法の 1 つは、 コールバック メソッドを渡すことです。</span><span class="sxs-lookup"><span data-stu-id="39ded-132">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="39ded-133">次に例を示します:</span><span class="sxs-lookup"><span data-stu-id="39ded-133">Here's an example:</span></span>

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

<span data-ttu-id="39ded-134">また、 `then()` メソッドの呼び出し `Office.onReady()`を、コールバックを渡す代わりにすることもできます。</span><span class="sxs-lookup"><span data-stu-id="39ded-134">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="39ded-135">たとえば、次のコードは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="39ded-135">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="39ded-136">これの同じ例では、 `async` と `await` キーワードをTypeScriptで使用します。</span><span class="sxs-lookup"><span data-stu-id="39ded-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="39ded-137">独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合、これらは *通常*  `Office.onReady()`への応答内に設置されます。</span><span class="sxs-lookup"><span data-stu-id="39ded-137">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the Office.initialize event.</span></span> <span data-ttu-id="39ded-138">例えば、[ JQuery の  ](https://jquery.com) `$(document).ready()` 関数は次のように参照されます。</span><span class="sxs-lookup"><span data-stu-id="39ded-138">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="39ded-139">ただし、この実習には例外があります。</span><span class="sxs-lookup"><span data-stu-id="39ded-139">However, there are exceptions to this practice.</span></span> <span data-ttu-id="39ded-140">たとえば、ブラウザーのツールを使用して UI をデバッグするため、ブラウザーでアドインを開く (Office ホスト内にsideload する代わりに）ことを考えます。</span><span class="sxs-lookup"><span data-stu-id="39ded-140">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="39ded-141">Office.js がブラウザーに読み込まれないので `onReady` は実行できず、 Office の中に呼び出される場合は`$(document).ready` は実行されません `onReady`。</span><span class="sxs-lookup"><span data-stu-id="39ded-141">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="39ded-142">別の例外: アドインの読み込み中に、作業ウィンドウに表示する進行状況のインジケーターを表示するようにします。</span><span class="sxs-lookup"><span data-stu-id="39ded-142">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="39ded-143">このシナリオでは、コードは、jQuery  `ready` を呼び出す必要があり、コールバックを使用して、進行状況のインジケーターを表示します。</span><span class="sxs-lookup"><span data-stu-id="39ded-143">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="39ded-144">Office では、 `onReady`のコールバックは、進行状況のインジケーターを最終的な UI に置き換えることができます。</span><span class="sxs-lookup"><span data-stu-id="39ded-144">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="39ded-145">Office.initialize を使用した初期化</span><span class="sxs-lookup"><span data-stu-id="39ded-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="39ded-146">Office.js ライブラリが完全に読み込まれ、ユーザーとの対話の準備が完了すると、initialize イベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="39ded-146">An initialize event fires when the Office.js library is fully loaded and ready for user interaction.</span></span> <span data-ttu-id="39ded-147">初期化ロジックを実装する `Office.initialize` にハンドラーを割り当てることができます。</span><span class="sxs-lookup"><span data-stu-id="39ded-147">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="39ded-148">次に示すのは、ユーザーのバージョンの Excel がアドインを呼び出すすべての Api をサポートしているかを確認する例です。</span><span class="sxs-lookup"><span data-stu-id="39ded-148">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="39ded-149">独自のハンドラーの初期化またはテストを含む追加の JavaScript フレームワークを使用する場合、これらは *通常*  `Office.initialize`への応答内に設置されます。</span><span class="sxs-lookup"><span data-stu-id="39ded-149">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the Office.initialize event.</span></span> <span data-ttu-id="39ded-150">(しかし、前に **Office.onReady() を使用した初期化** のセクションで説明した例外がこの場合も適用されます)。 [JQuery](https://jquery.com) の例では、 `$(document).ready()` 関数は次のように参照されるでしょう。</span><span class="sxs-lookup"><span data-stu-id="39ded-150">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="39ded-151">作業ウィンドウおよびコンテンツのアドインには、 `Office.initialize` が追加の _理由_ のパラメーターを提供します。</span><span class="sxs-lookup"><span data-stu-id="39ded-151">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="39ded-152">このパラメーターは、どのようにアドインが現在のドキュメントに追加されたかを指定します。</span><span class="sxs-lookup"><span data-stu-id="39ded-152">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="39ded-153">アドインが最初の挿入される場合と、文書内に既に存在していた場合に異なるロジックを提供するには、これを使用します。</span><span class="sxs-lookup"><span data-stu-id="39ded-153">For task pane and content add-ins, Office.initialize provides an additional reason parameter. This parameter can be used to determine how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="39ded-154">詳細については、「[Office.initialize イベント](https://docs.microsoft.com/javascript/api/office?view=office-js)」および「[InitializationReason 列挙型](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ded-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="39ded-155">現時点では、`Office.onReady()`も呼び出したかどうかに関係なく、`Office.Initialize`を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="39ded-155">Currently, you must set `Office.Initialize`, regardless of whether `Office.onReady()` is also called.</span></span> <span data-ttu-id="39ded-156"> `Office.Initialize\`が必要ない場合には、次の例に示すように空の関数を設定することができます。</span><span class="sxs-lookup"><span data-stu-id="39ded-156">If you have no use for `Office.Initialize`, you can set it to an empty function as shown in the following example.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="39ded-157">Office.initialize と Office.onReady の間の主な相違点</span><span class="sxs-lookup"><span data-stu-id="39ded-157">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="39ded-158">`Office.initialize`にハンドラーを 1 つだけを割り当てることができ、1 回だけは、Office のインフラストラクチャで呼び出すことができますが、 `Office.onReady()`の呼び出しは コードと異なる場所にして、異なるコールバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="39ded-158">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="39ded-159">例えば、コードは、カスタム スクリプトが初期化ロジックを実行するコールバックを読み込むとすぐに `Office.onReady()` をコールしますが、そのスクリプトが異なるコールバックで `Office.onReady()` を呼び出す作業ウィンドウにボタンを置くことができます。</span><span class="sxs-lookup"><span data-stu-id="39ded-159">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="39ded-160">その場合は、ボタンがクリックされたときに 2 番目のコールバックが実行されます。</span><span class="sxs-lookup"><span data-stu-id="39ded-160">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="39ded-161">`Office.initialize` イベントは、 Office.js 自分自身の初期化の内部プロセスの最後に発生します。</span><span class="sxs-lookup"><span data-stu-id="39ded-161">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="39ded-162">内部のプロセスが終了した後 *すぐ* に発生します。</span><span class="sxs-lookup"><span data-stu-id="39ded-162">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="39ded-163">イベントにハンドラーを割り当てるコードが、イベント発生後長時間実行した場合、ハンドラーは実行されません。</span><span class="sxs-lookup"><span data-stu-id="39ded-163">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="39ded-164">たとえば、WebPack タスク マネージャーを使用する場合は、Office.js が読み込まれた後で、カスタムjavascript を読み込む前に、ポリフィルのファイルをロードためのアドインのホーム ページを構成する場合があります。</span><span class="sxs-lookup"><span data-stu-id="39ded-164">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="39ded-165">この時点では、スクリプトはハンドラーをロードし、割り当てます。初期化 イベントは、すでに実行されています。</span><span class="sxs-lookup"><span data-stu-id="39ded-165">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="39ded-166">`Office.onReady()`呼び出すことは決して「手遅れ」ではありません。</span><span class="sxs-lookup"><span data-stu-id="39ded-166">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="39ded-167">Initialize イベントがすでに実行されている場合には、すぐにコールバックが実行されます。</span><span class="sxs-lookup"><span data-stu-id="39ded-167">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="39ded-168">スタートアップ ロジックがない場合でも、次の例に示すようにアドイン JavaScript をロードする時には、空の関数を`Office.initialize` に割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="39ded-168">Even if you have no start-up logic, you should assign an empty function to `Office.initialize` when your add-in JavaScript loads, as shown in the following example.</span></span> <span data-ttu-id="39ded-169">Office のホストとプラットフォームの組み合わせによっては、initialize イベントが発生し、指定されたイベント ハンドラ関数を実行するまで、作業ウィンドウはロードできません。</span><span class="sxs-lookup"><span data-stu-id="39ded-169">Some Office host and platform combinations won't load the task pane until the initialize event fires and the specified event handler function runs.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="39ded-170">Office JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="39ded-170">Office JavaScript API object model</span></span>

<span data-ttu-id="39ded-171">初期化されると、アドインはホスト (たとえば Excel、Outlook など) を操作できるようになります。</span><span class="sxs-lookup"><span data-stu-id="39ded-171">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="39ded-172">特定の使用パターンに関する詳細については、[Office JavaScript API オブジェクトモデル](office-javascript-api-object-model.md)ページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ded-172">The [Office JavaScript API object model](office-javascript-api-object-model.md)) page has more details on specific usage patterns.</span></span> <span data-ttu-id="39ded-173">[共有 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) および特定のホスト両方についても、詳細な参照ドキュメントがあります。</span><span class="sxs-lookup"><span data-stu-id="39ded-173">There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="39ded-174">APIサポートマトリックス</span><span class="sxs-lookup"><span data-stu-id="39ded-174">API support matrix</span></span>

<span data-ttu-id="39ded-175">次の表は、アドインの種類 (コンテンツ、作業ウィンドウ、および Outlook) 全体でサポートされている API と機能、および [1.1 アドイン マニフェスト スキーマと機能 (JavaScript API for Office v1.1 でサポート)](update-your-javascript-api-for-office-and-manifest-schema-version.md) を使用してアドインがサポートする Office のホスト アプリケーションを指定する際に、これらの API と機能をホストする Office アプリケーションについてまとめたものです。</span><span class="sxs-lookup"><span data-stu-id="39ded-175">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="39ded-176">**ホスト名**</span><span class="sxs-lookup"><span data-stu-id="39ded-176">**Host name**</span></span>|<span data-ttu-id="39ded-177">データベース</span><span class="sxs-lookup"><span data-stu-id="39ded-177">Database</span></span>|<span data-ttu-id="39ded-178">ブック</span><span class="sxs-lookup"><span data-stu-id="39ded-178">Workbook</span></span>|<span data-ttu-id="39ded-179">メールボックス</span><span class="sxs-lookup"><span data-stu-id="39ded-179">Mailbox</span></span>|<span data-ttu-id="39ded-180">プレゼンテーション</span><span class="sxs-lookup"><span data-stu-id="39ded-180">Presentation</span></span>|<span data-ttu-id="39ded-181">ドキュメント</span><span class="sxs-lookup"><span data-stu-id="39ded-181">Document</span></span>|<span data-ttu-id="39ded-182">プロジェクト</span><span class="sxs-lookup"><span data-stu-id="39ded-182">Project</span></span>|
||<span data-ttu-id="39ded-183">**サポートされる\*\*\*\*ホスト アプリケーション**</span><span class="sxs-lookup"><span data-stu-id="39ded-183">**Supported** **Host applications**</span></span>|<span data-ttu-id="39ded-184">Access Web アプリ</span><span class="sxs-lookup"><span data-stu-id="39ded-184">Access web apps</span></span>|<span data-ttu-id="39ded-185">Excel、</span><span class="sxs-lookup"><span data-stu-id="39ded-185">Excel,</span></span><br/><span data-ttu-id="39ded-186">Excel Online</span><span class="sxs-lookup"><span data-stu-id="39ded-186">Excel Online</span></span>|<span data-ttu-id="39ded-187">Outlook、</span><span class="sxs-lookup"><span data-stu-id="39ded-187">Outlook,</span></span><br/><span data-ttu-id="39ded-188">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="39ded-188">Outlook Web App,</span></span><br/><span data-ttu-id="39ded-189">デバイス用OWA</span><span class="sxs-lookup"><span data-stu-id="39ded-189">OWA for Devices</span></span>|<span data-ttu-id="39ded-190">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="39ded-190">PowerPoint,</span></span><br/><span data-ttu-id="39ded-191">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="39ded-191">PowerPoint Online</span></span>|<span data-ttu-id="39ded-192">Word</span><span class="sxs-lookup"><span data-stu-id="39ded-192">Word</span></span>|<span data-ttu-id="39ded-193">プロジェクト</span><span class="sxs-lookup"><span data-stu-id="39ded-193">Project</span></span>|
|<span data-ttu-id="39ded-194">**サポートされるアドインの種類**</span><span class="sxs-lookup"><span data-stu-id="39ded-194">**Supported add-in types**</span></span>|<span data-ttu-id="39ded-195">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ded-195">Content</span></span>|<span data-ttu-id="39ded-196">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-196">Y</span></span>|<span data-ttu-id="39ded-197">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-197">Y</span></span>||<span data-ttu-id="39ded-198">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-198">Y</span></span>|||
||<span data-ttu-id="39ded-199">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ded-199">Task pane</span></span>||<span data-ttu-id="39ded-200">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-200">Y</span></span>||<span data-ttu-id="39ded-201">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-201">Y</span></span>|<span data-ttu-id="39ded-202">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-202">Y</span></span>|<span data-ttu-id="39ded-203">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-203">Y</span></span>|
||<span data-ttu-id="39ded-204">Outlook</span><span class="sxs-lookup"><span data-stu-id="39ded-204">Outlook</span></span>|||<span data-ttu-id="39ded-205">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-205">Y</span></span>||||
|<span data-ttu-id="39ded-206">**サポートされている API 機能**</span><span class="sxs-lookup"><span data-stu-id="39ded-206">**Supported API features**</span></span>|<span data-ttu-id="39ded-207">テキストの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-207">Read/Write Text</span></span>||<span data-ttu-id="39ded-208">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-208">Y</span></span>||<span data-ttu-id="39ded-209">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-209">Y</span></span>|<span data-ttu-id="39ded-210">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-210">Y</span></span>|<span data-ttu-id="39ded-211">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-211">Y</span></span><br/><span data-ttu-id="39ded-212">(読み取り専用)</span><span class="sxs-lookup"><span data-stu-id="39ded-212">(Read only)</span></span>|
||<span data-ttu-id="39ded-213">マトリックスの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-213">Read/Write Matrix</span></span>||<span data-ttu-id="39ded-214">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-214">Y</span></span>|||<span data-ttu-id="39ded-215">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-215">Y</span></span>||
||<span data-ttu-id="39ded-216">テーブルの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-216">Read/Write Table</span></span>||<span data-ttu-id="39ded-217">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-217">Y</span></span>|||<span data-ttu-id="39ded-218">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-218">Y</span></span>||
||<span data-ttu-id="39ded-219">HTML の読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-219">Read/Write HTML</span></span>|||||<span data-ttu-id="39ded-220">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-220">Y</span></span>||
||<span data-ttu-id="39ded-221">読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-221">Read/Write</span></span><br/><span data-ttu-id="39ded-222">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="39ded-222">Office Open XML</span></span>|||||<span data-ttu-id="39ded-223">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-223">Y</span></span>||
||<span data-ttu-id="39ded-224">タスク、リソース、ビュー、フィールド プロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ded-224">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="39ded-225">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-225">Y</span></span>|
||<span data-ttu-id="39ded-226">選択変更イベント</span><span class="sxs-lookup"><span data-stu-id="39ded-226">Selection changed events</span></span>||<span data-ttu-id="39ded-227">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-227">Y</span></span>|||<span data-ttu-id="39ded-228">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-228">Y</span></span>||
||<span data-ttu-id="39ded-229">ドキュメント全体の取得</span><span class="sxs-lookup"><span data-stu-id="39ded-229">Get whole document</span></span>||||<span data-ttu-id="39ded-230">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-230">Y</span></span>|<span data-ttu-id="39ded-231">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-231">Y</span></span>||
||<span data-ttu-id="39ded-232">バインドとイベント バインド</span><span class="sxs-lookup"><span data-stu-id="39ded-232">Bindings and binding events</span></span>|<span data-ttu-id="39ded-233">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-233">Y</span></span><br/><span data-ttu-id="39ded-234">(完全および部分的なテーブル バインドのみ)</span><span class="sxs-lookup"><span data-stu-id="39ded-234">(Only full and partial table bindings)</span></span>|<span data-ttu-id="39ded-235">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-235">Y</span></span>|||<span data-ttu-id="39ded-236">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-236">Y</span></span>||
||<span data-ttu-id="39ded-237">カスタム XML パーツの読み取り/書き込み</span><span class="sxs-lookup"><span data-stu-id="39ded-237">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="39ded-238">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-238">Y</span></span>||
||<span data-ttu-id="39ded-239">アドイン状態データの保持 (設定)</span><span class="sxs-lookup"><span data-stu-id="39ded-239">Persist add-in state data (settings)</span></span>|<span data-ttu-id="39ded-240">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-240">Y</span></span><br/><span data-ttu-id="39ded-241">(ホスト アドインごと)</span><span class="sxs-lookup"><span data-stu-id="39ded-241">(Per host add-in)</span></span>|<span data-ttu-id="39ded-242">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-242">Y</span></span><br/><span data-ttu-id="39ded-243">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="39ded-243">(Per document)</span></span>|<span data-ttu-id="39ded-244">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-244">Y</span></span><br/><span data-ttu-id="39ded-245">(メールボックスごと)</span><span class="sxs-lookup"><span data-stu-id="39ded-245">(Per mailbox)</span></span>|<span data-ttu-id="39ded-246">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-246">Y</span></span><br/><span data-ttu-id="39ded-247">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="39ded-247">(Per document)</span></span>|<span data-ttu-id="39ded-248">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-248">Y</span></span><br/><span data-ttu-id="39ded-249">(ドキュメントごと)</span><span class="sxs-lookup"><span data-stu-id="39ded-249">(Per document)</span></span>||
||<span data-ttu-id="39ded-250">設定変更イベント</span><span class="sxs-lookup"><span data-stu-id="39ded-250">Settings changed events</span></span>|<span data-ttu-id="39ded-251">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-251">Y</span></span>|<span data-ttu-id="39ded-252">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-252">Y</span></span>||<span data-ttu-id="39ded-253">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-253">Y</span></span>|<span data-ttu-id="39ded-254">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-254">Y</span></span>||
||<span data-ttu-id="39ded-255">アクティブ ビュー モードの取得</span><span class="sxs-lookup"><span data-stu-id="39ded-255">Get active view mode</span></span><br/><span data-ttu-id="39ded-256">およびビュー変更イベントの取得</span><span class="sxs-lookup"><span data-stu-id="39ded-256">and view changed events</span></span>||||<span data-ttu-id="39ded-257">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-257">Y</span></span>|||
||<span data-ttu-id="39ded-258">ドキュメント内の</span><span class="sxs-lookup"><span data-stu-id="39ded-258">Navigate to locations</span></span><br/><span data-ttu-id="39ded-259">場所に移動</span><span class="sxs-lookup"><span data-stu-id="39ded-259">in the document</span></span>||<span data-ttu-id="39ded-260">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-260">Y</span></span>||<span data-ttu-id="39ded-261">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-261">Y</span></span>|<span data-ttu-id="39ded-262">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-262">Y</span></span>||
||<span data-ttu-id="39ded-263">ルールと RegEx を使用した</span><span class="sxs-lookup"><span data-stu-id="39ded-263">Activate contextually</span></span><br/><span data-ttu-id="39ded-264">文脈からのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="39ded-264">using rules and RegEx</span></span>|||<span data-ttu-id="39ded-265">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-265">Y</span></span>||||
||<span data-ttu-id="39ded-266">アイテム プロパティの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ded-266">Read Item properties</span></span>|||<span data-ttu-id="39ded-267">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-267">Y</span></span>||||
||<span data-ttu-id="39ded-268">ユーザー プロファイルの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ded-268">Read User profile</span></span>|||<span data-ttu-id="39ded-269">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-269">Y</span></span>||||
||<span data-ttu-id="39ded-270">添付ファイルの取得</span><span class="sxs-lookup"><span data-stu-id="39ded-270">Get attachments</span></span>|||<span data-ttu-id="39ded-271">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-271">Y</span></span>||||
||<span data-ttu-id="39ded-272">ユーザー ID トークンの取得</span><span class="sxs-lookup"><span data-stu-id="39ded-272">Get User identity token</span></span>|||<span data-ttu-id="39ded-273">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-273">Y</span></span>||||
||<span data-ttu-id="39ded-274">Exchange Web サービスの呼出</span><span class="sxs-lookup"><span data-stu-id="39ded-274">Call Exchange Web Services</span></span>|||<span data-ttu-id="39ded-275">Y</span><span class="sxs-lookup"><span data-stu-id="39ded-275">Y</span></span>||||
