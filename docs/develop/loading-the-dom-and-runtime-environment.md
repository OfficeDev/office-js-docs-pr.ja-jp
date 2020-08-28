---
title: DOM とランタイム環境を読み込む
description: DOM と Office アドインのランタイム環境を読み込む
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 02f950ca23d52b333f704c7d8aed431cb426a6f0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293276"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="a0a96-103">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="a0a96-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="a0a96-104">アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a0a96-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="a0a96-105">コンテンツまたは作業ウィンドウ アドインの起動</span><span class="sxs-lookup"><span data-stu-id="a0a96-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="a0a96-106">次の図では、Excel、PowerPoint、Project、または Word のコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="a0a96-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![コンテンツ アドインまたは作業ウィンドウ アドイン起動時のイベントのフロー](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="a0a96-108">コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="a0a96-109">ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="a0a96-110">Office クライアントアプリケーションは、アドインの XML マニフェストを AppSource、SharePoint のアプリカタログ、またはその作成元である共有フォルダーカタログから読み取ります。</span><span class="sxs-lookup"><span data-stu-id="a0a96-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="a0a96-111">Office クライアントアプリケーションは、ブラウザーコントロールでアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="a0a96-p101">次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a0a96-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="a0a96-114">ブラウザーコントロールが DOM と HTML 本文を読み込み、イベントのイベントハンドラーを呼び出し `window.onload` ます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="a0a96-115">Office クライアントアプリケーションは、ランタイム環境を読み込みます。これは、コンテンツ配布ネットワーク (CDN) サーバーから Office JavaScript API ライブラリファイルをダウンロードしてキャッシュし、 [office](/javascript/api/office)オブジェクトの[initialize](/javascript/api/office#office-initialize-reason-)イベントに対して、ハンドラーが割り当てられている場合は、アドインのイベントハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="a0a96-116">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="a0a96-117">との違いの詳細については `Office.initialize` `Office.onReady` 、「 [アドインを初期化する](initialize-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a0a96-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="a0a96-118">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="a0a96-119">Outlook アドインの起動</span><span class="sxs-lookup"><span data-stu-id="a0a96-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="a0a96-120">次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="a0a96-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Outlook アドイン起動時のイベントのフロー](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="a0a96-122">Outlook アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="a0a96-123">Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="a0a96-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="a0a96-124">ユーザーが Outlook でアイテムを選択します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="a0a96-125">選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="a0a96-p103">ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="a0a96-128">ブラウザーコントロールが DOM と HTML 本文を読み込み、イベントのイベントハンドラーを呼び出し `onload` ます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="a0a96-129">Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="a0a96-130">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="a0a96-131">との違いの詳細については `Office.initialize` `Office.onReady` 、「 [アドインを初期化する](initialize-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a0a96-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="a0a96-132">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="a0a96-133">読み込み状態のチェック</span><span class="sxs-lookup"><span data-stu-id="a0a96-133">Checking the load status</span></span>

<span data-ttu-id="a0a96-134">DOM とランタイム環境の両方で読み込みが完了したことを確認する方法の 1 つは、jQuery [.ready()](https://api.jquery.com/ready/) 関数を使用することです: `$(document).ready()`。</span><span class="sxs-lookup"><span data-stu-id="a0a96-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="a0a96-135">たとえば、次の `onReady` イベントハンドラーは、アドインの初期化に固有のコードが実行される前に、DOM が最初に読み込まれることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="a0a96-136">その後、 `onReady` ハンドラーは [メールボックス. item](/javascript/api/outlook/office.mailbox#item) プロパティを使用して、Outlook で現在選択されているアイテムを取得し、アドインの main 関数を呼び出し `initDialer` ます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox#item) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

<span data-ttu-id="a0a96-137">または、 `initialize` 次の例に示すように、同じコードをイベントハンドラーで使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

<span data-ttu-id="a0a96-138">この方法は、 `onReady` `initialize` Office アドインのハンドラーでも使用できます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="a0a96-139">ダイヤラー サンプル Outlook アドインでは、JavaScript のみを使用してこれらと同じ条件を確認するという少し異なる方法を使用しています。</span><span class="sxs-lookup"><span data-stu-id="a0a96-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a0a96-140">アドインに実行する初期化タスクがない場合でも、 `Office.onReady` 次の例に示されているように、少なくとも最小のイベントハンドラー関数の呼び出しを含める必要があり `Office.initialize` ます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="a0a96-141">`Office.onReady`イベントハンドラーの呼び出しまたは割り当てを行わない場合 `Office.initialize` 、アドインの起動時にエラーが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="a0a96-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="a0a96-142">また、ユーザーが Excel、PowerPoint、または Outlook などの Office Web クライアントでアドインを使用しようとすると、実行に失敗します。</span><span class="sxs-lookup"><span data-stu-id="a0a96-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="a0a96-143">アドインに複数のページが含まれている場合は、新しいページが読み込まれるときに、そのページがイベントハンドラーを呼び出すか、または割り当てる必要があり `Office.onReady` `Office.initialize` ます。</span><span class="sxs-lookup"><span data-stu-id="a0a96-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="a0a96-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="a0a96-144">See also</span></span>

- [<span data-ttu-id="a0a96-145">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="a0a96-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="a0a96-146">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="a0a96-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)
