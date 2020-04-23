---
title: DOM とランタイム環境を読み込む
description: DOM と Office アドインのランタイム環境を読み込む
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 7248f5b09a54552c3f16a9bc97bd4eae9795c8cd
ms.sourcegitcommit: 9da68c00ecc00a2f307757e0f5a903a8e31b7769
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/22/2020
ms.locfileid: "43785718"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="f6b10-103">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="f6b10-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="f6b10-104">アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="f6b10-105">コンテンツまたは作業ウィンドウ アドインの起動</span><span class="sxs-lookup"><span data-stu-id="f6b10-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="f6b10-106">次の図では、Excel、PowerPoint、Project、または Word のコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="f6b10-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![コンテンツ アドインまたは作業ウィンドウ アドイン起動時のイベントのフロー](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="f6b10-108">コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="f6b10-109">ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="f6b10-110">Office ホスト アプリケーションが、アドインの XML マニフェストを AppSource、SharePoint のアプリ カタログ、またはアドインの発生元である共有フォルダー カタログから読み取ります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-110">The Office host application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="f6b10-111">Office ホスト アプリケーションが、ブラウザー コントロールにアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-111">The Office host application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="f6b10-p101">次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="f6b10-114">ブラウザーコントロールが DOM と HTML 本文を読み込み、 `window.onload`イベントのイベントハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="f6b10-115">Office ホスト アプリケーションがランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、[Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するアドインのイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-115">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="f6b10-116">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="f6b10-117">との違いの詳細については、「[アドインを初期化する](initialize-add-in.md)」を参照してください。 `Office.onReady` `Office.initialize`</span><span class="sxs-lookup"><span data-stu-id="f6b10-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="f6b10-118">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="f6b10-119">Outlook アドインの起動</span><span class="sxs-lookup"><span data-stu-id="f6b10-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="f6b10-120">次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="f6b10-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Outlook アドイン起動時のイベントのフロー](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="f6b10-122">Outlook アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="f6b10-123">Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="f6b10-124">ユーザーが Outlook でアイテムを選択します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="f6b10-125">選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="f6b10-p103">ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="f6b10-128">ブラウザーコントロールが DOM と HTML 本文を読み込み、 `onload`イベントのイベントハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="f6b10-129">Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="f6b10-130">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="f6b10-131">との違いの詳細については、「[アドインを初期化する](initialize-add-in.md)」を参照してください。 `Office.onReady` `Office.initialize`</span><span class="sxs-lookup"><span data-stu-id="f6b10-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="f6b10-132">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="f6b10-133">読み込み状態のチェック</span><span class="sxs-lookup"><span data-stu-id="f6b10-133">Checking the load status</span></span>

<span data-ttu-id="f6b10-134">DOM とランタイム環境の両方で読み込みが完了したことを確認する方法の 1 つは、jQuery [.ready()](https://api.jquery.com/ready/) 関数を使用することです: `$(document).ready()`。</span><span class="sxs-lookup"><span data-stu-id="f6b10-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="f6b10-135">たとえば、次`onReady`のイベントハンドラーは、アドインの初期化に固有のコードが実行される前に、DOM が最初に読み込まれることを確認します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="f6b10-136">その後、 `onReady`ハンドラーは[メールボックス. item](/javascript/api/outlook/office.mailbox#item)プロパティを使用して、Outlook で現在選択されているアイテムを取得し、アドインの main 関数`initDialer`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox#item) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

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

<span data-ttu-id="f6b10-137">または、次の例に示すように`initialize` 、同じコードをイベントハンドラーで使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

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

<span data-ttu-id="f6b10-138">この方法は、 `onReady` Office アドインの`initialize`ハンドラーでも使用できます。</span><span class="sxs-lookup"><span data-stu-id="f6b10-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="f6b10-139">ダイヤラー サンプル Outlook アドインでは、JavaScript のみを使用してこれらと同じ条件を確認するという少し異なる方法を使用しています。</span><span class="sxs-lookup"><span data-stu-id="f6b10-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f6b10-140">アドインに実行する初期化タスクがない場合でも、次の例に示されているよう`Office.onReady`に、少なく`Office.initialize`とも最小のイベントハンドラー関数の呼び出しを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="f6b10-141">`Office.initialize`イベントハンドラーの呼び出し`Office.onReady`または割り当てを行わない場合、アドインの起動時にエラーが発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="f6b10-142">また、ユーザーが Excel、PowerPoint、または Outlook などの Office Web クライアントでアドインを使用しようとすると、実行に失敗します。</span><span class="sxs-lookup"><span data-stu-id="f6b10-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="f6b10-143">アドインに複数のページが含まれている場合は、新しいページが読み込まれるときに、 `Office.onReady`そのページが`Office.initialize`イベントハンドラーを呼び出すか、または割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="f6b10-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6b10-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="f6b10-144">See also</span></span>

- [<span data-ttu-id="f6b10-145">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="f6b10-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="f6b10-146">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="f6b10-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)
