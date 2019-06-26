---
title: DOM とランタイム環境を読み込む
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 78ddd10e9106e6668e2bb8cd40f58cbdb7b862d9
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128102"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="53d34-102">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="53d34-102">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="53d34-103">アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="53d34-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="53d34-104">コンテンツまたは作業ウィンドウ アドインの起動</span><span class="sxs-lookup"><span data-stu-id="53d34-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="53d34-105">次の図は、Excel、PowerPoint、Project、Word、または Access でコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="53d34-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![コンテンツ アドインまたは作業ウィンドウ アドイン起動時のイベントのフロー](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="53d34-107">コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="53d34-107">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="53d34-108">ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="53d34-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="53d34-109">Office ホスト アプリケーションが、アドインの XML マニフェストを AppSource、SharePoint のアプリ カタログ、またはアドインの発生元である共有フォルダー カタログから読み取ります。</span><span class="sxs-lookup"><span data-stu-id="53d34-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="53d34-110">Office ホスト アプリケーションが、ブラウザー コントロールにアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="53d34-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="53d34-p101">次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="53d34-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="53d34-113">ブラウザー コントロールが、DOM と HTML 本文を読み込み、**window.onload** イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="53d34-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>

5. <span data-ttu-id="53d34-114">Office ホスト アプリケーションがランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、[Office](/javascript/api/office#office-initialize) オブジェクトの [initialize](/javascript/api/office) イベントに対するアドインのイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="53d34-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="53d34-115">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="53d34-115">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="53d34-116">`Office.initialize` と `Office.onReady` の違いの詳細については、「[アドインの初期化](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="53d34-116">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>

6. <span data-ttu-id="53d34-117">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="53d34-117">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="53d34-118">Outlook アドインの起動</span><span class="sxs-lookup"><span data-stu-id="53d34-118">Startup of an Outlook add-in</span></span>

<span data-ttu-id="53d34-119">次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="53d34-119">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Outlook アドイン起動時のイベントのフロー](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="53d34-121">Outlook アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="53d34-121">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="53d34-122">Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="53d34-122">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="53d34-123">ユーザーが Outlook でアイテムを選択します。</span><span class="sxs-lookup"><span data-stu-id="53d34-123">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="53d34-124">選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。</span><span class="sxs-lookup"><span data-stu-id="53d34-124">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="53d34-p103">ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。</span><span class="sxs-lookup"><span data-stu-id="53d34-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="53d34-127">ブラウザー コントロールが DOM と HTML 本文を読み込んで、**onload** イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="53d34-127">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>

6. <span data-ttu-id="53d34-128">Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="53d34-128">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="53d34-129">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="53d34-129">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="53d34-130">`Office.initialize` と `Office.onReady` の違いの詳細については、「[アドインの初期化](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="53d34-130">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>

7. <span data-ttu-id="53d34-131">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="53d34-131">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="53d34-132">読み込み状態のチェック</span><span class="sxs-lookup"><span data-stu-id="53d34-132">Checking the load status</span></span>

<span data-ttu-id="53d34-133">DOM とランタイム環境の両方で読み込みが完了したことを確認する方法の 1 つは、jQuery [.ready()](https://api.jquery.com/ready/) 関数を使用することです: `$(document).ready()`。</span><span class="sxs-lookup"><span data-stu-id="53d34-133">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="53d34-134">たとえば、次の **onReady** イベント ハンドラーは、アドインの実行の初期化に固有のコードの前に、DOM が最初に読み込まれることを確認します。</span><span class="sxs-lookup"><span data-stu-id="53d34-134">For example, the following **onReady** event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="53d34-135">その後、**onReady** ハンドラーは [mailbox.item](/javascript/api/outlook/office.mailbox) プロパティを使用して、Outlook で現在選択されている項目を取得し、アドインのメイン関数 `initDialer` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="53d34-135">Subsequently, the **onReady** handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

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

<span data-ttu-id="53d34-136">また、次の例に示されているように、同じコードを **initialize** イベント ハンドラーで使用することができます。</span><span class="sxs-lookup"><span data-stu-id="53d34-136">Alternatively, you can use the same code in an  **initialize** event handler as shown in the following example.</span></span>

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

<span data-ttu-id="53d34-137">この方法は、任意の Office アドインの **onReady** または **initialize** ハンドラーで使用できます。</span><span class="sxs-lookup"><span data-stu-id="53d34-137">This same technique can be used in the **onReady** or **initialize** handlers of any Office Add-in.</span></span>

<span data-ttu-id="53d34-138">ダイヤラー サンプル Outlook アドインでは、JavaScript のみを使用してこれらと同じ条件を確認するという少し異なる方法を使用しています。</span><span class="sxs-lookup"><span data-stu-id="53d34-138">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="53d34-139">アドインに実行する初期化タスクがない場合でも、次の例に示されているように、**Office.onReady** の呼び出しを少なくとも 1 つ含めるか、最小のイベント ハンドラー関数 **Office.initialize** を割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="53d34-139">Even if your add-in has no initialization tasks to perform, you must include at least a call of **Office.onReady** or assign minimal **Office.initialize** event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="53d34-140">**Office.onReady** を呼び出したり、**Office.initialize** イベントを割り当てたりしない場合、アドインを開始するとエラーが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="53d34-140">If you do not call **Office.onReady** or assign an  **Office.initialize** event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="53d34-141">また、ユーザーが Excel、PowerPoint、または Outlook などの Office Web クライアントでアドインを使用しようとすると、実行に失敗します。</span><span class="sxs-lookup"><span data-stu-id="53d34-141">Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>
>
> <span data-ttu-id="53d34-142">アドインに複数のページが含まれる場合、新しいページが読み込まれるときに、そのページは **Office.onReady** を呼び出すか、**Office.initialize** イベント ハンドラーを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="53d34-142">If your add-in includes more than one page, whenever it loads a new page that page must either call **Office.onReady** or assign an  **Office.initialize** event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="53d34-143">関連項目</span><span class="sxs-lookup"><span data-stu-id="53d34-143">See also</span></span>

- [<span data-ttu-id="53d34-144">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="53d34-144">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
