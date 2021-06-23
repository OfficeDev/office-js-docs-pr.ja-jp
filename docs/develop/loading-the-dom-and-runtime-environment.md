---
title: DOM とランタイム環境を読み込む
description: DOM を読み込Officeアドインランタイム環境を使用します。
ms.date: 04/20/2021
localization_priority: Normal
ms.openlocfilehash: e66e6d5e30f5305dce35157280210a371ee3896f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076014"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="14ab5-103">DOM とランタイム環境を読み込む</span><span class="sxs-lookup"><span data-stu-id="14ab5-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="14ab5-104">アドインでは、DOM と Office アドイン両方のランタイム環境が、独自のカスタム ロジックを実行する前に読み込まれていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="14ab5-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="14ab5-105">コンテンツまたは作業ウィンドウ アドインの起動</span><span class="sxs-lookup"><span data-stu-id="14ab5-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="14ab5-106">次の図では、Excel、PowerPoint、Project、または Word のコンテンツ アドインまたは作業ウィンドウ アドインの起動に関連するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="14ab5-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![Flowまたは作業ウィンドウ アドインを開始する際のイベントの種類。](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="14ab5-108">コンテンツ アドインまたは作業ウィンドウ アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="14ab5-109">ユーザーは、既にアドインが含まれているドキュメントを開くか、ドキュメントにアドインを挿入します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="14ab5-110">クライアント Officeは、アドインの XML マニフェストを AppSource、SharePoint のアプリ カタログ、またはそれが作成元の共有フォルダー カタログから読み取ります。</span><span class="sxs-lookup"><span data-stu-id="14ab5-110">The Office client application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="14ab5-111">クライアント Officeブラウザー コントロールでアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="14ab5-111">The Office client application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="14ab5-p101">次の手順 4. と 5. は、同時に実行されることも、同時に実行されないこともあります。したがって、次の処理に進む前に、DOM とアドイン ランタイム環境の両方の読み込みが完了したことをアドインのコードで確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="14ab5-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="14ab5-114">ブラウザー コントロールは DOM と HTML 本文を読み込み、イベントのイベント ハンドラーを呼び出 `window.onload` します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="14ab5-115">Office クライアント アプリケーションはランタイム環境を読み込み、Office JavaScript API ライブラリ ファイルをコンテンツ配布ネットワーク (CDN) サーバーからダウンロードしてキャッシュし、ハンドラーが割り当てられている場合は[、Office](/javascript/api/office)オブジェクトの[](/javascript/api/office#office-initialize-reason-)initialize イベントに対してアドインのイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-115">The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="14ab5-116">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="14ab5-117">との区別の詳細については、「 `Office.initialize` アドイン `Office.onReady` の初期化 [」を参照してください](initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="14ab5-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="14ab5-118">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="14ab5-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="14ab5-119">Outlook アドインの起動</span><span class="sxs-lookup"><span data-stu-id="14ab5-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="14ab5-120">次の図は、デスクトップ、タブレット、スマートフォンで実行される Outlook アドインの起動に関係するイベントのフローを示しています。</span><span class="sxs-lookup"><span data-stu-id="14ab5-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![Flowを開始する際Outlookイベントの種類。](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="14ab5-122">Outlook アドインが起動すると、次のイベントが発生します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="14ab5-123">Outlook は起動時に、ユーザーの電子メール アカウント用にインストールされている Outlook アドインの XML マニフェストを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="14ab5-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="14ab5-124">ユーザーが Outlook でアイテムを選択します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="14ab5-125">選択されたアイテムが Outlook アドインのアクティブ化条件を満たしている場合は、Outlook がアドインをアクティブにし、ボタンを UI に表示します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="14ab5-p103">ユーザーがボタンをクリックして Outlook アドインを起動すると、Outlook がアプリの HTML ページをブラウザー コントロール内に表示します。次の手順 5 と 6 は同時に行われます。</span><span class="sxs-lookup"><span data-stu-id="14ab5-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="14ab5-128">ブラウザー コントロールは DOM と HTML 本文を読み込み、イベントのイベント ハンドラーを呼び出 `onload` します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="14ab5-129">Outlook がランタイム環境を読み込みます (このランタイム環境は、コンテンツ配布ネットワーク (CDN) サーバーから JavaScript API for JavaScript ライブラリ ファイルをダウンロードしてキャッシュします)。その後、ハンドラーが割り当てられている場合は、アドインの [Office](/javascript/api/office#office-initialize-reason-) オブジェクトの [initialize](/javascript/api/office) イベントに対するイベント ハンドラーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="14ab5-130">現時点では、コールバック (またはチェーンされた `then()` 関数) が `Office.onReady` ハンドラーに渡された (チェーンされた) かどうかも確認します。</span><span class="sxs-lookup"><span data-stu-id="14ab5-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="14ab5-131">との区別の詳細については、「 `Office.initialize` アドイン `Office.onReady` の初期化 [」を参照してください](initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="14ab5-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="14ab5-132">DOM と HTML 本文の読み込み、およびアドインの初期化が完了すると、アドインのメイン関数は処理を続行できます。</span><span class="sxs-lookup"><span data-stu-id="14ab5-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>

## <a name="see-also"></a><span data-ttu-id="14ab5-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="14ab5-133">See also</span></span>

- [<span data-ttu-id="14ab5-134">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="14ab5-134">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="14ab5-135">Office アドインを初期化する</span><span class="sxs-lookup"><span data-stu-id="14ab5-135">Initialize your Office Add-in</span></span>](initialize-add-in.md)
