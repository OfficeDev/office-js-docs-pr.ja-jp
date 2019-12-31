---
title: 作業ウィンドウからデバッガーをアタッチする
description: ''
ms.date: 09/09/2019
localization_priority: Priority
ms.openlocfilehash: 0b0cbfc2ba86368d0bdfbaa835ce6d904cc31081
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/31/2019
ms.locfileid: "40914910"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="7175e-102">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="7175e-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="7175e-p101">Windows での Office 2016 のビルド 77xx.xxxx 以降では、作業ウィンドウからデバッガーをアタッチすることができます。デバッガーのアタッチ機能によって、デバッガーが適切な Internet Explorer プロセスに直接アタッチされます。デバッガーは、Yeoman Generator、Visual Studio Code、Node.js、Angular、その他のツールのどれを使用しているかに関係なくアタッチすることができます。</span><span class="sxs-lookup"><span data-stu-id="7175e-p101">In Office 2016 on Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, Node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="7175e-106">**デバッガーのアタッチ** ツールを起動するのには、作業ウィンドウの右上隅を選択して**パーソナリティ** メニューを有効にします (以下の図の赤い円で示す通り)。</span><span class="sxs-lookup"><span data-stu-id="7175e-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="7175e-p102">現在サポートされているデバッガー ツールは、[Update 3](https://msdn.microsoft.com/library/mt752379.aspx) 以降を適用した [Visual Studio 2015](https://www.visualstudio.com/downloads/) だけです。Visual Studio をインストールしていない場合、**デバッガーのアタッチ** オプションを選択しても何も起こりません。</span><span class="sxs-lookup"><span data-stu-id="7175e-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="7175e-p103">**[デバッガーのアタッチ]** ツールでデバッグできるのは、クライアント側の JavaScript だけです。 Node.js サーバーなど、サーバー側のコードをデバッグするには、多くのオプションがあります。 Visual Studio Code でデバッグするための詳しい方法については、「[VS Code で Node.js をデバッグする](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)」を参照してください。 Visual Studio Code を使用していない場合は、「Node.js のデバッグ」または「{サーバー名} のデバッグ」で検索してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-p103">You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![[デバッガーのアタッチ] メニューのスクリーンショット](../images/attach-debugger.png)

<span data-ttu-id="7175e-p104">**デバッガーのアタッチ** を選択するこれにより、次の図のように、**Visual Studio Just-in-Time デバッガー** ダイアログ ボックスが起動します。</span><span class="sxs-lookup"><span data-stu-id="7175e-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![Visual Studio JIT デバッガー ダイアログのスクリーンショット](../images/visual-studio-debugger.png)

<span data-ttu-id="7175e-p105">Visual Studio では、**ソリューション エクスプローラー**内にコード ファイルが表示されます。   Visual Studio でデバッグするコードの行にブレークポイントを設定することができます。</span><span class="sxs-lookup"><span data-stu-id="7175e-p105">In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="7175e-119">[パーソナリティ] メニューが表示されない場合は、Visual Studio を使用してアドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="7175e-119">If you don't see the Personality menu, you can debug your add-in using Visual Studio.</span></span> <span data-ttu-id="7175e-120">Office で作業ウィンドウ アドインが開いていることを確認してから、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="7175e-120">Ensure your task pane add-in is open in Office, and then follow these steps:</span></span>
>
> 1. <span data-ttu-id="7175e-121">Visual Studio で、**[デバッグ]** > **[プロセスにアタッチ]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="7175e-121">In Visual Studio, choose **DEBUG** > **Attach to Process**.</span></span>
> 2. <span data-ttu-id="7175e-122">**使用可能なプロセス**で、[アドインが Internet Explorer または Microsoft Edge のどちらを使用しているか](../concepts/browsers-used-by-office-web-add-ins.md)に応じて、使用可能なすべての `Iexplore.exe` プロセス*または*使用可能なすべての `MicrosoftEdge*.exe` プロセスの*どちらか*を選択し、[**添付**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="7175e-122">In **Available processes**, choose *either* all of the available `Iexplore.exe` processes *or* all of the available `MicrosoftEdge*.exe` processes, depending on [whether your add-in is using Internet Explorer or Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md), and then choose the **Attach** button.</span></span>

<span data-ttu-id="7175e-123">Visual Studio でのデバッグの詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-123">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="7175e-124">DOM Explorer を Visual Studio で起動して使用するには、ブログ記事「[新しいプロジェクト テンプレートを使って見栄えの良い Office 用アプリをビルドする](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates)」の[ヒントとコツ](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) セクションのヒント 4 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-124">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="7175e-125">ブレークポイントの設定については、「[ブレークポイントの使用](/visualstudio/debugger/using-breakpoints?view=vs-2015)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-125">To set breakpoints, see [Using Breakpoints](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span></span>
-   <span data-ttu-id="7175e-126">F12 を使用するには、「[F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-126">To use F12, see [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>
-   <span data-ttu-id="7175e-127">Microsoft Edge 開発者ツールを使用するには、「[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7175e-127">To use the Microsoft Edge Developer Tools, see [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span>

## <a name="see-also"></a><span data-ttu-id="7175e-128">関連項目</span><span class="sxs-lookup"><span data-stu-id="7175e-128">See also</span></span>

- [<span data-ttu-id="7175e-129">Visual Studio で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="7175e-129">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="7175e-130">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="7175e-130">Publish your Office Add-in</span></span>](../publish/publish.md)
