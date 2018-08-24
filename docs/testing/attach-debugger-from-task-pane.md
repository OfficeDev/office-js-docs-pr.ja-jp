---
title: 作業ウィンドウからデバッガーをアタッチする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 28f7741a6d04f8f1492fec45649cb55a9447bdd7
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925151"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="1e23b-102">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="1e23b-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="1e23b-p101">Office 2016 for Windows のビルド 77xx.xxxx 以降では、作業ウィンドウからデバッガーをアタッチすることができます。デバッガーのアタッチ機能によって、デバッガーが適切な Internet Explorer プロセスに直接アタッチされます。デバッガーは、Yeoman Generator、Visual Studio Code、node.js、Angular、その他のツールのどれを使用しているかに関係なくアタッチすることができます。</span><span class="sxs-lookup"><span data-stu-id="1e23b-p101">In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="1e23b-106">**デバッガーのアタッチ** ツールを起動するのには、作業ウィンドウの右上隅を選択して**パーソナリティ** メニューを有効にします (以下の図の赤い円で示す通り)。</span><span class="sxs-lookup"><span data-stu-id="1e23b-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="1e23b-p102">現在サポートされているデバッガー ツールは、[Update 3](https://msdn.microsoft.com/library/mt752379.aspx) 以降を適用した [Visual Studio 2015](https://www.visualstudio.com/downloads/) だけです。Visual Studio をインストールしていない場合、**デバッガーのアタッチ** オプションを選択しても何も起こりません。</span><span class="sxs-lookup"><span data-stu-id="1e23b-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="1e23b-109">**[デバッガーのアタッチ]** ツールでデバッグできるのは、クライアント側の JavaScript だけです。</span><span class="sxs-lookup"><span data-stu-id="1e23b-109">You can only debug client-side JavaScript with the **Attach Debugger** tool.</span></span> <span data-ttu-id="1e23b-110">Node.js サーバーなど、サーバー側のコードをデバッグするには、多くのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="1e23b-110">To debug server-side code, such as with a Node.js server, you have many options.</span></span> <span data-ttu-id="1e23b-111">Visual Studio Code でデバッグするための詳しい方法については、「[VS Code で Node.js をデバッグする](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-111">For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging).</span></span> <span data-ttu-id="1e23b-112">Visual Studio Code を使用していない場合は、「Node.js のデバッグ」または「{サーバー名} のデバッグ」で検索してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-112">If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![[デバッガーのアタッチ] メニューのスクリーンショット](../images/attach-debugger.png)

<span data-ttu-id="1e23b-p104">**デバッガーのアタッチ** を選択するこれにより、次の図のように、**Visual Studio Just-in-Time デバッガー** ダイアログ ボックスが起動します。</span><span class="sxs-lookup"><span data-stu-id="1e23b-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![Visual Studio JIT デバッガー ダイアログのスクリーンショット](../images/visual-studio-debugger.png)

<span data-ttu-id="1e23b-117">Visual Studio では、**ソリューション エクスプローラー**内にコード ファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1e23b-117">In Visual Studio, you will see the code files in **Solution Explorer**.</span></span>   <span data-ttu-id="1e23b-118">Visual Studio でデバッグするコードの行にブレークポイントを設定することができます。</span><span class="sxs-lookup"><span data-stu-id="1e23b-118">You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

<span data-ttu-id="1e23b-119">Visual Studio でのデバッグの詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-119">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="1e23b-120">DOM Explorer を Visual Studio で起動して使用するには、ブログ記事「[新しいプロジェクト テンプレートを使って見栄えの良い Office 用アプリをビルドする](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates)」の[ヒントとコツ](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) セクションのヒント 4 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-120">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="1e23b-121">ブレークポイントの設定については、「[ブレークポイントの使用](https://msdn.microsoft.com/library/5557y8b4.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-121">To set breakpoints, see [Using Breakpoints](https://msdn.microsoft.com/library/5557y8b4.aspx).</span></span>
-   <span data-ttu-id="1e23b-122">F12 を使用するには、「[F12 開発者ツールの使用](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1e23b-122">To use F12, see [Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>

## <a name="see-also"></a><span data-ttu-id="1e23b-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="1e23b-123">See also</span></span>

- [<span data-ttu-id="1e23b-124">Visual Studio での Office アドインの作成とデバッグ</span><span class="sxs-lookup"><span data-stu-id="1e23b-124">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="1e23b-125">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="1e23b-125">Publish your Office Add-in</span></span>](../publish/publish.md)
