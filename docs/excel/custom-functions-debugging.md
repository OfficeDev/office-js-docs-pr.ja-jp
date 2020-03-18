---
ms.date: 07/10/2019
description: Excel でカスタム関数をデバッグします。
title: カスタム関数のデバッグ
localization_priority: Normal
ms.openlocfilehash: 4abd5f3da58c35485004b17f92b334b133cabd27
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719309"
---
# <a name="custom-functions-debugging"></a><span data-ttu-id="dab98-103">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="dab98-103">Custom functions debugging</span></span>

<span data-ttu-id="dab98-104">カスタム関数のデバッグは、使用しているプラットフォームによっては複数の方法で実行できます。</span><span class="sxs-lookup"><span data-stu-id="dab98-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="dab98-105">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="dab98-105">On Windows:</span></span>
- [<span data-ttu-id="dab98-106">Excel デスクトップと Visual Studio Code (VS コード) デバッガー</span><span class="sxs-lookup"><span data-stu-id="dab98-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="dab98-107">Excel on the web および VS コードデバッガー</span><span class="sxs-lookup"><span data-stu-id="dab98-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="dab98-108">Excel on the web およびブラウザーツール</span><span class="sxs-lookup"><span data-stu-id="dab98-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="dab98-109">コマンドライン</span><span class="sxs-lookup"><span data-stu-id="dab98-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="dab98-110">On Mac:</span><span class="sxs-lookup"><span data-stu-id="dab98-110">On Mac:</span></span>
- [<span data-ttu-id="dab98-111">Excel on the web およびブラウザーツール</span><span class="sxs-lookup"><span data-stu-id="dab98-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="dab98-112">コマンドライン</span><span class="sxs-lookup"><span data-stu-id="dab98-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="dab98-113">簡単にするために、この記事では、Visual Studio Code を使用した編集、タスクの実行、および場合によってはデバッグビューを使用するためのデバッグについて説明します。</span><span class="sxs-lookup"><span data-stu-id="dab98-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="dab98-114">別のエディターまたはコマンドラインツールを使用している場合は、この記事の最後にある[コマンドラインの手順](#commands-for-building-and-running-your-add-in)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dab98-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="dab98-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="dab98-115">Requirements</span></span>

<span data-ttu-id="dab98-116">デバッグを開始する前に、 [Office アドイン用の [ごみ箱] ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、カスタム関数プロジェクトを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dab98-116">Before starting to debug, you should use the [Yeoman generator for Office add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="dab98-117">カスタム関数プロジェクトを作成する方法のガイダンスについては、「[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dab98-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="dab98-118">Excel デスクトップ用の VS コードデバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="dab98-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="dab98-119">VS コードを使用して、デスクトップ上の Office Excel でカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="dab98-119">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="dab98-120">Mac 用のデスクトップデバッグは使用できませんが、[ブラウザーツールおよびコマンドラインを使用して、web 上で Excel をデバッグすることによって](#use-the-command-line-tools-to-debug)実現できます。</span><span class="sxs-lookup"><span data-stu-id="dab98-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="dab98-121">VS コードからアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="dab98-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="dab98-122">[VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dab98-123">[**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dab98-124">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="dab98-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dab98-125">[**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="dab98-126">VS コードデバッガーを開始する</span><span class="sxs-lookup"><span data-stu-id="dab98-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="dab98-127">[**表示 > デバッグ**] を選択するか、 **Ctrl + Shift + D キー**を押してデバッグビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="dab98-127">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="dab98-128">デバッグオプションで、[ **Excel デスクトップ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-128">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="dab98-129">**F5 キーを押し**て (または、[デバッグ] **-> メニューからデバッグ開始**)、デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="dab98-129">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="dab98-130">アドインが既にサイドロードで使用できる状態で、新しい Excel ブックが開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="dab98-131">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="dab98-131">Start debugging</span></span>

1. <span data-ttu-id="dab98-132">VS Code で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="dab98-133">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="dab98-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="dab98-134">Excel ブックで、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="dab98-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="dab98-135">この時点で、ブレークポイントを設定したコード行では、この時点で実行が停止します。</span><span class="sxs-lookup"><span data-stu-id="dab98-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="dab98-136">コードをステップ実行し、ウォッチポイントを設定して、必要な VS コードデバッグ機能を使用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="dab98-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="dab98-137">Microsoft Edge で Excel の VS コードデバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="dab98-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="dab98-138">VS コードを使用して、Microsoft Edge ブラウザー上の Excel でカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="dab98-138">You can use VS Code to debug custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="dab98-139">Microsoft Edge で VS コードを使用するには、 [Microsoft edge 拡張機能用のデバッガー](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)をインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="dab98-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="dab98-140">VS コードからアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="dab98-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="dab98-141">[VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dab98-142">[**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dab98-143">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="dab98-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dab98-144">[**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="dab98-145">VS コードデバッガーを開始する</span><span class="sxs-lookup"><span data-stu-id="dab98-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="dab98-146">[**表示 > デバッグ**] を選択するか、 **Ctrl + Shift + D キー**を押してデバッグビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="dab98-146">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="dab98-147">[デバッグオプション] で、[ **Office Online (Microsoft Edge)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-147">From the Debug options, choose **Office Online (Microsoft Edge)**.</span></span>
6. <span data-ttu-id="dab98-148">Microsoft Edge ブラウザーで Excel を開き、新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="dab98-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="dab98-149">リボンの [**共有**] を選択し、この新しいブックの URL のリンクをコピーします。</span><span class="sxs-lookup"><span data-stu-id="dab98-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="dab98-150">**F5 キーを押し**ます (または、[ **> デバッグ**] を選択して、メニューからデバッグを開始します)。デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="dab98-150">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="dab98-151">ドキュメントの URL の入力を求めるプロンプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="dab98-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="dab98-152">ブックの URL を貼り付け、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="dab98-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="dab98-153">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="dab98-153">Sideload your add-in</span></span>

1. <span data-ttu-id="dab98-154">リボンの [**挿入**] タブを選択し、 **[アドイン] セクションで**、[ **Office アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="dab98-155">[ **Office アドイン**] ダイアログボックスで、[**個人用アドイン**] タブ、[**個人用アドインの管理**]、[**個人用アドインのアップロード**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="dab98-157">アドインマニフェストファイルを**参照**し、[**アップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="dab98-159">ブレークポイントを設定する</span><span class="sxs-lookup"><span data-stu-id="dab98-159">Set breakpoints</span></span>
1. <span data-ttu-id="dab98-160">VS Code で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="dab98-161">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="dab98-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="dab98-162">Excel ブックで、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="dab98-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="dab98-163">ブラウザー開発者ツールを使用して、web 上の Excel でカスタム関数をデバッグする</span><span class="sxs-lookup"><span data-stu-id="dab98-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="dab98-164">ブラウザー開発者ツールを使用して、web 上の Excel でカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="dab98-164">You can use the browser developer tools to debug custom functions in Excel on the web.</span></span> <span data-ttu-id="dab98-165">次の手順は、Windows と macOS の両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="dab98-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="dab98-166">Visual Studio Code からアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="dab98-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="dab98-167">カスタム関数のルートプロジェクトフォルダーを[Visual Studio Code (VS コード)](https://code.visualstudio.com/)で開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dab98-168">[**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dab98-169">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="dab98-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dab98-170">[**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="dab98-171">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="dab98-171">Sideload your add-in</span></span>

1. <span data-ttu-id="dab98-172">[Microsoft Office on the web](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-172">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="dab98-173">新しい Excel ブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="dab98-174">リボンの [**挿入**] タブを開き、 **[アドイン] セクションで**、[ **Office アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="dab98-175">[ **Office アドイン**] ダイアログボックスで、[**個人用アドイン**] タブ、[**個人用アドインの管理**]、[**個人用アドインのアップロード**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="dab98-177">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="dab98-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="dab98-179">サイドロードしたドキュメントは、ドキュメントを開くたびにサイドロードされたままになります。</span><span class="sxs-lookup"><span data-stu-id="dab98-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="dab98-180">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="dab98-180">Start debugging</span></span>

1. <span data-ttu-id="dab98-181">開発者ツールをブラウザーで開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-181">Open developer tools in the browser.</span></span> <span data-ttu-id="dab98-182">Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="dab98-183">開発者ツールで、 **Cmd + p**または**Ctrl + p** (**node.js**または**functions**) を使用してソースコードスクリプトファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="dab98-184">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="dab98-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="dab98-185">コードを変更する必要がある場合は、VS コードで編集を行って変更を保存することができます。</span><span class="sxs-lookup"><span data-stu-id="dab98-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="dab98-186">ブラウザーを更新して、変更が読み込まれたことを確認します。</span><span class="sxs-lookup"><span data-stu-id="dab98-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="dab98-187">コマンドラインツールを使用してデバッグする</span><span class="sxs-lookup"><span data-stu-id="dab98-187">Use the command line tools to debug</span></span>

<span data-ttu-id="dab98-188">VS コードを使用していない場合は、コマンドライン (bash、PowerShell など) を使用してアドインを実行できます。</span><span class="sxs-lookup"><span data-stu-id="dab98-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="dab98-189">Web 上の Excel でコードをデバッグするには、ブラウザー開発者ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dab98-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="dab98-190">コマンドラインを使用して、デスクトップ版の Excel をデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="dab98-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="dab98-191">コマンドラインからを実行`npm run watch`すると、コードの変更が発生したときにを監視し、再構築します。</span><span class="sxs-lookup"><span data-stu-id="dab98-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="dab98-192">2番目のコマンドラインウィンドウを開きます (最初のウィンドウは、ウォッチの実行中にブロックされます)。</span><span class="sxs-lookup"><span data-stu-id="dab98-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="dab98-193">Excel のデスクトップバージョンでアドインを起動するには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="dab98-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="dab98-194">または、web 上の Excel でアドインを開始する場合は、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="dab98-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="dab98-195">Excel on the web では、アドインをサイドロードする必要もあります。</span><span class="sxs-lookup"><span data-stu-id="dab98-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="dab98-196">「[サイドロード](#sideload-your-add-in)を使用してアドインをサイドロードする」の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="dab98-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="dab98-197">その後、次のセクションに進み、デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="dab98-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="dab98-198">開発者ツールをブラウザーで開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-198">Open developer tools in the browser.</span></span> <span data-ttu-id="dab98-199">Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="dab98-200">[開発者ツール] で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。</span><span class="sxs-lookup"><span data-stu-id="dab98-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="dab98-201">カスタム関数のコードは、ファイルの末尾付近に配置されている場合があります。</span><span class="sxs-lookup"><span data-stu-id="dab98-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="dab98-202">カスタム関数のソースコードで、コードの行を選択してブレークポイントを適用します。</span><span class="sxs-lookup"><span data-stu-id="dab98-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="dab98-203">コードを変更する必要がある場合は、Visual Studio で編集を行って変更を保存することができます。</span><span class="sxs-lookup"><span data-stu-id="dab98-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="dab98-204">ブラウザーを更新して、変更が読み込まれたことを確認します。</span><span class="sxs-lookup"><span data-stu-id="dab98-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="dab98-205">アドインをビルドして実行するためのコマンド</span><span class="sxs-lookup"><span data-stu-id="dab98-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="dab98-206">使用可能なビルドタスクはいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="dab98-206">There are several build tasks available:</span></span>
- <span data-ttu-id="dab98-207">`npm run watch`: ソースファイルの保存時に開発用のビルドを作成し、自動的に再構築します。</span><span class="sxs-lookup"><span data-stu-id="dab98-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="dab98-208">`npm run build-dev`: 開発用ビルド</span><span class="sxs-lookup"><span data-stu-id="dab98-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="dab98-209">`npm run build`: 運用のためのビルド</span><span class="sxs-lookup"><span data-stu-id="dab98-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="dab98-210">`npm run dev-server`: 開発に使用する web サーバーを実行します。</span><span class="sxs-lookup"><span data-stu-id="dab98-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="dab98-211">次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。</span><span class="sxs-lookup"><span data-stu-id="dab98-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="dab98-212">`npm run start:desktop`: デスクトップ上で Excel を起動し、アドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="dab98-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="dab98-213">`npm run start:web`: Web 上で Excel を起動し、アドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="dab98-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="dab98-214">`npm run stop`: Excel およびデバッグを停止します。</span><span class="sxs-lookup"><span data-stu-id="dab98-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="dab98-215">次の手順</span><span class="sxs-lookup"><span data-stu-id="dab98-215">Next steps</span></span>
<span data-ttu-id="dab98-216">[カスタム関数の認証方法](custom-functions-authentication.md)について説明します。</span><span class="sxs-lookup"><span data-stu-id="dab98-216">Learn about [authentication practices in custom functions](custom-functions-authentication.md).</span></span> <span data-ttu-id="dab98-217">または、[カスタム関数の一意のアーキテクチャ](custom-functions-architecture.md)を確認します。</span><span class="sxs-lookup"><span data-stu-id="dab98-217">Or, review [custom function's unique architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="dab98-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="dab98-218">See also</span></span>

* [<span data-ttu-id="dab98-219">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="dab98-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="dab98-220">Excel のカスタム関数でのエラー処理 </span><span class="sxs-lookup"><span data-stu-id="dab98-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="dab98-221">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="dab98-221">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="dab98-222">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="dab98-222">Create custom functions in Excel</span></span>](custom-functions-overview.md)
