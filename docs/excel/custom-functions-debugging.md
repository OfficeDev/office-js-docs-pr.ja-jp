---
ms.date: 03/13/2019
description: Excel でカスタム関数をデバッグします。
title: カスタム関数のデバッグ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 08563ef630ebc457219c4c622328b84d13e6acab
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448760"
---
# <a name="custom-functions-debugging-preview"></a><span data-ttu-id="1b9be-103">カスタム関数のデバッグ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1b9be-103">Custom functions debugging (preview)</span></span>

<span data-ttu-id="1b9be-104">カスタム関数のデバッグは、使用しているプラットフォームによっては複数の方法で実行できます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="1b9be-105">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="1b9be-105">On Windows:</span></span>
- [<span data-ttu-id="1b9be-106">Excel デスクトップと Visual Studio code (VS コード) デバッガー</span><span class="sxs-lookup"><span data-stu-id="1b9be-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="1b9be-107">Excel Online および VS コードデバッガー</span><span class="sxs-lookup"><span data-stu-id="1b9be-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="1b9be-108">Excel Online およびブラウザーツール</span><span class="sxs-lookup"><span data-stu-id="1b9be-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="1b9be-109">コマンドライン</span><span class="sxs-lookup"><span data-stu-id="1b9be-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="1b9be-110">On Mac:</span><span class="sxs-lookup"><span data-stu-id="1b9be-110">On Mac:</span></span>
- [<span data-ttu-id="1b9be-111">Excel Online およびブラウザーツール</span><span class="sxs-lookup"><span data-stu-id="1b9be-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="1b9be-112">コマンドライン</span><span class="sxs-lookup"><span data-stu-id="1b9be-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> <span data-ttu-id="1b9be-113">簡単にするために、この記事では、Visual Studio Code を使用した編集、タスクの実行、および場合によってはデバッグビューを使用するためのデバッグについて説明します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="1b9be-114">別のエディターまたはコマンドラインツールを使用している場合は、この記事の最後にある[コマンドラインの手順](#use-the-command-line-tools-to-debug)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b9be-114">If you are using a different editor or command line tool, see the [command line instructions](#use-the-command-line-tools-to-debug) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="1b9be-115">要件</span><span class="sxs-lookup"><span data-stu-id="1b9be-115">Requirements</span></span>

<span data-ttu-id="1b9be-116">デバッグを開始する前に、Yo Office ジェネレーターを使用してカスタム関数アドインプロジェクトを作成し、プロジェクトに対して信頼できる自己署名証明書があることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-116">Before starting to debug, you should create a custom functions add-in project using the Yo Office generator and ensured that you have trusted self-signed certificates for your project.</span></span> <span data-ttu-id="1b9be-117">プロジェクトを作成する手順については、「[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b9be-117">For instructions to create a project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span> <span data-ttu-id="1b9be-118">証明書の信頼の手順については、「[自己署名証明書を信頼できるルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b9be-118">For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="1b9be-119">Excel デスクトップ用の VS コードデバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="1b9be-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="1b9be-120">VS コードを使用して、デスクトップ上の Office Excel でカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-120">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="1b9be-121">Mac のデスクトップデバッグは利用できませんが[、ブラウザーツールを使用して Excel Online をデバッグする](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)ことで実現できます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="1b9be-122">VS コードからアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="1b9be-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="1b9be-123">[VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="1b9be-124">[**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="1b9be-125">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="1b9be-126">[**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="1b9be-127">VS コードデバッガーを開始する</span><span class="sxs-lookup"><span data-stu-id="1b9be-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="1b9be-128">[ **view > Debug** ] を選択するか、 **Ctrl + Shift + D**を入力してデバッグビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-128">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="1b9be-129">デバッグオプションで、[ **Excel デスクトップ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-129">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="1b9be-130">**F5 キー**を選択するか、またはメニューからデバッグ**開始 >** を選択してデバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-130">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="1b9be-131">アドインが既にサイドロードで使用できる状態で、新しい Excel ブックが開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="1b9be-132">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="1b9be-132">Start debugging</span></span>

1. <span data-ttu-id="1b9be-133">VS Code で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-133">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="1b9be-134">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="1b9be-135">Excel ブックで、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="1b9be-136">この時点で、ブレークポイントを設定したコード行では、この時点で実行が停止します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="1b9be-137">コードをステップ実行し、ウォッチポイントを設定して、必要な VS コードデバッグ機能を使用できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="1b9be-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="1b9be-138">Microsoft Edge で Excel Online 用の VS コードデバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="1b9be-138">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="1b9be-139">Microsoft Edge ブラウザーで excel Online のカスタム関数をデバッグするには、VS コードを使用できます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-139">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="1b9be-140">microsoft edge で VS コードを使用するには、 [microsoft edge 拡張機能用のデバッガー](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)をインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="1b9be-141">VS コードからアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="1b9be-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="1b9be-142">[VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="1b9be-143">[**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="1b9be-144">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="1b9be-145">[**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="1b9be-146">VS コードデバッガーを開始する</span><span class="sxs-lookup"><span data-stu-id="1b9be-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="1b9be-147">[ **view > Debug** ] を選択するか、 **Ctrl + Shift + D**を入力してデバッグビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-147">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="1b9be-148">デバッグオプションで、[ **Office Online (エッジ)**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-148">From the Debug options, choose **Office Online (Edge)**.</span></span>
6. <span data-ttu-id="1b9be-149">Microsoft Edge ブラウザーを使用して excel online を開き、excel online を開き、新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-149">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="1b9be-150">リボンの [**共有**] を選択し、この新しいブックの URL のリンクをコピーします。</span><span class="sxs-lookup"><span data-stu-id="1b9be-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="1b9be-151">**F5 キー**を選択するか、[**デバッグ > 開始**] メニューからデバッグを開始してデバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-151">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="1b9be-152">ドキュメントの URL の入力を求めるプロンプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="1b9be-153">ブックの URL を貼り付け、enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="1b9be-154">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="1b9be-154">Sideload your add-in</span></span>   

1. <span data-ttu-id="1b9be-155">リボンの [**挿入**] タブを選択し、 \*\*\*\* [アドイン] セクションで、[ **Office アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-155">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="1b9be-156">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-156">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="1b9be-158">アドインマニフェストファイルを**参照**し、[**アップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="1b9be-160">ブレークポイントを設定する</span><span class="sxs-lookup"><span data-stu-id="1b9be-160">Set breakpoints</span></span>
1. <span data-ttu-id="1b9be-161">VS Code で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-161">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="1b9be-162">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="1b9be-163">Excel ブックで、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="1b9be-164">ブラウザー開発者ツールを使用して Excel Online のカスタム関数をデバッグする</span><span class="sxs-lookup"><span data-stu-id="1b9be-164">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="1b9be-165">ブラウザー開発者ツールを使用して、Excel Online のカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-165">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="1b9be-166">次の手順は、Windows と macOS の両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="1b9be-167">Visual Studio Code からアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="1b9be-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="1b9be-168">カスタム関数のルートプロジェクトフォルダーを[Visual Studio Code (VS コード)](https://code.visualstudio.com/)で開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="1b9be-169">[**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="1b9be-170">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="1b9be-171">[**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="1b9be-172">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="1b9be-172">Sideload your add-in</span></span>   

1. <span data-ttu-id="1b9be-173">[Microsoft Office Online](https://office.live.com/) を開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-173">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="1b9be-174">新しい Excel ブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="1b9be-175">リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-175">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="1b9be-176">**[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-176">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="1b9be-178">アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="1b9be-180">サイドロードしたドキュメントは、ドキュメントを開くたびにサイドロードされたままになります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="1b9be-181">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="1b9be-181">Start debugging</span></span>

1. <span data-ttu-id="1b9be-182">開発者ツールをブラウザーで開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-182">Open developer tools in the browser.</span></span> <span data-ttu-id="1b9be-183">Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="1b9be-184">開発者ツールで、 **Cmd + p**または**Ctrl + p** (node.js または functions) を使用してソースコードスクリプトファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (functions.js or functions.ts).</span></span>
3. <span data-ttu-id="1b9be-185">カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="1b9be-186">コードを変更する必要がある場合は、VS コードで編集を行って変更を保存することができます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="1b9be-187">ブラウザーを更新して、変更が読み込まれたことを確認します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="1b9be-188">コマンドラインツールを使用してデバッグする</span><span class="sxs-lookup"><span data-stu-id="1b9be-188">Use the command line tools to debug</span></span>

<span data-ttu-id="1b9be-189">VS コードを使用していない場合は、コマンドライン (bash、PowerShell など) を使用してアドインを実行できます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="1b9be-190">Excel Online でコードをデバッグするには、ブラウザー開発者ツールを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-190">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="1b9be-191">コマンドラインを使用して、デスクトップ版の Excel をデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="1b9be-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="1b9be-192">コマンドラインからを実行`npm run watch`すると、コードの変更が発生したときにを監視し、再構築します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="1b9be-193">2番目のコマンドラインウィンドウを開きます (最初のウィンドウは、ウォッチの実行中にブロックされます)。</span><span class="sxs-lookup"><span data-stu-id="1b9be-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="1b9be-194">Excel のデスクトップバージョンでアドインを起動するには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start desktop`
    
    <span data-ttu-id="1b9be-195">または、Excel Online でアドインを起動したい場合は、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-195">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start web`
    
    <span data-ttu-id="1b9be-196">Excel Online の場合は、アドインをサイドロードする必要もあります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-196">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="1b9be-197">「[サイドロード](#sideload-your-add-in)を使用してアドインをサイドロードする」の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="1b9be-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="1b9be-198">その後、次のセクションに進み、デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="1b9be-199">開発者ツールをブラウザーで開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-199">Open developer tools in the browser.</span></span> <span data-ttu-id="1b9be-200">Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="1b9be-201">[開発者ツール] で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-201">In developer tools, open your source code script file (functions.js or functions.ts).</span></span> <span data-ttu-id="1b9be-202">カスタム関数のコードは、ファイルの末尾付近に配置されている場合があります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="1b9be-203">カスタム関数のソースコードで、コードの行を選択してブレークポイントを適用します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="1b9be-204">コードを変更する必要がある場合は、Visual Studio で編集を行って変更を保存することができます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="1b9be-205">ブラウザーを更新して、変更が読み込まれたことを確認します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="1b9be-206">アドインをビルドして実行するためのコマンド</span><span class="sxs-lookup"><span data-stu-id="1b9be-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="1b9be-207">使用可能なビルドタスクはいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="1b9be-207">There are several build tasks available:</span></span>
- <span data-ttu-id="1b9be-208">`npm run watch`: ソースファイルの保存時に開発用のビルドを作成し、自動的に再構築します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="1b9be-209">`npm run build-dev`: 開発用ビルド</span><span class="sxs-lookup"><span data-stu-id="1b9be-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="1b9be-210">`npm run build`: 運用のためのビルド</span><span class="sxs-lookup"><span data-stu-id="1b9be-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="1b9be-211">`npm run dev-server`: 開発に使用する web サーバーを実行します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="1b9be-212">次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="1b9be-213">`npm run start desktop`: デスクトップ上で Excel を起動し、アドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-213">`npm run start desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="1b9be-214">`npm run start web`: Excel Online を起動して、アドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1b9be-214">`npm run start web`: Starts Excel Online and sideloads your add-in.</span></span>
- <span data-ttu-id="1b9be-215">`npm run stop`: Excel およびデバッグを停止します。</span><span class="sxs-lookup"><span data-stu-id="1b9be-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="see-also"></a><span data-ttu-id="1b9be-216">関連項目</span><span class="sxs-lookup"><span data-stu-id="1b9be-216">See also</span></span>

* [<span data-ttu-id="1b9be-217">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="1b9be-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="1b9be-218">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="1b9be-218">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="1b9be-219">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1b9be-219">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="1b9be-220">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="1b9be-220">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="1b9be-221">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="1b9be-221">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
