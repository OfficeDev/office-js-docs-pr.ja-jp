---
ms.date: 07/10/2020
description: 作業ウィンドウを使用しない Excel カスタム関数をデバッグする方法について説明します。
title: UI レスのカスタム関数のデバッグ
localization_priority: Normal
ms.openlocfilehash: 00065a465a22f83891dfb207943102b079e96a0f
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178077"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="65da9-103">UI レスのカスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="65da9-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="65da9-104">作業ウィンドウまたは他のユーザー インターフェイス要素 (UI レスのカスタム関数) を使用しないカスタム関数のデバッグは、使用しているプラットフォームに応じて複数の方法で実行できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-104">Debugging for custom functions that don't use a task pane or other user interface elements (UI-less custom functions) can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="65da9-105">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="65da9-105">On Windows:</span></span>
- [<span data-ttu-id="65da9-106">Excel Desktop および Visual Studio コード (VS Code) デバッガー</span><span class="sxs-lookup"><span data-stu-id="65da9-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="65da9-107">Excel on the web and VS Code Debugger</span><span class="sxs-lookup"><span data-stu-id="65da9-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="65da9-108">Excel on the web and browser tools</span><span class="sxs-lookup"><span data-stu-id="65da9-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="65da9-109">コマンド ライン</span><span class="sxs-lookup"><span data-stu-id="65da9-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="65da9-110">Mac の場合:</span><span class="sxs-lookup"><span data-stu-id="65da9-110">On Mac:</span></span>
- [<span data-ttu-id="65da9-111">Excel on the web and browser tools</span><span class="sxs-lookup"><span data-stu-id="65da9-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="65da9-112">コマンド ライン</span><span class="sxs-lookup"><span data-stu-id="65da9-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="65da9-113">わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグを示します。</span><span class="sxs-lookup"><span data-stu-id="65da9-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="65da9-114">別のエディターまたはコマンド ライン ツールを使用している場合は[](#commands-for-building-and-running-your-add-in)、この記事の最後にあるコマンド ラインの手順を参照してください。</span><span class="sxs-lookup"><span data-stu-id="65da9-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="65da9-115">要件</span><span class="sxs-lookup"><span data-stu-id="65da9-115">Requirements</span></span>

<span data-ttu-id="65da9-116">デバッグを開始する前に [、Yeoman](https://github.com/OfficeDev/generator-office) ジェネレーターを使用Officeカスタム関数プロジェクトを作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="65da9-116">Before starting to debug, you should use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="65da9-117">カスタム関数プロジェクトを作成する方法のガイダンスについては、カスタム関数の [チュートリアルを参照してください](../tutorials/excel-tutorial-create-custom-functions.md)。</span><span class="sxs-lookup"><span data-stu-id="65da9-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="65da9-118">Excel Desktop で VS Code デバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="65da9-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="65da9-119">VS Code を使用すると、デスクトップ上の Excel で UI レスのOfficeをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="65da9-119">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="65da9-120">Mac のデスクトップ デバッグは使用できませんが、ブラウザー ツールとコマンド ラインを使用して Web 上の [Excel をデバッグできます](#use-the-command-line-tools-to-debug))。</span><span class="sxs-lookup"><span data-stu-id="65da9-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="65da9-121">VS Code からアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="65da9-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="65da9-122">VS Code でカスタム関数ルート プロジェクト フォルダー [を開きます](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="65da9-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="65da9-123">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="65da9-124">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="65da9-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="65da9-125">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="65da9-126">VS Code デバッガーを起動する</span><span class="sxs-lookup"><span data-stu-id="65da9-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="65da9-127">[ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="65da9-127">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="65da9-128">[実行] ドロップダウン メニューから **、[Excel デスクトップ ] (エッジ クロム) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-128">From the Run drop-down menu, choose **Excel Desktop (Edge Chromium)**.</span></span>
6. <span data-ttu-id="65da9-129">デバッグ **を開始するには、[F5]** を選択します ( **または>から** [デバッグの開始] を選択します。</span><span class="sxs-lookup"><span data-stu-id="65da9-129">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="65da9-130">アドインが既にサイドロードされ、すぐに使用できる状態で、新しい Excel ブックが開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="65da9-131">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="65da9-131">Start debugging</span></span>

1. <span data-ttu-id="65da9-132">VS Code で、ソース コード スクリプトファイル (functions.js **または functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="65da9-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="65da9-133">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="65da9-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="65da9-134">Excel ブックに、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="65da9-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="65da9-135">この時点で、ブレークポイントを設定したコード行で実行が停止します。</span><span class="sxs-lookup"><span data-stu-id="65da9-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="65da9-136">これで、コードをステップ実行し、ウォッチを設定し、必要な VS Code デバッグ機能を使用できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="65da9-137">Microsoft Edge で Excel の VS Code デバッガーを使用する</span><span class="sxs-lookup"><span data-stu-id="65da9-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="65da9-138">VS Code を使用すると、Microsoft Edge ブラウザーの Excel で UI レスのカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="65da9-138">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="65da9-139">Microsoft Edge で VS Code を使用するには、デバッガー for [Microsoft Edge 拡張機能をインストールする必要](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) があります。</span><span class="sxs-lookup"><span data-stu-id="65da9-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="65da9-140">VS Code からアドインを実行する</span><span class="sxs-lookup"><span data-stu-id="65da9-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="65da9-141">VS Code でカスタム関数ルート プロジェクト フォルダー [を開きます](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="65da9-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="65da9-142">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="65da9-143">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="65da9-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="65da9-144">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="65da9-145">VS Code デバッガーを起動する</span><span class="sxs-lookup"><span data-stu-id="65da9-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="65da9-146">[ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="65da9-146">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="65da9-147">[デバッグ] オプションで、[オンライン] **Office (Edge Chromium) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-147">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="65da9-148">Microsoft Edge ブラウザーで Excel を開き、新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="65da9-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="65da9-149">リボン **で [共有** ] を選択し、この新しいブックの URL のリンクをコピーします。</span><span class="sxs-lookup"><span data-stu-id="65da9-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="65da9-150">デバッグ **を開始するには、[F5]** **(または>[** デバッグの開始] を選択します。</span><span class="sxs-lookup"><span data-stu-id="65da9-150">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="65da9-151">ドキュメントの URL を求めるプロンプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="65da9-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="65da9-152">ブックの URL に貼り付け、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="65da9-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="65da9-153">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="65da9-153">Sideload your add-in</span></span>

1. <span data-ttu-id="65da9-154">リボンの **[挿入**] タブを選択し、[アドイン] セクションで、[アドイン] Office **選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="65da9-155">[アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[マイアドインの管理] を選択し、[自分のアドインのアップロード]**を選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="65da9-157">**アドイン** マニフェスト ファイルを参照し、[アップロード] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="65da9-159">ブレークポイントの設定</span><span class="sxs-lookup"><span data-stu-id="65da9-159">Set breakpoints</span></span>
1. <span data-ttu-id="65da9-160">VS Code で、ソース コード スクリプトファイル (functions.js **または functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="65da9-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="65da9-161">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="65da9-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="65da9-162">Excel ブックに、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="65da9-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="65da9-163">ブラウザー開発者ツールを使用して、Web 上の Excel でカスタム関数をデバッグする</span><span class="sxs-lookup"><span data-stu-id="65da9-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="65da9-164">ブラウザー開発者ツールを使用して、Web 上の Excel で UI レスのカスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="65da9-164">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="65da9-165">次の手順は、Windows と macOS の両方で機能します。</span><span class="sxs-lookup"><span data-stu-id="65da9-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="65da9-166">コードからアドインをVisual Studioする</span><span class="sxs-lookup"><span data-stu-id="65da9-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="65da9-167">カスタム関数ルート プロジェクト フォルダーを [コード] [(VS Code) Visual Studio開きます](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="65da9-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="65da9-168">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="65da9-169">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="65da9-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="65da9-170">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="65da9-171">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="65da9-171">Sideload your add-in</span></span>

1. <span data-ttu-id="65da9-172">Web [Officeを開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="65da9-172">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="65da9-173">新しい Excel ブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="65da9-174">リボンの **[挿入**] タブを開き、[アドイン] セクションで、[アドイン] Office **選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="65da9-175">[アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[マイアドインの管理] を選択し、[自分のアドインのアップロード]**を選択します**。</span><span class="sxs-lookup"><span data-stu-id="65da9-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="65da9-177">アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="65da9-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="65da9-179">ドキュメントにサイドロードすると、ドキュメントを開くごとにサイドロードされたままです。</span><span class="sxs-lookup"><span data-stu-id="65da9-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="65da9-180">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="65da9-180">Start debugging</span></span>

1. <span data-ttu-id="65da9-181">ブラウザーで開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-181">Open developer tools in the browser.</span></span> <span data-ttu-id="65da9-182">Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="65da9-183">開発者ツールで **、Cmd + P** または Ctrl + **P** (functions.jsまたは **functions.ts)** を **使用して** ソース コード スクリプト ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="65da9-184">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="65da9-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="65da9-185">コードを変更する必要がある場合は、VS Code で編集を行い、変更を保存できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="65da9-186">ブラウザーを更新して、読み込まれた変更を確認します。</span><span class="sxs-lookup"><span data-stu-id="65da9-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="65da9-187">コマンド ライン ツールを使用してデバッグする</span><span class="sxs-lookup"><span data-stu-id="65da9-187">Use the command line tools to debug</span></span>

<span data-ttu-id="65da9-188">VS Code を使用していない場合は、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="65da9-189">ブラウザー開発者ツールを使用して、Web 上の Excel でコードをデバッグする必要があります。</span><span class="sxs-lookup"><span data-stu-id="65da9-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="65da9-190">コマンド ラインを使用してデスクトップ バージョンの Excel をデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="65da9-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="65da9-191">コマンド ラインから実行して `npm run watch` 、コードの変更が発生した場合の監視と再構築を行います。</span><span class="sxs-lookup"><span data-stu-id="65da9-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="65da9-192">2 番目のコマンド ライン ウィンドウを開きます (最初のウィンドウはウォッチの実行中にブロックされます)。</span><span class="sxs-lookup"><span data-stu-id="65da9-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="65da9-193">デスクトップ バージョンの Excel でアドインを起動する場合は、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="65da9-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="65da9-194">または、Web 上の Excel でアドインを起動する場合は、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="65da9-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="65da9-195">Web 上の Excel の場合は、アドインをサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="65da9-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="65da9-196">「アドインを [サイドロードする」の手順に従って](#sideload-your-add-in) 、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="65da9-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="65da9-197">次に、次のセクションに進み、デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="65da9-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="65da9-198">ブラウザーで開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-198">Open developer tools in the browser.</span></span> <span data-ttu-id="65da9-199">Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="65da9-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="65da9-200">開発者ツールで、ソース コード スクリプト ファイル(functions.jsまたは **functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="65da9-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="65da9-201">カスタム関数コードは、ファイルの末尾近くに位置している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="65da9-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="65da9-202">カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。</span><span class="sxs-lookup"><span data-stu-id="65da9-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="65da9-203">コードを変更する必要がある場合は、編集を行い、Visual Studio保存できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="65da9-204">ブラウザーを更新して、読み込まれた変更を確認します。</span><span class="sxs-lookup"><span data-stu-id="65da9-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="65da9-205">アドインを構築および実行するコマンド</span><span class="sxs-lookup"><span data-stu-id="65da9-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="65da9-206">使用できるビルド タスクは複数あります。</span><span class="sxs-lookup"><span data-stu-id="65da9-206">There are several build tasks available:</span></span>
- <span data-ttu-id="65da9-207">`npm run watch`: 開発用のビルドと、ソース ファイルの保存時に自動的に再構築する</span><span class="sxs-lookup"><span data-stu-id="65da9-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="65da9-208">`npm run build-dev`: 一度開発用にビルドする</span><span class="sxs-lookup"><span data-stu-id="65da9-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="65da9-209">`npm run build`: 実稼働用のビルド</span><span class="sxs-lookup"><span data-stu-id="65da9-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="65da9-210">`npm run dev-server`: 開発に使用する Web サーバーを実行します。</span><span class="sxs-lookup"><span data-stu-id="65da9-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="65da9-211">次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。</span><span class="sxs-lookup"><span data-stu-id="65da9-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="65da9-212">`npm run start:desktop`: デスクトップ上で Excel を起動し、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="65da9-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="65da9-213">`npm run start:web`: Web 上で Excel を起動し、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="65da9-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="65da9-214">`npm run stop`: Excel とデバッグを停止します。</span><span class="sxs-lookup"><span data-stu-id="65da9-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="65da9-215">次の手順</span><span class="sxs-lookup"><span data-stu-id="65da9-215">Next steps</span></span>
<span data-ttu-id="65da9-216">UI レス [のカスタム関数の認証方法について説明します](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="65da9-216">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="65da9-217">関連項目</span><span class="sxs-lookup"><span data-stu-id="65da9-217">See also</span></span>

* [<span data-ttu-id="65da9-218">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="65da9-218">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="65da9-219">Excel のカスタム関数でのエラー処理 </span><span class="sxs-lookup"><span data-stu-id="65da9-219">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="65da9-220">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="65da9-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
