---
ms.date: 04/12/2021
description: 作業ウィンドウを使用しないExcel関数をデバッグする方法について説明します。
title: UI レスのカスタム関数のデバッグ
localization_priority: Normal
ms.openlocfilehash: a692f376cb5c874fa4d510d3459469d803e643f7
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075937"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="06cb7-103">UI レスのカスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="06cb7-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="06cb7-104">この記事では、作業ウィンドウまたは他のユーザー インターフェイス要素 (UI レスのカスタム関数) を使用しないカスタム関数のデバッグのみについて説明します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="06cb7-105">オンWindows:</span><span class="sxs-lookup"><span data-stu-id="06cb7-105">On Windows:</span></span>
- [<span data-ttu-id="06cb7-106">ExcelデスクトップおよびVisual Studio Code (VS Code) デバッガー</span><span class="sxs-lookup"><span data-stu-id="06cb7-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="06cb7-107">Excel on the webとVS Codeデバッガー</span><span class="sxs-lookup"><span data-stu-id="06cb7-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="06cb7-108">Excel on the webブラウザー ツール</span><span class="sxs-lookup"><span data-stu-id="06cb7-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="06cb7-109">コマンド ライン</span><span class="sxs-lookup"><span data-stu-id="06cb7-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="06cb7-110">Mac の場合:</span><span class="sxs-lookup"><span data-stu-id="06cb7-110">On Mac:</span></span>
- [<span data-ttu-id="06cb7-111">Excel on the webブラウザー ツール</span><span class="sxs-lookup"><span data-stu-id="06cb7-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="06cb7-112">コマンド ライン</span><span class="sxs-lookup"><span data-stu-id="06cb7-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="06cb7-113">わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグを示します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="06cb7-114">別のエディターまたはコマンド ライン ツールを使用している場合は[](#commands-for-building-and-running-your-add-in)、この記事の最後にあるコマンド ラインの手順を参照してください。</span><span class="sxs-lookup"><span data-stu-id="06cb7-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="06cb7-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="06cb7-115">Requirements</span></span>

<span data-ttu-id="06cb7-116">このデバッグ プロセスは、作業 **ウィンドウ** や他の UI 要素を使用しない UI レスのカスタム関数でのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="06cb7-117">UI レスのカスタム関数を作成するには、「Excel のカスタム関数を作成する」チュートリアルの手順に従い[、Office](../tutorials/excel-tutorial-create-custom-functions.md)アドイン用の[Yeoman](https://www.npmjs.com/package/generator-office)ジェネレーターによってインストールされている作業ウィンドウと UI 要素をすべて削除します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="06cb7-118">このデバッグ プロセスは、共有ランタイムを使用するカスタム関数プロジェクトと [互換性がない点に注意してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="06cb7-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="06cb7-119">デスクトップにVS CodeデバッガーをExcelする</span><span class="sxs-lookup"><span data-stu-id="06cb7-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="06cb7-120">デスクトップ上VS Code UI レスカスタム関数をデバッグするには、Office Excelを使用します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="06cb7-121">Mac のデスクトップ デバッグは使用できませんが、ブラウザー ツールとコマンド ラインを使用してデバッグ[Excel on the web)。](#use-the-command-line-tools-to-debug)</span><span class="sxs-lookup"><span data-stu-id="06cb7-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="06cb7-122">アドインを実行するには、次のVS Code</span><span class="sxs-lookup"><span data-stu-id="06cb7-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="06cb7-123">カスタム関数ルート プロジェクト フォルダーを開きます[。VS Code。](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="06cb7-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="06cb7-124">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="06cb7-125">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="06cb7-126">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="06cb7-127">デバッガーのVS Codeする</span><span class="sxs-lookup"><span data-stu-id="06cb7-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="06cb7-128">[ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="06cb7-129">[実行] ドロップダウン メニューから、[デスクトップ **(Excel関数) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-129">From the Run drop-down menu, choose **Excel Desktop (Custom Functions)**.</span></span>
6. <span data-ttu-id="06cb7-130">デバッグ **を開始するには、[F5]** を選択します ( **または>から** [デバッグの開始] を選択します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="06cb7-131">新しいExcelブックが開き、アドインが既にサイドロードされ、すぐに使用できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="06cb7-132">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="06cb7-132">Start debugging</span></span>

1. <span data-ttu-id="06cb7-133">このVS Code、ソース コード スクリプト ファイル (functions.jsまたは **functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="06cb7-134">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="06cb7-135">ブックのExcel、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="06cb7-136">この時点で、ブレークポイントを設定したコード行で実行が停止します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="06cb7-137">これで、コードをステップ実行し、ウォッチを設定し、必要なデバッグVS Codeを使用できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="06cb7-138">アプリケーション内のVS CodeデバッガーをExcel使用Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="06cb7-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="06cb7-139">このコマンドを使用VS Code、UI レスのカスタム関数を、Excelブラウザー Microsoft Edgeできます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="06cb7-140">この機能をVS CodeするにはMicrosoft Edge拡張機能用の[デバッガーをインストールMicrosoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)があります。</span><span class="sxs-lookup"><span data-stu-id="06cb7-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="06cb7-141">アドインを実行するには、次のVS Code</span><span class="sxs-lookup"><span data-stu-id="06cb7-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="06cb7-142">カスタム関数ルート プロジェクト フォルダーを開きます[。VS Code。](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="06cb7-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="06cb7-143">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="06cb7-144">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="06cb7-145">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="06cb7-146">デバッガーのVS Codeする</span><span class="sxs-lookup"><span data-stu-id="06cb7-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="06cb7-147">[ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="06cb7-148">[デバッグ] オプションで、[オンライン] **Office (エッジ Chromium) を選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="06cb7-149">ブラウザー Excel開Microsoft Edge新しいブックを作成します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="06cb7-150">リボン **で [共有** ] を選択し、この新しいブックの URL のリンクをコピーします。</span><span class="sxs-lookup"><span data-stu-id="06cb7-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="06cb7-151">デバッグ **を開始するには、[F5]** **(または>[** デバッグの開始] を選択します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="06cb7-152">ドキュメントの URL を求めるプロンプトが表示されます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="06cb7-153">ブックの URL に貼り付け、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="06cb7-154">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="06cb7-154">Sideload your add-in</span></span>

1. <span data-ttu-id="06cb7-155">リボンの **[挿入**] タブを選択し、[アドイン] セクションで、[アドイン] Office **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="06cb7-156">[アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="06cb7-158">**アドイン** マニフェスト ファイルを参照し、[次へ] を **アップロード。**</span><span class="sxs-lookup"><span data-stu-id="06cb7-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="06cb7-160">ブレークポイントの設定</span><span class="sxs-lookup"><span data-stu-id="06cb7-160">Set breakpoints</span></span>
1. <span data-ttu-id="06cb7-161">このVS Code、ソース コード スクリプト ファイル (functions.jsまたは **functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="06cb7-162">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="06cb7-163">ブックのExcel、カスタム関数を使用する数式を入力します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="06cb7-164">ブラウザー開発者ツールを使用して、ブラウザーのカスタム関数をデバッグExcel on the web</span><span class="sxs-lookup"><span data-stu-id="06cb7-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="06cb7-165">ブラウザー開発者ツールを使用して、UI レスのカスタム関数をデバッグExcel on the web。</span><span class="sxs-lookup"><span data-stu-id="06cb7-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="06cb7-166">次の手順は、Windows macOS の両方で機能します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="06cb7-167">アドインをアプリから実行Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="06cb7-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="06cb7-168">カスタム関数ルート プロジェクト フォルダーを Visual Studio Code [(VS Code) で開きます](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="06cb7-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="06cb7-169">[ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="06cb7-170">これにより、ファイルの変更が監視され、再構築されます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="06cb7-171">[ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="06cb7-172">アドインのサイドロード</span><span class="sxs-lookup"><span data-stu-id="06cb7-172">Sideload your add-in</span></span>

1. <span data-ttu-id="06cb7-173">[ファイル[Office on the web] を開きます](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="06cb7-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="06cb7-174">新しいブックを開Excelします。</span><span class="sxs-lookup"><span data-stu-id="06cb7-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="06cb7-175">リボンの **[挿入**] タブを開き、[アドイン] セクションで、[アドイン] Office **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="06cb7-176">[アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="06cb7-178">アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="06cb7-180">ドキュメントにサイドロードすると、ドキュメントを開くごとにサイドロードされたままです。</span><span class="sxs-lookup"><span data-stu-id="06cb7-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="06cb7-181">デバッグを開始する</span><span class="sxs-lookup"><span data-stu-id="06cb7-181">Start debugging</span></span>

1. <span data-ttu-id="06cb7-182">ブラウザーで開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-182">Open developer tools in the browser.</span></span> <span data-ttu-id="06cb7-183">Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="06cb7-184">開発者ツールで **、Cmd + P** または Ctrl + **P** (functions.jsまたは **functions.ts)** を **使用して** ソース コード スクリプト ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="06cb7-185">[カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="06cb7-186">コードを変更する必要がある場合は、編集を行い、VS Code保存できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="06cb7-187">ブラウザーを更新して、読み込まれた変更を確認します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="06cb7-188">コマンド ライン ツールを使用してデバッグする</span><span class="sxs-lookup"><span data-stu-id="06cb7-188">Use the command line tools to debug</span></span>

<span data-ttu-id="06cb7-189">アドインを使用していないVS Code、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="06cb7-190">ブラウザー開発者ツールを使用して、ブラウザーのコードをデバッグする必要Excel on the web。</span><span class="sxs-lookup"><span data-stu-id="06cb7-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="06cb7-191">コマンド ラインを使用してデスクトップ バージョンExcelデバッグすることはできません。</span><span class="sxs-lookup"><span data-stu-id="06cb7-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="06cb7-192">コマンド ラインから実行して `npm run watch` 、コードの変更が発生した場合の監視と再構築を行います。</span><span class="sxs-lookup"><span data-stu-id="06cb7-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="06cb7-193">2 番目のコマンド ライン ウィンドウを開きます (最初のウィンドウはウォッチの実行中にブロックされます)。</span><span class="sxs-lookup"><span data-stu-id="06cb7-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="06cb7-194">デスクトップ バージョンのデスクトップ バージョンでアドインを起動する場合はExcelコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="06cb7-195">または、次のコマンドを実行するためにアドインExcel on the web開始する場合</span><span class="sxs-lookup"><span data-stu-id="06cb7-195">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="06cb7-196">このExcel on the webアドインをサイドロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="06cb7-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="06cb7-197">「アドインを [サイドロードする」の手順に従って](#sideload-your-add-in) 、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="06cb7-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="06cb7-198">次に、次のセクションに進み、デバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="06cb7-199">ブラウザーで開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-199">Open developer tools in the browser.</span></span> <span data-ttu-id="06cb7-200">Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="06cb7-201">開発者ツールで、ソース コード スクリプト ファイル(functions.jsまたは **functions.ts) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="06cb7-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="06cb7-202">カスタム関数コードは、ファイルの末尾近くに位置している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="06cb7-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="06cb7-203">カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="06cb7-204">コードを変更する必要がある場合は、編集を行い、Visual Studio保存できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="06cb7-205">ブラウザーを更新して、読み込まれた変更を確認します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="06cb7-206">アドインを構築および実行するコマンド</span><span class="sxs-lookup"><span data-stu-id="06cb7-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="06cb7-207">使用できるビルド タスクは複数あります。</span><span class="sxs-lookup"><span data-stu-id="06cb7-207">There are several build tasks available:</span></span>
- <span data-ttu-id="06cb7-208">`npm run watch`: 開発用のビルドと、ソース ファイルの保存時に自動的に再構築する</span><span class="sxs-lookup"><span data-stu-id="06cb7-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="06cb7-209">`npm run build-dev`: 一度開発用にビルドする</span><span class="sxs-lookup"><span data-stu-id="06cb7-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="06cb7-210">`npm run build`: 実稼働用のビルド</span><span class="sxs-lookup"><span data-stu-id="06cb7-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="06cb7-211">`npm run dev-server`: 開発に使用する Web サーバーを実行します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="06cb7-212">次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。</span><span class="sxs-lookup"><span data-stu-id="06cb7-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="06cb7-213">`npm run start:desktop`: デスクトップExcelを開始し、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="06cb7-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="06cb7-214">`npm run start:web`: アドインExcel on the webを開始し、サイドロードします。</span><span class="sxs-lookup"><span data-stu-id="06cb7-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="06cb7-215">`npm run stop`: デバッグExcel停止します。</span><span class="sxs-lookup"><span data-stu-id="06cb7-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="06cb7-216">次の手順</span><span class="sxs-lookup"><span data-stu-id="06cb7-216">Next steps</span></span>
<span data-ttu-id="06cb7-217">UI レス [のカスタム関数の認証方法について説明します](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="06cb7-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="06cb7-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="06cb7-218">See also</span></span>

* [<span data-ttu-id="06cb7-219">カスタム関数のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="06cb7-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="06cb7-220">Excel のカスタム関数でのエラー処理 </span><span class="sxs-lookup"><span data-stu-id="06cb7-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="06cb7-221">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="06cb7-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
