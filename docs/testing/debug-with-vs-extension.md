---
title: Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能
description: アドイン デバッガー Visual Studio Code拡張機能Microsoft Office使用して、アドインのOfficeデバッグします。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 264a5d43a8b4f0faf7d6216664d30d7c8b64cccc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077121"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="2a390-103">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="2a390-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="2a390-104">Visual Studio Code の Microsoft Office アドイン デバッガー拡張機能を使用すると、Office アドインを元の webView (EdgeHTML) ランタイムで Microsoft Edge に対してデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="2a390-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="2a390-105">WebView2 に対するデバッグMicrosoft Edge (Chromiumベース) については、この[記事を参照してください。](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="2a390-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="2a390-106">このデバッグ モードは動的で、コードの実行中にブレークポイントを設定できます。</span><span class="sxs-lookup"><span data-stu-id="2a390-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="2a390-107">デバッガーが接続されている間、コード内の変更をすぐに確認できます。すべてデバッグ セッションを失う必要はありません。</span><span class="sxs-lookup"><span data-stu-id="2a390-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="2a390-108">コードの変更も保持されます。そのため、コードに対する複数の変更の結果を確認できます。</span><span class="sxs-lookup"><span data-stu-id="2a390-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="2a390-109">次の図は、この拡張機能の動作を示しています。</span><span class="sxs-lookup"><span data-stu-id="2a390-109">The following image shows this extension in action.</span></span>

![Officeアドイン デバッガー拡張機能は、アドインのセクションExcelデバッグします。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="2a390-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="2a390-111">Prerequisites</span></span>

- <span data-ttu-id="2a390-112">[Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)</span><span class="sxs-lookup"><span data-stu-id="2a390-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="2a390-113">Node.js (バージョン 10 以上)</span><span class="sxs-lookup"><span data-stu-id="2a390-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="2a390-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="2a390-114">Windows 10</span></span>
- [<span data-ttu-id="2a390-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="2a390-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="2a390-116">これらの手順では、コマンド ラインの使用経験、基本的な JavaScript の理解、および Yo Office ジェネレーターを使用する前に Office アドイン プロジェクトを作成したと仮定します。</span><span class="sxs-lookup"><span data-stu-id="2a390-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="2a390-117">前にこれを行ったことがない場合は、次のようなチュートリアルの 1 つを参照Excel Office[検討してください](../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="2a390-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="2a390-118">デバッガーをインストールして使用する</span><span class="sxs-lookup"><span data-stu-id="2a390-118">Install and use the debugger</span></span>

1. <span data-ttu-id="2a390-119">アドイン プロジェクトを作成する必要がある場合は[、Yo Officeを使用して作成します](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。</span><span class="sxs-lookup"><span data-stu-id="2a390-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="2a390-120">コマンド ライン内のプロンプトに従って、プロジェクトをセットアップします。</span><span class="sxs-lookup"><span data-stu-id="2a390-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="2a390-121">ニーズに合わせて任意の言語または種類のプロジェクトを選択できます。</span><span class="sxs-lookup"><span data-stu-id="2a390-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="2a390-122">プロジェクトが既に存在する場合は、手順 1 をスキップして、手順 2 に進みます。</span><span class="sxs-lookup"><span data-stu-id="2a390-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="2a390-123">管理者としてコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a390-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="2a390-124">![コマンド プロンプト のオプション ([管理者として実行] を含む) Windows 10。](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="2a390-124">![Command prompt options, including "run as administrator" in Windows 10.](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="2a390-125">プロジェクト ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="2a390-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="2a390-126">次のコマンドを実行して、プロジェクトを管理者Visual Studio Code開きます。</span><span class="sxs-lookup"><span data-stu-id="2a390-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="2a390-127">ファイルVisual Studio Code開いた後、手動でプロジェクト フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="2a390-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="2a390-128">管理者としてVisual Studio Codeを開く場合は、管理者として実行オプションを選択し、Visual Studio Codeで管理者を検索した後Windows。</span><span class="sxs-lookup"><span data-stu-id="2a390-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="2a390-129">VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a390-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="2a390-130">"Microsoft Office アドイン デバッガー" 拡張機能を検索してインストールします。</span><span class="sxs-lookup"><span data-stu-id="2a390-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="2a390-131">プロジェクトの .vscode フォルダーで、**launch.json** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a390-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="2a390-132">セクションに次のコードを追加 `configurations` します。</span><span class="sxs-lookup"><span data-stu-id="2a390-132">Add the following code to the `configurations` section:</span></span>

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. <span data-ttu-id="2a390-133">コピーした JSON のセクションで、"url" セクションを探します。</span><span class="sxs-lookup"><span data-stu-id="2a390-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="2a390-134">この URL では、大文字の HOST テキストを、アドインをホストしているアプリケーションに置き換Office必要があります。</span><span class="sxs-lookup"><span data-stu-id="2a390-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="2a390-135">たとえば、Office アドインが Excel 用の場合、URL 値は https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0" になります。</span><span class="sxs-lookup"><span data-stu-id="2a390-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="2a390-136">コマンド プロンプトを開き、プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="2a390-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="2a390-137">コマンドを実行 `npm start` して開発サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="2a390-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="2a390-138">アドインがクライアントに読み込まれるOffice作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="2a390-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="2a390-139">[デバッグ] **Visual Studio Codeし、[** デバッグの表示] を>、Ctrl + Shift **+ D** と入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="2a390-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="2a390-140">[デバッグ] オプションで、[**アドインに接続Office選択します**。**[F5]** を選択するか、メニューから **[デバッグ - >デバッグ** の開始] を選択してデバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="2a390-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="2a390-141">プロジェクトの作業ウィンドウ ファイルにブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="2a390-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="2a390-142">コード行の横にホバー Visual Studio Code表示される赤い円を選択すると、ブレークポイントを設定できます。</span><span class="sxs-lookup"><span data-stu-id="2a390-142">You can set breakpoints in Visual Studio Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![赤い円は、次のコード行にVisual Studio Code。](../images/set-breakpoint.jpg)

12. <span data-ttu-id="2a390-144">アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="2a390-144">Run your add-in.</span></span> <span data-ttu-id="2a390-145">ブレークポイントがヒットし、ローカル変数を検査できます。</span><span class="sxs-lookup"><span data-stu-id="2a390-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="2a390-146">関連項目</span><span class="sxs-lookup"><span data-stu-id="2a390-146">See also</span></span>

* [<span data-ttu-id="2a390-147">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="2a390-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="2a390-148">Windows 10 で開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="2a390-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="2a390-149">Microsoft Edge WebView2 (Chromium ベース) を使用した Windows 上のアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="2a390-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
