---
title: Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能
description: アドイン デバッガー Visual Studioコード拡張機能Microsoft Office使用して、アドインのOfficeデバッグします。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 83791d5d60238288e3059809b8b8c02b1f4f768f
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840112"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="3f44c-103">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="3f44c-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="3f44c-104">コードMicrosoft Officeアドイン デバッガー拡張機能Visual Studioでは、Edge ランタイムに対して Office アドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="3f44c-105">このデバッグ モードは動的で、コードの実行中にブレークポイントを設定できます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="3f44c-106">デバッガーがアタッチされている間は、デバッグ セッションを失わずに、コードの変更をすぐに確認できます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="3f44c-107">コードの変更も保持されます。そのため、コードに対する複数の変更の結果を確認できます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="3f44c-108">次の図は、この拡張機能の動作を示しています。</span><span class="sxs-lookup"><span data-stu-id="3f44c-108">The following image shows this extension in action.</span></span>

![Officeアドインのセクションをデバッグするアドイン デバッガー拡張機能](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="3f44c-110">前提条件</span><span class="sxs-lookup"><span data-stu-id="3f44c-110">Prerequisites</span></span>

- <span data-ttu-id="3f44c-111">[Visual Studioコード](https://code.visualstudio.com/) (管理者として実行する必要があります)</span><span class="sxs-lookup"><span data-stu-id="3f44c-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="3f44c-112">Node.js (バージョン 10 以上)</span><span class="sxs-lookup"><span data-stu-id="3f44c-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="3f44c-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="3f44c-113">Windows 10</span></span>
- [<span data-ttu-id="3f44c-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="3f44c-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="3f44c-115">これらの手順は、コマンド ラインの使用経験、基本的な JavaScript の理解、Yo Office ジェネレーターを使用する前に Office アドイン プロジェクトを作成した経験を前提にしています。</span><span class="sxs-lookup"><span data-stu-id="3f44c-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="3f44c-116">まだこれを行っていない場合は、次の [Excel](../tutorials/excel-tutorial.md)やアドインのチュートリアルOfficeチュートリアルのいずれかを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3f44c-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="3f44c-117">デバッガーをインストールして使用する</span><span class="sxs-lookup"><span data-stu-id="3f44c-117">Install and use the debugger</span></span>

1. <span data-ttu-id="3f44c-118">アドイン プロジェクトを作成する必要がある場合は、Yo Office ジェネレーターを使用 [して作成します](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。</span><span class="sxs-lookup"><span data-stu-id="3f44c-118">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="3f44c-119">コマンド ライン内のプロンプトに従って、プロジェクトを設定します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="3f44c-120">必要に応じて、任意の言語または種類のプロジェクトを選択できます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-120">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="3f44c-121">プロジェクトが既に存在する場合は、手順 1 をスキップして手順 2 に進みます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-121">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="3f44c-122">管理者としてコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="3f44c-123">![Windows 10 のコマンド プロンプト オプション ("管理者として実行" を含む)](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="3f44c-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="3f44c-124">プロジェクト ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="3f44c-125">次のコマンドを実行して、管理者として Visual Studioコードでプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="3f44c-126">コードVisual Studio開いた後、手動でプロジェクト フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="3f44c-127">管理者として Visual Studio コードを開く場合は、Windowsでコードを検索した後、Visual Studio コードを開く際に管理者として実行オプションを選択します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="3f44c-128">VS Code 内で **Ctrl + Shift + X** キーを押して拡張機能バーを開きます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="3f44c-129">"Microsoft Office アドイン デバッガー" 拡張機能を検索してインストールします。</span><span class="sxs-lookup"><span data-stu-id="3f44c-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="3f44c-130">プロジェクトの .vscode フォルダーで、ファイルlaunch.js **開** きます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="3f44c-131">セクションに次のコードを追加 `configurations` します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-131">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="3f44c-132">コピーした JSON のセクションで、"url" セクションを探します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="3f44c-133">この URL では、大文字の HOST テキストを、アドインをホストしているアプリケーションに置Officeがあります。</span><span class="sxs-lookup"><span data-stu-id="3f44c-133">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office add-in.</span></span> <span data-ttu-id="3f44c-134">たとえば、Office アドインが Excel 用の場合、URL 値は https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0" になります。</span><span class="sxs-lookup"><span data-stu-id="3f44c-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="3f44c-135">コマンド プロンプトを開き、プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="3f44c-136">コマンドを実行 `npm start` して開発サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="3f44c-137">アドインがクライアントに読み込Office作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="3f44c-138">コードにVisual Studioし、[デバッグ] の[>] を選択するか **、Ctrl + Shift + D** キーを押してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="3f44c-139">[デバッグ] オプションで、[アドイン **にアタッチOffice選択します**。F5 **キーを押** するか、[デバッグ] **->メニュー** から [デバッグの開始] を選択してデバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="3f44c-140">プロジェクトの作業ウィンドウ ファイルにブレークポイントを設定します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="3f44c-141">VS Code でブレークポイントを設定するには、コード行の横にカーソルを合わせると、表示される赤い円を選択します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![VS Code のコード行に赤い円が表示される](../images/set-breakpoint.jpg)

12. <span data-ttu-id="3f44c-143">アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="3f44c-143">Run your add-in.</span></span> <span data-ttu-id="3f44c-144">ブレークポイントにヒットしたと表示され、ローカル変数を検査できます。</span><span class="sxs-lookup"><span data-stu-id="3f44c-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="3f44c-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="3f44c-145">See also</span></span>

* [<span data-ttu-id="3f44c-146">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="3f44c-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="3f44c-147">Windows 10 で開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="3f44c-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="3f44c-148">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="3f44c-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)