---
title: Microsoft Edge WebView2 (Chromium ベース) を使用した Windows 上のアドインをデバッグする
description: VS Code で拡張機能 Debugger for Microsoft Edge を使用し、Microsoft Edge WebView2 (Chromium ベース) を使用した Office アドインをデバッグする方法について説明します。
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 6a62718147fbb5d2e8a6819066425737d853cbf0
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350177"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a><span data-ttu-id="23a50-103">Edge Chromium WebView2 を使用して Windows でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="23a50-103">Debug add-ins on Windows using Edge Chromium WebView2</span></span>

<span data-ttu-id="23a50-104">Windows 上で動作する Office アドインは、VS Code の拡張機能 Debugger for Microsoft Edge を使用することで、Edge Chromium WebView2 ランタイムに対してデバッグを行うことができます。</span><span class="sxs-lookup"><span data-stu-id="23a50-104">Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="23a50-105">前提条件</span><span class="sxs-lookup"><span data-stu-id="23a50-105">Prerequisites</span></span>

- <span data-ttu-id="23a50-106">[Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)</span><span class="sxs-lookup"><span data-stu-id="23a50-106">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="23a50-107">Node.js (バージョン 10 以上)</span><span class="sxs-lookup"><span data-stu-id="23a50-107">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="23a50-108">Windows 10</span><span class="sxs-lookup"><span data-stu-id="23a50-108">Windows 10</span></span>
- [<span data-ttu-id="23a50-109">Microsoft Edge Chromium は Windows Insider に提供しています</span><span class="sxs-lookup"><span data-stu-id="23a50-109">Microsoft Edge Chromium available to Windows Insiders</span></span>](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="23a50-110">デバッガーをインストールして使用する</span><span class="sxs-lookup"><span data-stu-id="23a50-110">Install and use the debugger</span></span>

1. <span data-ttu-id="23a50-111">[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用してプロジェクトを作成してください。これを行うには、「[Outlook アドインのクイック スタート](../quickstarts/outlook-quickstart.md)」などのクイック スタート ガイドのいずれかをご利用ください。</span><span class="sxs-lookup"><span data-stu-id="23a50-111">Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.</span></span>

    > [!TIP]
    > <span data-ttu-id="23a50-112">Yeoman ジェネレーター ベースのアドインを使用していない場合は、レジストリ キーを調整する必要があります。</span><span class="sxs-lookup"><span data-stu-id="23a50-112">If you aren't using a Yeoman generator based add-in, you need to adjust a registry key.</span></span> <span data-ttu-id="23a50-113">プロジェクトのルート フォルダーで、コマンドラインを使用して以下を実行します: `office-add-in-debugging start <your manifest path>`。</span><span class="sxs-lookup"><span data-stu-id="23a50-113">While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.</span></span>

1. <span data-ttu-id="23a50-114">VS Code でプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="23a50-114">Open your project in VS Code.</span></span> <span data-ttu-id="23a50-115">VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。</span><span class="sxs-lookup"><span data-stu-id="23a50-115">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="23a50-116">「Debugger for Microsoft Edge」で拡張機能を検索し、これをインストールします。</span><span class="sxs-lookup"><span data-stu-id="23a50-116">Search for the "Debugger for Microsoft Edge" extension and install it.</span></span>

1. <span data-ttu-id="23a50-117">プロジェクトの **.vscode** フォルダーで、**launch.json** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="23a50-117">In the **.vscode** folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="23a50-118">構成セクションに以下のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="23a50-118">Add the following code to the configurations section.</span></span>

      ```JSON
        {
          "name": "Debug Office Add-in (Edge Chromium)",
          "type": "edge",
          "request": "attach",
          "useWebView": "advanced",
          "port": 9229,
          "timeout": 600000,
          "webRoot": "${workspaceRoot}",
        },
      ```

1. <span data-ttu-id="23a50-119">次に、**[表示]、[デバッグ]** の順に選択するか、**Ctrl キー + Shift キー + D キー** を入力してデバッグ ビューに切り替えます。</span><span class="sxs-lookup"><span data-stu-id="23a50-119">Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

1. <span data-ttu-id="23a50-120">デバッグ オプションから、**Excel Desktop (Edge Chromium)** などのホスト アプリケーション用に Edge Chromium オプションを選択します。</span><span class="sxs-lookup"><span data-stu-id="23a50-120">From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**.</span></span> <span data-ttu-id="23a50-121">**F5** キーを選択するか、メニューから **[デバッグ]、[デバッグの開始]** の順に選択してデバッグを開始します。</span><span class="sxs-lookup"><span data-stu-id="23a50-121">Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.</span></span>

1. <span data-ttu-id="23a50-122">これで、Excel などのホスト アプリケーションでアドインを使用する準備ができました。</span><span class="sxs-lookup"><span data-stu-id="23a50-122">In the host application, such as Excel, your add-in is now ready to use.</span></span> <span data-ttu-id="23a50-123">**[作業ウィンドウの表示]** を選択するか、その他のアドイン コマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="23a50-123">Select **Show Taskpane** or run any other add-in command.</span></span> <span data-ttu-id="23a50-124">ダイアログ ボックスが表示され、以下が表示されます。</span><span class="sxs-lookup"><span data-stu-id="23a50-124">A dialog box will appear, reading:</span></span>

    > <span data-ttu-id="23a50-125">WebView は読み込み時に停止します。</span><span class="sxs-lookup"><span data-stu-id="23a50-125">WebView Stop On Load.</span></span>
    > <span data-ttu-id="23a50-126">WebView をデバッグするには、拡張機能 Microsoft Debugger for Edge を使用して VS Code を WebView のインスタンスにアタッチし、[OK] をクリックして続行します。</span><span class="sxs-lookup"><span data-stu-id="23a50-126">To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue.</span></span> <span data-ttu-id="23a50-127">今後このダイアログが表示されないようにするには、[キャンセル] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="23a50-127">To prevent this dialog from appearing in the future, click Cancel."</span></span>

    <span data-ttu-id="23a50-128">**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="23a50-128">Select **OK**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="23a50-129">**[キャンセル]** を選択すると、このアドインのインスタンスの実行中はダイアログが表示されなくなります。</span><span class="sxs-lookup"><span data-stu-id="23a50-129">If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running.</span></span> <span data-ttu-id="23a50-130">ただし、アドインを再起動すると、ダイアログはもう一度表示されます。</span><span class="sxs-lookup"><span data-stu-id="23a50-130">However, if you restart your add-in, you'll see the dialog again.</span></span>

1. <span data-ttu-id="23a50-131">これで、プロジェクトのコードにブレークポイントを設定し、デバッグを実行できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="23a50-131">You're now able to set breakpoints in your project's code and debug.</span></span>

## <a name="see-also"></a><span data-ttu-id="23a50-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="23a50-132">See also</span></span>

- [<span data-ttu-id="23a50-133">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="23a50-133">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
- [<span data-ttu-id="23a50-134">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="23a50-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="23a50-135">作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="23a50-135">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)