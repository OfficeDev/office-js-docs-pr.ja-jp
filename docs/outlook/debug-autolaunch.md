---
title: イベント ベースのOutlook アドインのデバッグ (プレビュー)
description: イベント ベースのアクティブ化を実装するOutlook アドインをデバッグする方法について説明します。
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555286"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="b1647-103">イベント ベースのOutlook アドインのデバッグ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b1647-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="b1647-104">この記事では、アドインにイベント [ベースのアクティブ化](autolaunch.md) を実装する際のデバッグ ガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b1647-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="b1647-105">イベントベースのアクティブ化機能は現在プレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="b1647-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1647-106">このデバッグ機能は、Microsoft 365 サブスクリプションを使用するWindowsでOutlookでプレビューする場合にのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="b1647-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="b1647-107">詳細については、この記事の「 [イベントベースのアクティブ化機能のプレビュー デバッグ](#preview-debugging-for-the-event-based-activation-feature) 」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b1647-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="b1647-108">この記事では、デバッグを有効にする主要なステージについて説明します。</span><span class="sxs-lookup"><span data-stu-id="b1647-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="b1647-109">アドインにデバッグ用のマークを付けます</span><span class="sxs-lookup"><span data-stu-id="b1647-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="b1647-110">Visual Studio Codeの構成</span><span class="sxs-lookup"><span data-stu-id="b1647-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="b1647-111">Visual Studio Codeを添付する</span><span class="sxs-lookup"><span data-stu-id="b1647-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="b1647-112">Debug</span><span class="sxs-lookup"><span data-stu-id="b1647-112">Debug</span></span>](#debug)

<span data-ttu-id="b1647-113">アドイン プロジェクトを作成するには、いくつかのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="b1647-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="b1647-114">使用するオプションによって、手順が異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="b1647-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="b1647-115">この場合、アドインの Office に Yeoman ジェネレーターを使用してアドイン プロジェクトを作成した場合 ([たとえば、イベント ベースのアクティブ化のチュートリアル](autolaunch.md)を実行する場合)、**ヨーオフィス** の手順に従って、それ以外の場合は「**その他** の手順」に従います。</span><span class="sxs-lookup"><span data-stu-id="b1647-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="b1647-116">Visual Studio Codeは、少なくともバージョン 1.56.1 である必要があります。</span><span class="sxs-lookup"><span data-stu-id="b1647-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="b1647-117">イベント ベースのアクティブ化機能のプレビュー デバッグ</span><span class="sxs-lookup"><span data-stu-id="b1647-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="b1647-118">イベントベースのアクティブ化機能のデバッグ機能を試してみるようお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b1647-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="b1647-119">GitHubを通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせください(このページの最後にある **フィードバック** セクションを参照)。</span><span class="sxs-lookup"><span data-stu-id="b1647-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="b1647-120">WindowsでOutlookに対してこの機能をプレビューするには、最小必要なビルドは 16.0.13729.2000 です。</span><span class="sxs-lookup"><span data-stu-id="b1647-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="b1647-121">ベータ版ビルドOfficeアクセスするには[、Office Insider プログラム](https://insider.office.com)に参加してください。</span><span class="sxs-lookup"><span data-stu-id="b1647-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="b1647-122">アドインにデバッグ用のマークを付けます</span><span class="sxs-lookup"><span data-stu-id="b1647-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="b1647-123">レジストリ キーを設定 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` する:</span><span class="sxs-lookup"><span data-stu-id="b1647-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="b1647-124">`[Add-in ID]` はアドイン マニフェストの **ID** です。</span><span class="sxs-lookup"><span data-stu-id="b1647-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="b1647-125">**yo office**: コマンド ライン ウィンドウで、アドイン フォルダーのルートに移動し、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="b1647-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="b1647-126">このコマンドでは、コードのビルドとローカル サーバーの起動に加えて、 `UseDirectDebugger` このアドインのレジストリ キーを に設定する必要があります `1` 。</span><span class="sxs-lookup"><span data-stu-id="b1647-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="b1647-127">**その他**: レジストリ キーを `UseDirectDebugger` に追加 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` します。</span><span class="sxs-lookup"><span data-stu-id="b1647-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="b1647-128">`[Add-in ID]`アドイン マニフェストの **ID** で置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b1647-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="b1647-129">レジストリ キーを `1` に設定します。</span><span class="sxs-lookup"><span data-stu-id="b1647-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="b1647-130">デスクトップOutlook起動します (または、既に開いている場合はOutlookを再起動します)。</span><span class="sxs-lookup"><span data-stu-id="b1647-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="b1647-131">新しいメッセージまたは予定を作成します。</span><span class="sxs-lookup"><span data-stu-id="b1647-131">Compose a new message or appointment.</span></span> <span data-ttu-id="b1647-132">次のダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b1647-132">You should see the following dialog.</span></span> <span data-ttu-id="b1647-133">ダイアログをまだ操作 *しないでください* 。</span><span class="sxs-lookup"><span data-stu-id="b1647-133">Do *not* interact with the dialog yet.</span></span>

    ![デバッグ イベント ベースのハンドラー ダイアログのスクリーンショット](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="b1647-135">Visual Studio Codeの構成</span><span class="sxs-lookup"><span data-stu-id="b1647-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="b1647-136">ヨーオフィス</span><span class="sxs-lookup"><span data-stu-id="b1647-136">yo office</span></span>

1. <span data-ttu-id="b1647-137">コマンド ライン ウィンドウに戻り、Visual Studio Code開きます。</span><span class="sxs-lookup"><span data-stu-id="b1647-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="b1647-138">Visual Studio Codeで **、./.vscode/launch.jsのファイルを** 開き、次の抜粋を構成のリストに追加します。</span><span class="sxs-lookup"><span data-stu-id="b1647-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="b1647-139">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="b1647-139">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a><span data-ttu-id="b1647-140">その他</span><span class="sxs-lookup"><span data-stu-id="b1647-140">Other</span></span>

1. <span data-ttu-id="b1647-141">**デバッグ** という名前の新しいフォルダーを作成します (**デスクトップ** フォルダーの場合があります)。</span><span class="sxs-lookup"><span data-stu-id="b1647-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="b1647-142">Visual Studio Code を開きます。</span><span class="sxs-lookup"><span data-stu-id="b1647-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="b1647-143">[**ファイル**  >  **を開くフォルダ]** に移動し、作成したフォルダに移動して、[**フォルダの選択**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="b1647-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="b1647-144">アクティビティ バーで、[ **デバッグ** ] 項目 (Ctrl + Shift + D) を選択します。</span><span class="sxs-lookup"><span data-stu-id="b1647-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![アクティビティ バーのデバッグ アイコンのスクリーンショット](../images/vs-code-debug.png)

1. <span data-ttu-id="b1647-146">[ **ファイルにlaunch.jsを作成]リンクを選択します** 。</span><span class="sxs-lookup"><span data-stu-id="b1647-146">Select the **create a launch.json file** link.</span></span>

    ![Visual Studio Codeでファイルにlaunch.jsを作成するためのリンクのスクリーンショット](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="b1647-148">[ **環境の選択** ] ドロップダウンで、[ **エッジ: 起動** ] を選択してファイルにlaunch.jsを作成します。</span><span class="sxs-lookup"><span data-stu-id="b1647-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="b1647-149">構成の一覧に次の抜粋を追加します。</span><span class="sxs-lookup"><span data-stu-id="b1647-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="b1647-150">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="b1647-150">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a><span data-ttu-id="b1647-151">Visual Studio Codeを添付する</span><span class="sxs-lookup"><span data-stu-id="b1647-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="b1647-152">アドインの **bundle.js** を検索するには、Windows エクスプローラーで次のフォルダーを開き、アドインの **ID** (マニフェスト内にあります) を検索します。</span><span class="sxs-lookup"><span data-stu-id="b1647-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="b1647-153">この ID のプレフィックスが付いたフォルダーを開き、その完全なパスをコピーします。</span><span class="sxs-lookup"><span data-stu-id="b1647-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="b1647-154">Visual Studio Codeで、そのフォルダから **bundle.js** 開きます。</span><span class="sxs-lookup"><span data-stu-id="b1647-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="b1647-155">ファイル パスのパターンは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="b1647-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="b1647-156">デバッガーを停止する位置bundle.jsにブレークポイントを配置します。</span><span class="sxs-lookup"><span data-stu-id="b1647-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="b1647-157">**[DEBUG]** ドロップダウンで、[**直接デバッグ**] という名前を選択し、[**実行**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="b1647-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![[Visual Studio Codeデバッグ] ドロップダウンの構成オプションから直接デバッグを選択するスクリーンショット](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="b1647-159">Debug</span><span class="sxs-lookup"><span data-stu-id="b1647-159">Debug</span></span>

1. <span data-ttu-id="b1647-160">デバッガーがアタッチされていることを確認したら、Outlookに戻り、イベント ベースの **デバッグ ハンドラー** ダイアログで **[OK] を** クリックします。</span><span class="sxs-lookup"><span data-stu-id="b1647-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="b1647-161">Visual Studio Codeでブレークポイントにヒットして、イベントベースのアクティブ化コードをデバッグできるようになりました。</span><span class="sxs-lookup"><span data-stu-id="b1647-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="b1647-162">デバッグを停止する</span><span class="sxs-lookup"><span data-stu-id="b1647-162">Stop debugging</span></span>

<span data-ttu-id="b1647-163">現在のOutlookデスクトップ セッションの残りの部分のデバッグを停止するには、[**イベント ベースのデバッグ ハンドラー** ] ダイアログ ボックスで [**キャンセル**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="b1647-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="b1647-164">デバッグを再度有効にするには、デスクトップOutlook再起動します。</span><span class="sxs-lookup"><span data-stu-id="b1647-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="b1647-165">**イベントに基づくデバッグ ハンドラ** ダイアログがポップアップして、後続のOutlook セッションのデバッグを停止しないようにするには、関連付けられたレジストリ キーを削除するか、その値を : に設定 `0` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` します。</span><span class="sxs-lookup"><span data-stu-id="b1647-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="b1647-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="b1647-166">See also</span></span>

- [<span data-ttu-id="b1647-167">イベント ベースのアクティブ化用にOutlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="b1647-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="b1647-168">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="b1647-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
