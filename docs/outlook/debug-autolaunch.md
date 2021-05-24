---
title: イベント ベースのアドインOutlookデバッグする (プレビュー)
description: イベント ベースのアクティブ化を実装Outlookアドインをデバッグする方法について説明します。
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
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="74507-103">イベント ベースのアドインOutlookデバッグする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="74507-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="74507-104">この記事では、アドインでイベント ベースの [ライセンス認証を](autolaunch.md) 実装する場合のデバッグ ガイダンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="74507-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="74507-105">イベント ベースのアクティブ化機能は現在プレビュー中です。</span><span class="sxs-lookup"><span data-stu-id="74507-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="74507-106">このデバッグ機能は、サブスクリプションを使用したOutlookのWindowsプレビューでのみMicrosoft 365されます。</span><span class="sxs-lookup"><span data-stu-id="74507-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="74507-107">詳細については、この記事の「イベント ベースのアクティブ [化機能の](#preview-debugging-for-the-event-based-activation-feature) プレビュー デバッグ」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="74507-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="74507-108">この記事では、デバッグを有効にする重要な段階について説明します。</span><span class="sxs-lookup"><span data-stu-id="74507-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="74507-109">デバッグ用にアドインをマークする</span><span class="sxs-lookup"><span data-stu-id="74507-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="74507-110">構成Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="74507-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="74507-111">添付Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="74507-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="74507-112">Debug</span><span class="sxs-lookup"><span data-stu-id="74507-112">Debug</span></span>](#debug)

<span data-ttu-id="74507-113">アドイン プロジェクトを作成するには、いくつかのオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="74507-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="74507-114">使用しているオプションによっては、手順が異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="74507-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="74507-115">このような場合は、Office アドインに Yeoman ジェネレーターを使用してアドイン プロジェクトを作成した場合 (たとえば、イベント ベースのライセンス認証のチュートリアルを実行します)、office のヨーヨー 手順に従い、それ以外の場合は、その他の手順に従います。 [](autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="74507-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="74507-116">Visual Studio Codeバージョン 1.56.1 以上である必要があります。</span><span class="sxs-lookup"><span data-stu-id="74507-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="74507-117">イベント ベースのアクティブ化機能のデバッグをプレビューする</span><span class="sxs-lookup"><span data-stu-id="74507-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="74507-118">イベント ベースのアクティブ化機能のデバッグ機能を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="74507-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="74507-119">このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。</span><span class="sxs-lookup"><span data-stu-id="74507-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="74507-120">この機能を Outlook Windowsでプレビューするには、必要な最小ビルドは 16.0.13729.20000 です。</span><span class="sxs-lookup"><span data-stu-id="74507-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="74507-121">ベータビルドへのアクセスOffice、Insider プログラムOffice[参加してください](https://insider.office.com)。</span><span class="sxs-lookup"><span data-stu-id="74507-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="74507-122">デバッグ用にアドインをマークする</span><span class="sxs-lookup"><span data-stu-id="74507-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="74507-123">レジストリ キーを設定します `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。</span><span class="sxs-lookup"><span data-stu-id="74507-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="74507-124">`[Add-in ID]` は **、アドイン** マニフェストの ID です。</span><span class="sxs-lookup"><span data-stu-id="74507-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="74507-125">**yo office**: コマンド ライン ウィンドウで、アドイン フォルダーのルートに移動し、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="74507-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="74507-126">コードを構築し、ローカル サーバーを起動する以外に、このコマンドは、このアドインのレジストリ キーを `UseDirectDebugger` に設定する必要があります `1` 。</span><span class="sxs-lookup"><span data-stu-id="74507-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="74507-127">**その他**: の下 `UseDirectDebugger` にレジストリ キーを追加 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` します。</span><span class="sxs-lookup"><span data-stu-id="74507-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="74507-128">アドイン `[Add-in ID]` マニフェスト **の Id** に置き換える。</span><span class="sxs-lookup"><span data-stu-id="74507-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="74507-129">レジストリ キーをに設定します `1` 。</span><span class="sxs-lookup"><span data-stu-id="74507-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="74507-130">デスクトップOutlook起動します (またはOutlook開いている場合は再起動します)。</span><span class="sxs-lookup"><span data-stu-id="74507-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="74507-131">新しいメッセージまたは予定を作成します。</span><span class="sxs-lookup"><span data-stu-id="74507-131">Compose a new message or appointment.</span></span> <span data-ttu-id="74507-132">次のダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="74507-132">You should see the following dialog.</span></span> <span data-ttu-id="74507-133">ダイアログ *を* まだ操作しないでください。</span><span class="sxs-lookup"><span data-stu-id="74507-133">Do *not* interact with the dialog yet.</span></span>

    ![イベント ベースのハンドラー のデバッグ ダイアログのスクリーンショット](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="74507-135">構成Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="74507-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="74507-136">yo office</span><span class="sxs-lookup"><span data-stu-id="74507-136">yo office</span></span>

1. <span data-ttu-id="74507-137">コマンド ライン ウィンドウに戻り、コマンド ライン ウィンドウVisual Studio Code。</span><span class="sxs-lookup"><span data-stu-id="74507-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="74507-138">このVisual Studio Code **./.vscode/launch.js** を開き、構成の一覧に次の抜粋を追加します。</span><span class="sxs-lookup"><span data-stu-id="74507-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="74507-139">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="74507-139">Save your changes.</span></span>

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

### <a name="other"></a><span data-ttu-id="74507-140">その他</span><span class="sxs-lookup"><span data-stu-id="74507-140">Other</span></span>

1. <span data-ttu-id="74507-141">デバッグという名前の新しい **フォルダーを作成** します (おそらくデスクトップ フォルダー **に)。**</span><span class="sxs-lookup"><span data-stu-id="74507-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="74507-142">Visual Studio Code を開きます。</span><span class="sxs-lookup"><span data-stu-id="74507-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="74507-143">[ファイルを **開**  >  **くフォルダー]** に移動し、作成したフォルダーに移動し、[フォルダーの選択]**を選択します**。</span><span class="sxs-lookup"><span data-stu-id="74507-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="74507-144">[アクティビティ バー] で、[デバッグ] **アイテム** (Ctrl + Shift + D) を選択します。</span><span class="sxs-lookup"><span data-stu-id="74507-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![アクティビティ バーの [デバッグ] アイコンのスクリーンショット](../images/vs-code-debug.png)

1. <span data-ttu-id="74507-146">[ファイルに **対してlaunch.jsを作成する] リンクを選択** します。</span><span class="sxs-lookup"><span data-stu-id="74507-146">Select the **create a launch.json file** link.</span></span>

    ![ページ内のファイルにlaunch.jsを作成するリンクVisual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="74507-148">[環境 **の選択] ドロップダウン** で、[ **エッジ:** 起動] を選択して、launch.jsを作成します。</span><span class="sxs-lookup"><span data-stu-id="74507-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="74507-149">構成の一覧に次の抜粋を追加します。</span><span class="sxs-lookup"><span data-stu-id="74507-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="74507-150">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="74507-150">Save your changes.</span></span>

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

## <a name="attach-visual-studio-code"></a><span data-ttu-id="74507-151">添付Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="74507-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="74507-152">アドインのbundle.jsを見 **つける** には、Windows エクスプローラーで次のフォルダーを開き、アドインの **ID** (マニフェストにある) を検索します。</span><span class="sxs-lookup"><span data-stu-id="74507-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="74507-153">この ID のプレフィックスが付いたフォルダーを開き、完全なパスをコピーします。</span><span class="sxs-lookup"><span data-stu-id="74507-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="74507-154">このVisual Studio Code、その **フォルダーbundle.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="74507-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="74507-155">ファイル パスのパターンは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="74507-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="74507-156">デバッガーを停止bundle.js場所にブレークポイントを配置します。</span><span class="sxs-lookup"><span data-stu-id="74507-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="74507-157">[ **デバッグ] ドロップダウンで** 、[直接デバッグ] という **名前を選択し**、[実行] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="74507-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![[デバッグ] ドロップダウンの構成オプションから [直接デバッグ] を選択Visual Studio Codeスクリーンショット](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="74507-159">Debug</span><span class="sxs-lookup"><span data-stu-id="74507-159">Debug</span></span>

1. <span data-ttu-id="74507-160">デバッガーが接続されているのを確認した後、Outlook に戻り、[イベント ベースのハンドラーのデバッグ] ダイアログで **[OK] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="74507-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="74507-161">これで、イベント ベースのアクティブ化コードVisual Studio Codeデバッグを有効にすることで、ブレークポイントをヒットできます。</span><span class="sxs-lookup"><span data-stu-id="74507-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="74507-162">デバッグを停止する</span><span class="sxs-lookup"><span data-stu-id="74507-162">Stop debugging</span></span>

<span data-ttu-id="74507-163">現在のデスクトップ セッションの残りのOutlook停止するには、[イベント ベースのハンドラーのデバッグ] ダイアログで、[キャンセル] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="74507-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="74507-164">デバッグを再び有効にするには、デスクトップOutlookします。</span><span class="sxs-lookup"><span data-stu-id="74507-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="74507-165">イベント ベースの **ハンドラー の** デバッグ ダイアログがポップアップし、後続の Outlook セッションのデバッグを停止するには、関連付けられたレジストリ キーを削除するか、その値を : に設定します `0` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。</span><span class="sxs-lookup"><span data-stu-id="74507-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="74507-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="74507-166">See also</span></span>

- [<span data-ttu-id="74507-167">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="74507-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="74507-168">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="74507-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
