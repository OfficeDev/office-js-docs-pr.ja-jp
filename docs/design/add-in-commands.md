---
title: アドイン コマンドの基本概念
description: Office アドインの一部として、カスタム リボン ボタンやメニュー項目を Office に追加する方法について説明します。
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093876"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="5d0dc-103">Excel、PowerPoint、Word のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="5d0dc-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="5d0dc-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span></span> <span data-ttu-id="5d0dc-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span></span> <span data-ttu-id="5d0dc-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> <span data-ttu-id="5d0dc-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="5d0dc-108">機能の概要については、ビデオ「[Office アプリ リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="5d0dc-109">SharePoint catalogs do not support add-in commands.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-109">SharePoint catalogs do not support add-in commands.</span></span> <span data-ttu-id="5d0dc-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5d0dc-111">アドイン コマンドは、Outlook でもサポートされています。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="5d0dc-112">詳細については、「[Outlook のアドイン コマンド](../outlook/add-in-commands-for-outlook.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="5d0dc-113">*図 1. Excel デスクトップで実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="5d0dc-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel のアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

<span data-ttu-id="5d0dc-115">*図 2. Excel on the web で実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="5d0dc-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Excel on the web のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="5d0dc-117">コマンドの機能</span><span class="sxs-lookup"><span data-stu-id="5d0dc-117">Command capabilities</span></span>

<span data-ttu-id="5d0dc-118">現在は、次のコマンド機能がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="5d0dc-119">現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="5d0dc-120">拡張点</span><span class="sxs-lookup"><span data-stu-id="5d0dc-120">Extension points</span></span>

- <span data-ttu-id="5d0dc-121">リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="5d0dc-122">コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="5d0dc-123">コントロールの種類</span><span class="sxs-lookup"><span data-stu-id="5d0dc-123">Control types</span></span>

- <span data-ttu-id="5d0dc-124">単純なボタン: 特定のアクションをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="5d0dc-125">メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="5d0dc-126">アクション</span><span class="sxs-lookup"><span data-stu-id="5d0dc-126">Actions</span></span>

- <span data-ttu-id="5d0dc-127">ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="5d0dc-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span></span> <span data-ttu-id="5d0dc-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="5d0dc-130">既定で有効または無効になっている状態 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="5d0dc-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="5d0dc-131">アドイン起動時にコマンドを有効にするか無効にするかを指定したり、プログラムによって設定を変更したりできます。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="5d0dc-132">この機能はプレビュー段階にあり、すべてのホストまたはシナリオでサポートされるわけではありません。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="5d0dc-133">詳細については、「[アドイン コマンドを有効または無効にする](disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="5d0dc-134">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5d0dc-134">Supported platforms</span></span>

<span data-ttu-id="5d0dc-135">現在、アドイン コマンドは次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="5d0dc-136">Windows 上の Office (ビルド 16.0.6769 以降、Microsoft 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="5d0dc-136">Office on Windows (build 16.0.6769+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="5d0dc-137">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5d0dc-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="5d0dc-138">Mac 上の Office (ビルド 15.33 以降、Microsoft 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="5d0dc-138">Office on Mac (build 15.33+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="5d0dc-139">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="5d0dc-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="5d0dc-140">Office on the web</span><span class="sxs-lookup"><span data-stu-id="5d0dc-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="5d0dc-141">Outlook でのサポートについては、「[Outlook のアドイン コマンド](../outlook/add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="5d0dc-142">デバッグ</span><span class="sxs-lookup"><span data-stu-id="5d0dc-142">Debugging</span></span>

<span data-ttu-id="5d0dc-143">アドイン コマンドをデバッグするには、Office on the web で実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="5d0dc-144">詳細については、「[Office on the web でアドインをデバッグする](../testing/debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="5d0dc-145">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="5d0dc-145">Best practices</span></span>

<span data-ttu-id="5d0dc-146">アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="5d0dc-147">Use commands to represent a specific action with a clear and specific outcome for users.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-147">Use commands to represent a specific action with a clear and specific outcome for users.</span></span> <span data-ttu-id="5d0dc-148">Do not combine multiple actions in a single button.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-148">Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="5d0dc-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span></span> <span data-ttu-id="5d0dc-150">Minimize the number of steps an action takes to complete.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-150">Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="5d0dc-151">Office アプリ リボンにコマンドを配置するために。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="5d0dc-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span></span> <span data-ttu-id="5d0dc-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span></span> <span data-ttu-id="5d0dc-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="5d0dc-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="5d0dc-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span></span> <span data-ttu-id="5d0dc-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span><span class="sxs-lookup"><span data-stu-id="5d0dc-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="5d0dc-157">6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="5d0dc-158">Name your group to match the name of your add-in.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-158">Name your group to match the name of your add-in.</span></span> <span data-ttu-id="5d0dc-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span><span class="sxs-lookup"><span data-stu-id="5d0dc-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="5d0dc-160">アドインの使用スペースを増やす余分なボタンを追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="5d0dc-161">占有領域が大きすぎるアドインは [AppSource 検証](/legal/marketplace/certification-policies)を通過しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="5d0dc-162">すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="5d0dc-163">コマンドをサポートしていないホストでも動作するアドインのバージョンを提供します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="5d0dc-164">1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) ホストとコマンド非対応 (作業ウィンドウとして) ホストの両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="5d0dc-165">*図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*</span><span class="sxs-lookup"><span data-stu-id="5d0dc-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを示すスクリーンショット](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="5d0dc-167">次の手順</span><span class="sxs-lookup"><span data-stu-id="5d0dc-167">Next steps</span></span>

<span data-ttu-id="5d0dc-168">アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="5d0dc-169">マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](../reference/manifest/versionoverrides.md)」のリファレンス資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5d0dc-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
