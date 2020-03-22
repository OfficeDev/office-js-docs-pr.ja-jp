---
title: アドイン コマンドの基本概念
description: Office Web アドインの一部として、カスタム リボン ボタンやメニュー項目を Office に追加する方法について説明します。
ms.date: 02/11/2020
localization_priority: Priority
ms.openlocfilehash: 6395b087ea191b37e9398096038dacfd66ed263c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890557"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="300e9-103">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="300e9-103">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="300e9-p101">アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。ユーザーがアドイン コマンドを選択すると、JavaScript コードを実行したり、アドインのページを作業ウィンドウに表示するなどのアクションが開始されます。アドイン コマンドは、ユーザーがアドインを検索して使用ために役立ちます。これにより、アドインの導入と再利用を促進し、顧客維持率を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="300e9-108">機能の概要については、ビデオ「[Office リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="300e9-108">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="300e9-p102">SharePoint カタログは、アドイン コマンドをサポートしません。[集中展開](../publish/centralized-deployment.md)または [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) でアドイン コマンドを展開するか、または[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)を使ってテストのためのアドイン コマンドを展開できます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="300e9-111">*図 1. Excel デスクトップで実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="300e9-111">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel のアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

<span data-ttu-id="300e9-113">*図 2. Excel on the web で実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="300e9-113">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Excel on the web のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="300e9-115">コマンドの機能</span><span class="sxs-lookup"><span data-stu-id="300e9-115">Command capabilities</span></span>

<span data-ttu-id="300e9-116">現在は、次のコマンド機能がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="300e9-116">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="300e9-117">現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="300e9-117">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="300e9-118">拡張点</span><span class="sxs-lookup"><span data-stu-id="300e9-118">Extension points</span></span>

- <span data-ttu-id="300e9-119">リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。</span><span class="sxs-lookup"><span data-stu-id="300e9-119">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="300e9-120">コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。</span><span class="sxs-lookup"><span data-stu-id="300e9-120">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="300e9-121">コントロールの種類</span><span class="sxs-lookup"><span data-stu-id="300e9-121">Control types</span></span>

- <span data-ttu-id="300e9-122">単純なボタン: 特定のアクションをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="300e9-122">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="300e9-123">メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。</span><span class="sxs-lookup"><span data-stu-id="300e9-123">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="300e9-124">アクション</span><span class="sxs-lookup"><span data-stu-id="300e9-124">Actions</span></span>

- <span data-ttu-id="300e9-125">ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="300e9-125">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="300e9-p103">ExecuteFunction: 非表示の HTML ページをロードして、JavaScript 関数を実行します。関数内で UI を表示するには (エラー、進行状況、追加入力など)、[displayDialog](/javascript/api/office/office.ui) API を使用できます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="300e9-128">既定で有効または無効になっている状態 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="300e9-128">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="300e9-129">アドイン起動時にコマンドを有効にするか無効にするかを指定したり、プログラムによって設定を変更したりできます。</span><span class="sxs-lookup"><span data-stu-id="300e9-129">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span> 

> [!NOTE]
> <span data-ttu-id="300e9-130">この機能はプレビュー段階にあり、すべてのホストまたはシナリオでサポートされるわけではありません。</span><span class="sxs-lookup"><span data-stu-id="300e9-130">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="300e9-131">詳細については、「[アドイン コマンドを有効または無効にする](disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="300e9-131">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="300e9-132">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="300e9-132">Supported platforms</span></span>

<span data-ttu-id="300e9-133">現在、アドイン コマンドは次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="300e9-133">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="300e9-134">Windows 版 Outlook 2016 (ビルド 16.0.4678.1000 以降)</span><span class="sxs-lookup"><span data-stu-id="300e9-134">Outlook 2016 on Windows (build 16.0.4678.1000+)</span></span>
- <span data-ttu-id="300e9-135">Windows 上の Office (ビルド 16.0.6769 以降、Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="300e9-135">Office on Windows (build 16.0.6769+, connected to Office 365 subscription)</span></span>
- <span data-ttu-id="300e9-136">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="300e9-136">Office 2019 on Windows</span></span>
- <span data-ttu-id="300e9-137">Mac 上の Office (ビルド 15.33 以降、Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="300e9-137">Office on Mac (build 15.33+, connected to Office 365 subscription)</span></span>
- <span data-ttu-id="300e9-138">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="300e9-138">Office 2019 on Mac</span></span>
- <span data-ttu-id="300e9-139">Office on the web</span><span class="sxs-lookup"><span data-stu-id="300e9-139">Office on the web</span></span>

## <a name="debugging"></a><span data-ttu-id="300e9-140">デバッグ</span><span class="sxs-lookup"><span data-stu-id="300e9-140">Debugging</span></span>

<span data-ttu-id="300e9-141">アドイン コマンドをデバッグするには、Office on the web で実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="300e9-141">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="300e9-142">詳細については、「[Office on the web でアドインをデバッグする](../testing/debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="300e9-142">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="300e9-143">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="300e9-143">Best practices</span></span>

<span data-ttu-id="300e9-144">アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。</span><span class="sxs-lookup"><span data-stu-id="300e9-144">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="300e9-p106">ユーザーに対して、特定のアクションとともにアクションの結果を明確かつ具体的に表すコマンドを使用します。複数のアクションを 1 つのボタンにまとめないでください。</span><span class="sxs-lookup"><span data-stu-id="300e9-p106">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="300e9-p107">アドイン内の一般的なタスクをより効率的に実行できるように、アクションは細分化して提供します。1 つのアクションを完了するまでのステップ数は最小限に抑えます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p107">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="300e9-149">Office リボンにコマンドを配置するために。</span><span class="sxs-lookup"><span data-stu-id="300e9-149">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="300e9-p108">提供する機能が適応する場合は既存のタブ (挿入、レビューなど) にコマンドを配置します。たとえば、アドインを使用することでユーザーがメディアを挿入できる場合は、[挿入] タブにグループを追加します。Office のすべてのバージョンで、すべてのタブが使用可能なわけではない点に注意してください。詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="300e9-p108">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="300e9-p109">別のタブに機能が適応せず、トップ レベル コマンドが 6 個未満の場合は、[ホーム] タブにコマンドを配置します。Office on the web やデスクトップなど、Office の複数のバージョン間でアドインを操作する必要があり、タブがどのバージョンでも利用できるわけではない場合 (たとえば、[デザイン] タブは Office on the web にはありません) は、[ホーム] タブにコマンドを追加できます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p109">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="300e9-155">6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。</span><span class="sxs-lookup"><span data-stu-id="300e9-155">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="300e9-p110">グループに、アドインの名前と一致する名前を指定します。グループが複数ある場合は、そのグループのコマンドが提供する機能に基づいた名前を各グループに付けます。</span><span class="sxs-lookup"><span data-stu-id="300e9-p110">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="300e9-158">アドインの使用スペースを増やす余分なボタンを追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="300e9-158">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="300e9-159">占有領域が大きすぎるアドインは [AppSource 検証](/legal/marketplace/certification-policies)を通過しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="300e9-159">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="300e9-160">すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="300e9-160">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="300e9-161">コマンドをサポートしていないホストでも動作するアドインのバージョンを提供します。</span><span class="sxs-lookup"><span data-stu-id="300e9-161">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="300e9-162">1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) ホストとコマンド非対応 (作業ウィンドウとして) ホストの両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="300e9-162">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="300e9-163">*図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*</span><span class="sxs-lookup"><span data-stu-id="300e9-163">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを示すスクリーンショット](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="300e9-165">次の手順</span><span class="sxs-lookup"><span data-stu-id="300e9-165">Next steps</span></span>

<span data-ttu-id="300e9-166">アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。</span><span class="sxs-lookup"><span data-stu-id="300e9-166">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="300e9-167">マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](../reference/manifest/versionoverrides.md)」のリファレンス資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="300e9-167">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
