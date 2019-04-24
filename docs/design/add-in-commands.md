---
title: Excel、Word、PowerPoint のアドイン コマンド
description: ''
ms.date: 04/05/2019
localization_priority: Priority
ms.openlocfilehash: 0a2a28e09061248851e8ee1037d0219c7fcda1a5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448737"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="d55d3-102">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="d55d3-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="d55d3-p101">アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。ユーザーがアドイン コマンドを選択すると、JavaScript コードを実行したり、アドインのページを作業ウィンドウに表示するなどのアクションが開始されます。アドイン コマンドは、ユーザーがアドインを検索して使用ために役立ちます。これにより、アドインの導入と再利用を促進し、顧客維持率を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="d55d3-107">機能の概要については、ビデオ「[Office リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="d55d3-p102">SharePoint カタログは、アドイン コマンドをサポートしません。[集中展開](../publish/centralized-deployment.md)または [AppSource](/office/dev/store/submit-to-the-office-store) でアドイン コマンドを展開するか、または[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)を使ってテストのためのアドイン コマンドを展開できます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="d55d3-110">*図 1. Excel デスクトップで実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="d55d3-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel のアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

<span data-ttu-id="d55d3-112">*図 2. Excel Online で実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="d55d3-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Excel Online のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="d55d3-114">コマンドの機能</span><span class="sxs-lookup"><span data-stu-id="d55d3-114">Command capabilities</span></span>

<span data-ttu-id="d55d3-115">現在は、次のコマンド機能がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d55d3-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="d55d3-116">現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="d55d3-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="d55d3-117">**拡張点**</span><span class="sxs-lookup"><span data-stu-id="d55d3-117">**Extension points**</span></span>

- <span data-ttu-id="d55d3-118">リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="d55d3-119">コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-119">Context menus - Extend selected context menus.</span></span>

<span data-ttu-id="d55d3-120">**コントロールの種類**</span><span class="sxs-lookup"><span data-stu-id="d55d3-120">**Control types**</span></span>

- <span data-ttu-id="d55d3-121">単純なボタン: 特定のアクションをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="d55d3-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="d55d3-122">メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。</span><span class="sxs-lookup"><span data-stu-id="d55d3-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="d55d3-123">**アクション**</span><span class="sxs-lookup"><span data-stu-id="d55d3-123">**Actions**</span></span>

- <span data-ttu-id="d55d3-124">ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="d55d3-p103">ExecuteFunction: 非表示の HTML ページをロードして、JavaScript 関数を実行します。関数内で UI を表示するには (エラー、進行状況、追加入力など)、[displayDialog](/javascript/api/office/office.ui) API を使用できます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="d55d3-127">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="d55d3-127">Supported platforms</span></span>

<span data-ttu-id="d55d3-128">現在、アドイン コマンドは次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d55d3-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="d55d3-129">Outlook 2016 for Windows 以降 (ビルド 16.0.6769 以降)</span><span class="sxs-lookup"><span data-stu-id="d55d3-129">Office 2016 or later for Windows (build 16.0.6769+)</span></span>
- <span data-ttu-id="d55d3-130">Office for Mac (ビルド 15.33 以降)</span><span class="sxs-lookup"><span data-stu-id="d55d3-130">Office for Mac (build 15.33+)</span></span>
- <span data-ttu-id="d55d3-131">Office Online</span><span class="sxs-lookup"><span data-stu-id="d55d3-131">Office Online</span></span>

<span data-ttu-id="d55d3-132">その他のプラットフォームが近日中に公開されます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-132">More platforms are coming soon.</span></span>

## <a name="debugging"></a><span data-ttu-id="d55d3-133">デバッグ</span><span class="sxs-lookup"><span data-stu-id="d55d3-133">Debugging</span></span>

<span data-ttu-id="d55d3-134">アドイン コマンドをデバッグするには、Office Online で実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d55d3-134">To debug an Add-in Command, you must run it in Office Online.</span></span> <span data-ttu-id="d55d3-135">詳細については、「[Office Online でアドインをデバッグする](../testing/debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-135">For details, see [Debug add-ins in Office Online](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="d55d3-136">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="d55d3-136">Best practices</span></span>

<span data-ttu-id="d55d3-137">アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-137">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="d55d3-p105">ユーザーに対して、特定のアクションとともにアクションの結果を明確かつ具体的に表すコマンドを使用します。複数のアクションを 1 つのボタンにまとめないでください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p105">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="d55d3-p106">アドイン内の一般的なタスクをより効率的に実行できるように、アクションは細分化して提供します。1 つのアクションを完了するまでのステップ数は最小限に抑えます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p106">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="d55d3-142">Office リボンにコマンドを配置するために。</span><span class="sxs-lookup"><span data-stu-id="d55d3-142">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="d55d3-p107">提供する機能が適応する場合は既存のタブ (挿入、レビューなど) にコマンドを配置します。たとえば、アドインを使用することでユーザーがメディアを挿入できる場合は、[挿入] タブにグループを追加します。Office のすべてのバージョンで、すべてのタブが使用可能なわけではない点に注意してください。詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p107">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> 
    - <span data-ttu-id="d55d3-p108">別のタブに機能が適応せず、トップ レベル コマンドが 6 個未満の場合は、[ホーム] タブにコマンドを配置します。Office デスクトップと Office Online など、Office の複数のバージョン間でアドインを操作する必要があり、タブがどのバージョンでも利用できるわけではない場合 (たとえば、[デザイン] タブは Office Online にはありません) は、[ホーム] タブにコマンドを追加できます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p108">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="d55d3-148">6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-148">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="d55d3-p109">グループに、アドインの名前と一致する名前を指定します。グループが複数ある場合は、そのグループのコマンドが提供する機能に基づいた名前を各グループに付けます。</span><span class="sxs-lookup"><span data-stu-id="d55d3-p109">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="d55d3-151">アドインの使用スペースを増やす余分なボタンを追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-151">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="d55d3-152">占有領域が大きすぎるアドインは [AppSource 検証](/office/dev/store/validation-policies)を通過しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="d55d3-152">Add-ins that take up too much space might not pass [AppSource validation](/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="d55d3-153">すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-153">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="d55d3-154">コマンドをサポートしていないホストでも動作するアドインのバージョンを提供します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-154">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="d55d3-155">1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) ホストとコマンド非対応 (作業ウィンドウとして) ホストの両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="d55d3-155">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="d55d3-156">*図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*</span><span class="sxs-lookup"><span data-stu-id="d55d3-156">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを示すスクリーンショット](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="d55d3-158">次の手順</span><span class="sxs-lookup"><span data-stu-id="d55d3-158">Next steps</span></span>

<span data-ttu-id="d55d3-159">アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。</span><span class="sxs-lookup"><span data-stu-id="d55d3-159">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="d55d3-160">マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](/office/dev/add-ins/reference/manifest/versionoverrides)」のリファレンス資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d55d3-160">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) reference content.</span></span>
