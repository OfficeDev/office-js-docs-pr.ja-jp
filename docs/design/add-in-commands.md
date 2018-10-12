---
title: Excel、Word、PowerPoint のアドイン コマンド
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 3db846e9d28e063d959fd617bf8c50ab5cb5ec86
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506267"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="cccae-102">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="cccae-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="cccae-p101">アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。ユーザーがアドイン コマンドを選択すると、JavaScript コードを実行したり、アドインのページを作業ウィンドウに表示するなどのアクションが開始されます。アドイン コマンドは、ユーザーがアドインを検索して使用ために役立ちます。これにより、アドインの導入と再利用を促進し、顧客維持率を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="cccae-107">機能の概要については、ビデオ「[Office リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cccae-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="cccae-p102">SharePoint カタログは、アドイン コマンドをサポートしません。[集中展開](../publish/centralized-deployment.md)または [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store) でアドイン コマンドを展開するか、または[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)を使ってテストのためのアドイン コマンドを展開できます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="cccae-110">*図 1. Excel デスクトップで実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="cccae-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel のアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

<span data-ttu-id="cccae-112">*図 2. Excel Online で実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="cccae-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Excel Online のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="cccae-114">コマンドの機能</span><span class="sxs-lookup"><span data-stu-id="cccae-114">Command capabilities</span></span>
<span data-ttu-id="cccae-115">現在は、次のコマンド機能がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cccae-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="cccae-116">現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="cccae-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="cccae-117">**拡張点**</span><span class="sxs-lookup"><span data-stu-id="cccae-117">**Extension points**</span></span>

- <span data-ttu-id="cccae-118">リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。</span><span class="sxs-lookup"><span data-stu-id="cccae-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="cccae-119">コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。</span><span class="sxs-lookup"><span data-stu-id="cccae-119">Context menus - Extend selected context menus.</span></span> 

<span data-ttu-id="cccae-120">**コントロールの種類**</span><span class="sxs-lookup"><span data-stu-id="cccae-120">**Control types**</span></span>

- <span data-ttu-id="cccae-121">単純なボタン: 特定のアクションをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="cccae-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="cccae-122">メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。</span><span class="sxs-lookup"><span data-stu-id="cccae-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="cccae-123">**アクション**</span><span class="sxs-lookup"><span data-stu-id="cccae-123">**Actions**</span></span>

- <span data-ttu-id="cccae-124">ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="cccae-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="cccae-p103">ExecuteFunction: 非表示の HTML ページをロードして、JavaScript 関数を実行します。関数内で UI を表示するには (エラー、進行状況、追加入力など)、[displayDialog](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) API を使用できます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="cccae-127">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="cccae-127">Supported platforms</span></span>

<span data-ttu-id="cccae-128">現在、アドイン コマンドは次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cccae-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="cccae-129">Windows デスクトップ 2016 版 Office (ビルド 16.0.6769 以降)</span><span class="sxs-lookup"><span data-stu-id="cccae-129">Office for Windows Desktop 2016 (build 16.0.6769+)</span></span>
- <span data-ttu-id="cccae-130">Office for Mac (ビルド 15.33 以降)</span><span class="sxs-lookup"><span data-stu-id="cccae-130">Office for Mac (build 15.33+)</span></span>
- <span data-ttu-id="cccae-131">Office Online</span><span class="sxs-lookup"><span data-stu-id="cccae-131">Office Online</span></span> 

<span data-ttu-id="cccae-132">その他のプラットフォームが近日中に公開されます。</span><span class="sxs-lookup"><span data-stu-id="cccae-132">More platforms are coming soon.</span></span>

## <a name="best-practices"></a><span data-ttu-id="cccae-133">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cccae-133">Best practices</span></span>

<span data-ttu-id="cccae-134">アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。</span><span class="sxs-lookup"><span data-stu-id="cccae-134">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="cccae-p104">ユーザーに対して、特定のアクションとともにアクションの結果を明確かつ具体的に表すコマンドを使用します。複数のアクションを 1 つのボタンにまとめないでください。</span><span class="sxs-lookup"><span data-stu-id="cccae-p104">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="cccae-p105">アドイン内の一般的なタスクをより効率的に実行できるように、アクションは細分化して提供します。1 つのアクションを完了するまでのステップ数は最小限に抑えます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p105">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="cccae-139">Office リボンにコマンドを配置するために。</span><span class="sxs-lookup"><span data-stu-id="cccae-139">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="cccae-p106">提供する機能が適応する場合は既存のタブ (挿入、レビューなど) にコマンドを配置します。たとえば、アドインを使用することでユーザーがメディアを挿入できる場合は、[挿入] タブにグループを追加します。Office のすべてのバージョンで、すべてのタブが使用可能なわけではない点に注意してください。詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cccae-p106">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> 
    - <span data-ttu-id="cccae-p107">別のタブに機能が適応せず、トップ レベル コマンドが 6 個未満の場合は、[ホーム] タブにコマンドを配置します。Office デスクトップと Office Online など、Office の複数のバージョン間でアドインを操作する必要があり、タブがどのバージョンでも利用できるわけではない場合 (たとえば、[デザイン] タブは Office Online にはありません) は、[ホーム] タブにコマンドを追加できます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p107">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="cccae-145">6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。</span><span class="sxs-lookup"><span data-stu-id="cccae-145">Place commands on a custom tab if you have more than six top-level commands.</span></span> 
    - <span data-ttu-id="cccae-p108">グループに、アドインの名前と一致する名前を指定します。グループが複数ある場合は、そのグループのコマンドが提供する機能に基づいた名前を各グループに付けます。</span><span class="sxs-lookup"><span data-stu-id="cccae-p108">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="cccae-148">アドインの使用スペースを増やす余分なボタンを追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="cccae-148">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="cccae-149">占有領域が大きすぎるアドインは [AppSource 検証](https://docs.microsoft.com/office/dev/store/validation-policies)を通過しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="cccae-149">Add-ins that take up too much space might not pass [AppSource validation](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="cccae-150">すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="cccae-150">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="cccae-p109">コマンドをサポートしていないホストでも動作するアドインのバージョンを提供します。1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) ホストとコマンド非対応 (作業ウィンドウとして) ホストの両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="cccae-p109">Provide a version of your add-in that also works on hosts that do not support commands. A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a taskpane) hosts.</span></span>

   <span data-ttu-id="cccae-153">*図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*</span><span class="sxs-lookup"><span data-stu-id="cccae-153">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを示すスクリーンショット](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="cccae-155">次の手順</span><span class="sxs-lookup"><span data-stu-id="cccae-155">Next steps</span></span>

<span data-ttu-id="cccae-156">アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。</span><span class="sxs-lookup"><span data-stu-id="cccae-156">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="cccae-157">マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js)」のリファレンス資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cccae-157">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js) reference content.</span></span>




