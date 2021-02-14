---
title: アドイン コマンドの基本概念
description: Office アドインの一部として、カスタム リボン ボタンやメニュー項目を Office に追加する方法について説明します。
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: c9d69b21be5cca0c37feb14f43649b55df532466
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173949"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="8e3b5-103">Excel、PowerPoint、Word のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="8e3b5-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="8e3b5-p101">アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。ユーザーがアドイン コマンドを選択すると、JavaScript コードを実行したり、アドインのページを作業ウィンドウに表示するなどのアクションが開始されます。アドイン コマンドは、ユーザーがアドインを検索して使用ために役立ちます。これにより、アドインの導入と再利用を促進し、顧客維持率を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="8e3b5-108">機能の概要については、ビデオ「[Office アプリ リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-p102">SharePoint カタログは、アドイン コマンドをサポートしません。[集中展開](../publish/centralized-deployment.md)または [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) でアドイン コマンドを展開するか、または[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)を使ってテストのためのアドイン コマンドを展開できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8e3b5-111">アドイン コマンドは、Outlook でもサポートされています。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="8e3b5-112">詳細については、「[Outlook のアドイン コマンド](../outlook/add-in-commands-for-outlook.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="8e3b5-113">*図 1. Excel デスクトップで実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="8e3b5-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel のリボンで強調表示されているアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

<span data-ttu-id="8e3b5-115">*図 2. Excel on the web で実行するコマンドを含むアドイン*</span><span class="sxs-lookup"><span data-stu-id="8e3b5-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Excel on the web のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="8e3b5-117">コマンドの機能</span><span class="sxs-lookup"><span data-stu-id="8e3b5-117">Command capabilities</span></span>

<span data-ttu-id="8e3b5-118">現在は、次のコマンド機能がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-119">現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="8e3b5-120">拡張点</span><span class="sxs-lookup"><span data-stu-id="8e3b5-120">Extension points</span></span>

- <span data-ttu-id="8e3b5-121">リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="8e3b5-122">コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="8e3b5-123">コントロールの種類</span><span class="sxs-lookup"><span data-stu-id="8e3b5-123">Control types</span></span>

- <span data-ttu-id="8e3b5-124">単純なボタン: 特定のアクションをトリガーします。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="8e3b5-125">メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="8e3b5-126">アクション</span><span class="sxs-lookup"><span data-stu-id="8e3b5-126">Actions</span></span>

- <span data-ttu-id="8e3b5-127">ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="8e3b5-p104">ExecuteFunction: 非表示の HTML ページをロードして、JavaScript 関数を実行します。関数内で UI を表示するには (エラー、進行状況、追加入力など)、[displayDialog](/javascript/api/office/office.ui) API を使用できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status"></a><span data-ttu-id="8e3b5-130">既定で有効または無効になっている状態 </span><span class="sxs-lookup"><span data-stu-id="8e3b5-130">Default Enabled or Disabled Status</span></span>

<span data-ttu-id="8e3b5-131">アドイン起動時にコマンドを有効にするか無効にするかを指定したり、プログラムによって設定を変更したりできます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-132">この機能はすべての Office アプリケーションまたはシナリオでサポートされてはいません。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-132">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="8e3b5-133">詳細については、「[アドイン コマンドを有効または無効にする](disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

### <a name="position-on-the-ribbon-preview"></a><span data-ttu-id="8e3b5-134">リボンの位置 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8e3b5-134">Position on the ribbon (preview)</span></span>

<span data-ttu-id="8e3b5-135">「ホームタブのすぐ右側」など、Office アプリケーションのリボンのどこにカスタム タブを表示するかを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-135">You can specify where a custom tab appears on the Office application's ribbon, such as "just to the right of the Home tab".</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-136">この機能はすべての Office アプリケーションまたはシナリオでサポートされてはいません。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-136">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="8e3b5-137">詳細については、「[リボンにカスタムタブを配置する](custom-tab-placement.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-137">For more information, see [Position a custom tab on the ribbon](custom-tab-placement.md).</span></span>

### <a name="integration-of-built-in-office-buttons-preview"></a><span data-ttu-id="8e3b5-138">組み込みの Office ボタンの統合 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8e3b5-138">Integration of built-in Office buttons (preview)</span></span>

<span data-ttu-id="8e3b5-139">組み込みの Office リボン ボタンをカスタム コマンド グループとカスタム リボン タブに挿入できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-139">You can insert the built-in Office ribbon buttons into your custom command groups and custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-140">この機能はすべての Office アプリケーションまたはシナリオでサポートされてはいません。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-140">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="8e3b5-141">詳細については、「[組み込みの Office ボタンをカスタム タブに統合する](built-in-button-integration.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-141">For more information, see [Integrate built-in Office buttons into custom tabs](built-in-button-integration.md).</span></span>

### <a name="contextual-tabs-preview"></a><span data-ttu-id="8e3b5-142">コンテキスト タブ (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8e3b5-142">Contextual tabs (preview)</span></span>

<span data-ttu-id="8e3b5-143">Excel でグラフが選択されている場合など、特定のコンテキストでのみタブがリボンに表示されるように指定できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-143">You can specify that a tab is only visible on the ribbon in certain contexts, such as when a chart is selected in Excel.</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-144">この機能はすべての Office アプリケーションまたはシナリオでサポートされてはいません。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-144">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="8e3b5-145">詳細については、「[Office アドインでカスタム コンテキスト タブを作成する (プレビュー)](contextual-tabs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-145">For more information, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="8e3b5-146">サポートされるプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="8e3b5-146">Supported platforms</span></span>

<span data-ttu-id="8e3b5-147">現在アドイン コマンドは、以前に[コマンドの機能](#command-capabilities)のサブ セクションで指定された制限を除いて、次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-147">Add-in commands are currently supported on the following platforms, except for limitations specified in the subsections of [Command capabilities](#command-capabilities) earlier.</span></span>

- <span data-ttu-id="8e3b5-148">Windows 上の Office (ビルド 16.0.6769 以降、Microsoft 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="8e3b5-148">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="8e3b5-149">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="8e3b5-149">Office 2019 on Windows</span></span>
- <span data-ttu-id="8e3b5-150">Mac 上の Office (ビルド 15.33 以降、Microsoft 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="8e3b5-150">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="8e3b5-151">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="8e3b5-151">Office 2019 on Mac</span></span>
- <span data-ttu-id="8e3b5-152">Office on the web</span><span class="sxs-lookup"><span data-stu-id="8e3b5-152">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="8e3b5-153">Outlook でのサポートについては、「[Outlook のアドイン コマンド](../outlook/add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-153">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="8e3b5-154">デバッグ</span><span class="sxs-lookup"><span data-stu-id="8e3b5-154">Debugging</span></span>

<span data-ttu-id="8e3b5-155">アドイン コマンドをデバッグするには、Office on the web で実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-155">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="8e3b5-156">詳細については、「[Office on the web でアドインをデバッグする](../testing/debug-add-ins-in-office-online.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-156">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="8e3b5-157">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8e3b5-157">Best practices</span></span>

<span data-ttu-id="8e3b5-158">アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-158">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="8e3b5-p110">ユーザーに対して、特定のアクションとともにアクションの結果を明確かつ具体的に表すコマンドを使用します。複数のアクションを 1 つのボタンにまとめないでください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p110">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="8e3b5-p111">アドイン内の一般的なタスクをより効率的に実行できるように、アクションは細分化して提供します。1 つのアクションを完了するまでのステップ数は最小限に抑えます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p111">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="8e3b5-163">Office アプリ リボンにコマンドを配置するために。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-163">For the placement of your commands in the Office app ribbon:</span></span>
  - <span data-ttu-id="8e3b5-p112">提供する機能が適応する場合は既存のタブ (挿入、レビューなど) にコマンドを配置します。たとえば、アドインを使用することでユーザーがメディアを挿入できる場合は、[挿入] タブにグループを追加します。Office のすべてのバージョンで、すべてのタブが使用可能なわけではない点に注意してください。詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p112">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
  - <span data-ttu-id="8e3b5-p113">別のタブに機能が適応せず、トップ レベル コマンドが 6 個未満の場合は、[ホーム] タブにコマンドを配置します。Office on the web やデスクトップなど、Office の複数のバージョン間でアドインを操作する必要があり、タブがどのバージョンでも利用できるわけではない場合 (たとえば、[デザイン] タブは Office on the web にはありません) は、[ホーム] タブにコマンドを追加できます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p113">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
  - <span data-ttu-id="8e3b5-169">6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-169">Place commands on a custom tab if you have more than six top-level commands.</span></span>
  - <span data-ttu-id="8e3b5-p114">グループに、アドインの名前と一致する名前を指定します。グループが複数ある場合は、そのグループのコマンドが提供する機能に基づいた名前を各グループに付けます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-p114">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
  - <span data-ttu-id="8e3b5-172">アドインの使用スペースを増やす余分なボタンを追加しないでください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-172">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>
  - <span data-ttu-id="8e3b5-173">ユーザーがドキュメントを操作する主な方法がアドインである場合を除き、カスタムタブを [ホーム] タブの左側に配置したり、ドキュメントを開いたときに既定でフォーカスを設定したりしないでください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-173">Do not position a custom tab to the left of the Home tab, or give it focus by default when the document opens, unless your add-in is the primary way users will interact with the document.</span></span> <span data-ttu-id="8e3b5-174">アドインの不便さを過度に目立たせ、ユーザーや管理者を悩ませます。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-174">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span>
  - <span data-ttu-id="8e3b5-175">アドインがユーザーがドキュメントを操作する主な方法であり、カスタム リボン タブがある場合は、ユーザーが頻繁に必要とする Office 機能のボタンをタブに統合することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-175">If your add-in is the primary way users interact with the document and you have a custom ribbon tab, consider integrating into the tab the buttons for the Office functions that users will frequently need.</span></span>
  - <span data-ttu-id="8e3b5-176">カスタム タブで提供される機能を特定のコンテキストでのみ使用できるようにする必要がある場合は、[カスタム コンテキスト タブ](contextual-tabs.md)を使用します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-176">If the functionality that is provided with a custom tab should only be available in certain contexts, use [custom contextual tabs](contextual-tabs.md).</span></span> <span data-ttu-id="8e3b5-177">カスタム コンテキスト タブを使用する場合は、[カスタム コンテキスト タブをサポートしていないプラットフォームでアドインを実行する場合のフォールバック エクスペリエンス](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)を実装します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-177">If you use custom contextual tabs, make sure to implement a [fallback experience for when your add-in runs on platforms that don't support custom contextual tabs](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

  > [!NOTE]
  > <span data-ttu-id="8e3b5-178">占有領域が大きすぎるアドインは [AppSource 検証](/legal/marketplace/certification-policies)を通過しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-178">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="8e3b5-179">すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-179">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="8e3b5-180">コマンドをサポートしていない Office アプリケーションでも動作するアドインのバージョンを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-180">Provide a version of your add-in that also works on Office applications that do not support commands.</span></span> <span data-ttu-id="8e3b5-181">1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) アプリケーションとコマンド非対応 (作業ウィンドウとして) アプリケーションの両方で動作します。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-181">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) applications.</span></span>

   <span data-ttu-id="8e3b5-182">*図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*</span><span class="sxs-lookup"><span data-stu-id="8e3b5-182">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを比較するスクリーンショット。](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a><span data-ttu-id="8e3b5-185">次の手順</span><span class="sxs-lookup"><span data-stu-id="8e3b5-185">Next steps</span></span>

<span data-ttu-id="8e3b5-186">アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-186">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="8e3b5-187">マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](../reference/manifest/versionoverrides.md)」のリファレンス資料をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8e3b5-187">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
