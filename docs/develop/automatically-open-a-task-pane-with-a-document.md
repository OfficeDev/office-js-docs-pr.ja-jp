---
title: ドキュメントで作業ウィンドウを自動的に開く
description: ドキュメントを開いたときに自動的に Office アドインを開くように構成する方法について説明します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 85b421a569ccb83c3d07f0f10fd4767929332f96
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093708"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="0292d-103">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="0292d-103">Automatically open a task pane with a document</span></span>

<span data-ttu-id="0292d-104">Office アドインでアドインコマンドを使用して、office アプリのリボンにボタンを追加することにより、Office UI を拡張することができます。</span><span class="sxs-lookup"><span data-stu-id="0292d-104">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office app ribbon.</span></span> <span data-ttu-id="0292d-105">ユーザーがコマンド ボタンをクリックすると、アクション (作業ウィンドウを開くなど) が実行されます。</span><span class="sxs-lookup"><span data-stu-id="0292d-105">When users click your command button, an action occurs, such as opening a task pane.</span></span>

<span data-ttu-id="0292d-106">いくつかのシナリオでは、ドキュメントを開いたときに、ユーザーの明示的な操作なしで、自動的に作業ウィンドウを開くことが必要になります。</span><span class="sxs-lookup"><span data-stu-id="0292d-106">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction.</span></span> <span data-ttu-id="0292d-107">AddInCommands 1.1 要件セットに導入されている、作業ウィンドウの Autoopen 機能は、作業ウィンドウを自動的に開く必要があるシナリオで使用できます。</span><span class="sxs-lookup"><span data-stu-id="0292d-107">You can use the autoopen task pane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span>


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="0292d-108">Autoopen 機能と作業ウィンドウの挿入の相違点</span><span class="sxs-lookup"><span data-stu-id="0292d-108">How is the autoopen feature different from inserting a task pane?</span></span>

<span data-ttu-id="0292d-109">ユーザーがアドイン コマンドを使用しないアドイン (Office 2013 で実行するアドインなど) を起動すると、それらはドキュメントに挿入され、そのドキュメントに永続化されます。</span><span class="sxs-lookup"><span data-stu-id="0292d-109">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document.</span></span> <span data-ttu-id="0292d-110">その結果として、別のユーザーがドキュメントを開くと、そのユーザーにアドインのインストールを求めるダイアログが表示され、作業ウィンドウが開きます。</span><span class="sxs-lookup"><span data-stu-id="0292d-110">As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens.</span></span> <span data-ttu-id="0292d-111">このモデルの課題は、多くの場合、ユーザーがドキュメントでアドインを永続化したくないということです。</span><span class="sxs-lookup"><span data-stu-id="0292d-111">The challenge with this model is that in many cases, users don't want the add-in to persist in the document.</span></span> <span data-ttu-id="0292d-112">たとえば、Word ドキュメントで辞書アドインを使用する学生は、そのドキュメントを同級生や教師が開いたときに、アドインのインストールを求めるダイアログが表示されることを望まない場合もあります。</span><span class="sxs-lookup"><span data-stu-id="0292d-112">For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>

<span data-ttu-id="0292d-113">Autoopen 機能では、特定のドキュメントに特定の作業ウィンドウ アドインを永続化させるかどうかをユーザーが明示的に定義できます。</span><span class="sxs-lookup"><span data-stu-id="0292d-113">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span>

## <a name="support-and-availability"></a><span data-ttu-id="0292d-114">サポートと可用性</span><span class="sxs-lookup"><span data-stu-id="0292d-114">Support and availability</span></span>

<span data-ttu-id="0292d-115">autoopen 機能は現在</span><span class="sxs-lookup"><span data-stu-id="0292d-115">The autoopen feature is currently</span></span> <!-- in **developer preview** and it is only --> <span data-ttu-id="0292d-116">次の製品およびプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0292d-116">supported in the following products and platforms.</span></span>

|<span data-ttu-id="0292d-117">**製品**</span><span class="sxs-lookup"><span data-stu-id="0292d-117">**Products**</span></span>|<span data-ttu-id="0292d-118">**プラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="0292d-118">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="0292d-119">Word</span><span class="sxs-lookup"><span data-stu-id="0292d-119">Word</span></span></li><li><span data-ttu-id="0292d-120">Excel</span><span class="sxs-lookup"><span data-stu-id="0292d-120">Excel</span></span></li><li><span data-ttu-id="0292d-121">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0292d-121">PowerPoint</span></span></li></ul>|<span data-ttu-id="0292d-122">すべての製品でサポートされているプラットフォーム: </span><span class="sxs-lookup"><span data-stu-id="0292d-122">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="0292d-123">Office on Windows Desktop.</span><span class="sxs-lookup"><span data-stu-id="0292d-123">Office on Windows Desktop.</span></span> <span data-ttu-id="0292d-124">Build 16.0.8121.1000+</span><span class="sxs-lookup"><span data-stu-id="0292d-124">Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="0292d-125">Office on Mac.</span><span class="sxs-lookup"><span data-stu-id="0292d-125">Office on Mac.</span></span> <span data-ttu-id="0292d-126">Build 15.34.17051500+</span><span class="sxs-lookup"><span data-stu-id="0292d-126">Build 15.34.17051500+</span></span></li><li><span data-ttu-id="0292d-127">Office on the web</span><span class="sxs-lookup"><span data-stu-id="0292d-127">Office on the web</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="0292d-128">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0292d-128">Best practices</span></span>

<span data-ttu-id="0292d-129">Autoopen 機能を使用するときには、次に示すベスト プラクティスを適用してください。</span><span class="sxs-lookup"><span data-stu-id="0292d-129">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="0292d-130">Autoopen 機能は、アドイン ユーザーの作業効率の向上に役立つ場合に使用します。たとえば、次の場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="0292d-130">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
  - <span data-ttu-id="0292d-131">When the document needs the add-in in order to function properly.</span><span class="sxs-lookup"><span data-stu-id="0292d-131">When the document needs the add-in in order to function properly.</span></span> <span data-ttu-id="0292d-132">For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in.</span><span class="sxs-lookup"><span data-stu-id="0292d-132">For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in.</span></span> <span data-ttu-id="0292d-133">The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span><span class="sxs-lookup"><span data-stu-id="0292d-133">The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span>
  - <span data-ttu-id="0292d-134">When the user will most likely always use the add-in with a particular document.</span><span class="sxs-lookup"><span data-stu-id="0292d-134">When the user will most likely always use the add-in with a particular document.</span></span> <span data-ttu-id="0292d-135">For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span><span class="sxs-lookup"><span data-stu-id="0292d-135">For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span>
- <span data-ttu-id="0292d-136">Allow users to turn on or turn off the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="0292d-136">Allow users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="0292d-137">Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span><span class="sxs-lookup"><span data-stu-id="0292d-137">Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="0292d-138">要件セット検出を使用して autoopen 機能が使用可能かどうかを判断し、そうでない場合はフォールバック動作を提供します。</span><span class="sxs-lookup"><span data-stu-id="0292d-138">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn't.</span></span>
- <span data-ttu-id="0292d-139">アドインの使用率を人為的に増やすために、Autoopen 機能を使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="0292d-139">Don't use the autoopen feature to artificially increase usage of your add-in.</span></span> <span data-ttu-id="0292d-140">特定のドキュメントでアドインを自動的に開くことが適切でない場合は、この機能によってユーザーに迷惑を持たせる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="0292d-140">If it doesn't make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span>

    > [!NOTE]
    > <span data-ttu-id="0292d-141">Microsoft では、Autoopen 機能の乱用を見つけた場合は、そのアドインを AppSource から排除することがあります。</span><span class="sxs-lookup"><span data-stu-id="0292d-141">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span>

- <span data-ttu-id="0292d-142">Don't use this feature to pin multiple task panes.</span><span class="sxs-lookup"><span data-stu-id="0292d-142">Don't use this feature to pin multiple task panes.</span></span> <span data-ttu-id="0292d-143">You can only set one pane of your add-in to open automatically with a document.</span><span class="sxs-lookup"><span data-stu-id="0292d-143">You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="0292d-144">実装</span><span class="sxs-lookup"><span data-stu-id="0292d-144">Implementation</span></span>

<span data-ttu-id="0292d-145">Autoopen 機能を実装するには: </span><span class="sxs-lookup"><span data-stu-id="0292d-145">To implement the autoopen feature:</span></span>

- <span data-ttu-id="0292d-146">自動的に開く作業ウィンドウを指定します。</span><span class="sxs-lookup"><span data-stu-id="0292d-146">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="0292d-147">作業ウィンドウを自動的に開くドキュメントにタグ設定します。</span><span class="sxs-lookup"><span data-stu-id="0292d-147">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0292d-148">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device.</span><span class="sxs-lookup"><span data-stu-id="0292d-148">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device.</span></span> <span data-ttu-id="0292d-149">If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored.</span><span class="sxs-lookup"><span data-stu-id="0292d-149">If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored.</span></span> <span data-ttu-id="0292d-150">If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span><span class="sxs-lookup"><span data-stu-id="0292d-150">If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span>

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="0292d-151">手順 1: 開く作業ウィンドウを指定する</span><span class="sxs-lookup"><span data-stu-id="0292d-151">Step 1: Specify the task pane to open</span></span>

<span data-ttu-id="0292d-152">To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**.</span><span class="sxs-lookup"><span data-stu-id="0292d-152">To specify the task pane to open automatically, set the [TaskpaneId](../reference/manifest/action.md#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**.</span></span> <span data-ttu-id="0292d-153">You can only set this value on one task pane.</span><span class="sxs-lookup"><span data-stu-id="0292d-153">You can only set this value on one task pane.</span></span> <span data-ttu-id="0292d-154">If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span><span class="sxs-lookup"><span data-stu-id="0292d-154">If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span>

<span data-ttu-id="0292d-155">次の例では、Office.AutoShowTaskpaneWithDocument に設定された TaskPaneId の値を示しています。</span><span class="sxs-lookup"><span data-stu-id="0292d-155">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="0292d-156">手順 2:作業ウィンドウを自動的に開くよう、ドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="0292d-156">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="0292d-157">You can tag the document to trigger the autoopen feature in one of two ways.</span><span class="sxs-lookup"><span data-stu-id="0292d-157">You can tag the document to trigger the autoopen feature in one of two ways.</span></span> <span data-ttu-id="0292d-158">Pick the alternative that works best for your scenario.</span><span class="sxs-lookup"><span data-stu-id="0292d-158">Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="0292d-159">クライアント側でドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="0292d-159">Tag the document on the client side</span></span>

<span data-ttu-id="0292d-160">Office.js の [settings.set](/javascript/api/office/office.settings) メソッドを使用して、**Office.AutoShowTaskpaneWithDocument** を **true** に設定します。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="0292d-160">Use the Office.js [settings.set](/javascript/api/office/office.settings) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="0292d-161">このメソッドは、アドインの対話式操作の一環としてドキュメントにタグを設定する必要がある場合に使用します (たとえば、ユーザーがバインディングを作成した直後に、または自動的にウィンドウを開くことを示すオプションを選択した直後に使用します)。</span><span class="sxs-lookup"><span data-stu-id="0292d-161">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="0292d-162">Open XML を使用してドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="0292d-162">Use Open XML to tag the document</span></span>

<span data-ttu-id="0292d-163">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="0292d-163">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature.</span></span> <span data-ttu-id="0292d-164">For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span><span class="sxs-lookup"><span data-stu-id="0292d-164">For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span>

<span data-ttu-id="0292d-165">次に示す 2 つの Open XML パートをドキュメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="0292d-165">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="0292d-166">`webextension` パート</span><span class="sxs-lookup"><span data-stu-id="0292d-166">A `webextension` part</span></span>
- <span data-ttu-id="0292d-167">`taskpane` パート</span><span class="sxs-lookup"><span data-stu-id="0292d-167">A `taskpane` part</span></span>

<span data-ttu-id="0292d-168">次の例は、`webextension` パートを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="0292d-168">The following example shows how to add the `webextension` part.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
   <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="0292d-169">`webextension` パートには、プロパティ バッグと **Office.AutoShowTaskpaneWithDocument** という名前のプロパティが含まれています。このプロパティは、`true` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0292d-169">The `webextension` part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.</span></span>

<span data-ttu-id="0292d-170">また、`webextension` パートには、属性が `id`、`storeType`、`store`、および `version` のストアまたはカタログへの参照も含まれています。</span><span class="sxs-lookup"><span data-stu-id="0292d-170">The `webextension` part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`.</span></span> <span data-ttu-id="0292d-171">Autoopen 機能に関連する `storeType` の値は、4 つのみです。</span><span class="sxs-lookup"><span data-stu-id="0292d-171">Of the `storeType` values, only four are relevant to the autoopen feature.</span></span> <span data-ttu-id="0292d-172">その他の 3 つの属性の値は、次の表に示すように、`storeType` の値に応じて決まります。</span><span class="sxs-lookup"><span data-stu-id="0292d-172">The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span>

| <span data-ttu-id="0292d-173">**`storeType` 値**</span><span class="sxs-lookup"><span data-stu-id="0292d-173">**`storeType` value**</span></span> | <span data-ttu-id="0292d-174">**`id` 値**</span><span class="sxs-lookup"><span data-stu-id="0292d-174">**`id` value**</span></span>    |<span data-ttu-id="0292d-175">**`store` 値**</span><span class="sxs-lookup"><span data-stu-id="0292d-175">**`store` value**</span></span> | <span data-ttu-id="0292d-176">**`version` 値**</span><span class="sxs-lookup"><span data-stu-id="0292d-176">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="0292d-177">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="0292d-177">OMEX (AppSource)</span></span>|<span data-ttu-id="0292d-178">アドインの AppSource アセット ID (注を参照)</span><span class="sxs-lookup"><span data-stu-id="0292d-178">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="0292d-179">AppSource のロケール (たとえば、"en-us")。</span><span class="sxs-lookup"><span data-stu-id="0292d-179">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="0292d-180">AppSource カタログのバージョン (注を参照)</span><span class="sxs-lookup"><span data-stu-id="0292d-180">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="0292d-181">FileSystem (ネットワーク共有)</span><span class="sxs-lookup"><span data-stu-id="0292d-181">FileSystem (a network share)</span></span>|<span data-ttu-id="0292d-182">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="0292d-182">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0292d-183">ネットワーク共有のパス。例: "\\\\MyComputer\\MySharedFolder"。</span><span class="sxs-lookup"><span data-stu-id="0292d-183">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="0292d-184">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="0292d-184">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="0292d-185">EXCatalog (Exchange サーバー経由の展開)</span><span class="sxs-lookup"><span data-stu-id="0292d-185">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="0292d-186">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="0292d-186">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0292d-187">"EXCatalog"。</span><span class="sxs-lookup"><span data-stu-id="0292d-187">"EXCatalog".</span></span> <span data-ttu-id="0292d-188">EXCatalog 行は、Microsoft 365 管理センターで一元展開を使用するアドインで使用する行です。</span><span class="sxs-lookup"><span data-stu-id="0292d-188">EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Microsoft 365 admin center.</span></span>|<span data-ttu-id="0292d-189">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="0292d-189">The version in the add-in manifest.</span></span>
|<span data-ttu-id="0292d-190">Registry (システム レジストリ)</span><span class="sxs-lookup"><span data-stu-id="0292d-190">Registry (System registry)</span></span>|<span data-ttu-id="0292d-191">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="0292d-191">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="0292d-192">"developer"</span><span class="sxs-lookup"><span data-stu-id="0292d-192">"developer"</span></span>|<span data-ttu-id="0292d-193">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="0292d-193">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="0292d-194">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in.</span><span class="sxs-lookup"><span data-stu-id="0292d-194">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in.</span></span> <span data-ttu-id="0292d-195">The asset ID appears in the address bar in the browser.</span><span class="sxs-lookup"><span data-stu-id="0292d-195">The asset ID appears in the address bar in the browser.</span></span> <span data-ttu-id="0292d-196">The version is listed in the **Details** section of the page.</span><span class="sxs-lookup"><span data-stu-id="0292d-196">The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="0292d-197">webextension マークアップの詳細については、「[[MS-OWEXML] 2.2.5.WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0292d-197">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="0292d-198">次の例は、`taskpane` パートを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="0292d-198">The following example shows how to add the `taskpane` part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="0292d-199">この例では、`visibility` 属性が "0" に設定されている点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="0292d-199">Note that in this example, the `visibility` attribute is set to "0".</span></span> <span data-ttu-id="0292d-200">これは、webextension パートと `taskpane` パートの追加後に、初めてドキュメントを開いたときに、ユーザーはリボンの **[アドイン]** ボタンからアドインをインストールする必要があることを意味します。</span><span class="sxs-lookup"><span data-stu-id="0292d-200">This means that after the webextension and `taskpane` parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon.</span></span> <span data-ttu-id="0292d-201">それ以降は、ファイルを開いたときに、アドイン作業ウィンドが自動的に開きます。</span><span class="sxs-lookup"><span data-stu-id="0292d-201">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="0292d-202">また、`visibility` を "0" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用できるようにもなります。</span><span class="sxs-lookup"><span data-stu-id="0292d-202">Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="0292d-203">具体的には、スクリプトでドキュメント設定の **Office.AutoShowTaskpaneWithDocument** を `true` または `false` に設定します </span><span class="sxs-lookup"><span data-stu-id="0292d-203">Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`.</span></span> <span data-ttu-id="0292d-204">(詳細については、「[クライアント側でドキュメントにタグを設定する](#tag-the-document-on-the-client-side)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="0292d-204">(For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span>

<span data-ttu-id="0292d-205">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened.</span><span class="sxs-lookup"><span data-stu-id="0292d-205">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened.</span></span> <span data-ttu-id="0292d-206">The user is prompted to trust the add-in, and when trust is granted, the add-in opens.</span><span class="sxs-lookup"><span data-stu-id="0292d-206">The user is prompted to trust the add-in, and when trust is granted, the add-in opens.</span></span> <span data-ttu-id="0292d-207">Thereafter, the add-in task pane opens automatically when the file is opened.</span><span class="sxs-lookup"><span data-stu-id="0292d-207">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="0292d-208">However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span><span class="sxs-lookup"><span data-stu-id="0292d-208">However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span>

<span data-ttu-id="0292d-209">アドインとドキュメントのテンプレートまたはコンテンツが緊密に統合されていて、ユーザーが Autoopen 機能をオフにすることない場合は、`visibility` を "1" に設定することが適切な選択になります。</span><span class="sxs-lookup"><span data-stu-id="0292d-209">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span>

> [!NOTE]
> <span data-ttu-id="0292d-210">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1.</span><span class="sxs-lookup"><span data-stu-id="0292d-210">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1.</span></span> <span data-ttu-id="0292d-211">You can only do this via Open XML.</span><span class="sxs-lookup"><span data-stu-id="0292d-211">You can only do this via Open XML.</span></span>

<span data-ttu-id="0292d-212">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated.</span><span class="sxs-lookup"><span data-stu-id="0292d-212">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated.</span></span> <span data-ttu-id="0292d-213">Office will detect and provide the appropriate attribute values.</span><span class="sxs-lookup"><span data-stu-id="0292d-213">Office will detect and provide the appropriate attribute values.</span></span> <span data-ttu-id="0292d-214">You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span><span class="sxs-lookup"><span data-stu-id="0292d-214">You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-task-panes"></a><span data-ttu-id="0292d-215">作業ウィンドウ表示のテストと検証</span><span class="sxs-lookup"><span data-stu-id="0292d-215">Test and verify opening task panes</span></span>

<span data-ttu-id="0292d-216">Microsoft 365 管理センターを介して一元的な展開を使用して作業ウィンドウを自動的に開くように、アドインのテストバージョンを展開することができます。</span><span class="sxs-lookup"><span data-stu-id="0292d-216">You can deploy a test version of your add-in that will automatically open a task pane using Centralized Deployment via the Microsoft 365 admin center.</span></span> <span data-ttu-id="0292d-217">次の例では、EXCatalog のストア版を使用して一元展開カタログからアドインを挿入する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="0292d-217">The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="0292d-218">前の例をテストするには、Microsoft 365 サブスクリプションを使用して一元展開を試行し、アドインが想定どおりに動作することを確認します。</span><span class="sxs-lookup"><span data-stu-id="0292d-218">You can test the previous example by using your Microsoft 365 subscription to try out Centralized Deployment and verify that your add-in works as expected.</span></span> <span data-ttu-id="0292d-219">Microsoft 365 サブスクリプションをまだお持ちでない場合は、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することによって、更新可能な90日間の microsoft 365 サブスクリプションを無料で入手できます。</span><span class="sxs-lookup"><span data-stu-id="0292d-219">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="see-also"></a><span data-ttu-id="0292d-220">関連項目</span><span class="sxs-lookup"><span data-stu-id="0292d-220">See also</span></span>

<span data-ttu-id="0292d-221">Autoopen 機能の使用方法を示すサンプルについては、「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0292d-221">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span></span>
<span data-ttu-id="0292d-222">[Microsoft 365 開発者プログラムに参加](/office/developer-program/office-365-developer-program)します。</span><span class="sxs-lookup"><span data-stu-id="0292d-222">[Join the Microsoft 365 developer program](/office/developer-program/office-365-developer-program).</span></span>
