---
title: ドキュメントで作業ウィンドウを自動的に開く
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: a231255200d6edd1fc923a82711c8c24819bf914
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870241"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="b2019-102">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="b2019-102">Automatically open a task pane with a document</span></span>

<span data-ttu-id="b2019-p101">Office アドインでアドイン コマンドを使用して、Office リボンにボタンを追加することで Office UI を拡張できます。ユーザーがコマンド ボタンをクリックすると、アクション (作業ウィンドウを開くなど) が実行されます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p101">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane.</span></span>

<span data-ttu-id="b2019-105">いくつかのシナリオでは、ドキュメントを開いたときに、ユーザーの明示的な操作なしで、自動的に作業ウィンドウを開くことが必要になります。</span><span class="sxs-lookup"><span data-stu-id="b2019-105">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction.</span></span> <span data-ttu-id="b2019-106">AddInCommands 1.1 要件セットに導入されている、作業ウィンドウの Autoopen 機能は、作業ウィンドウを自動的に開く必要があるシナリオで使用できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-106">You can use the autoopen task pane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span>


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="b2019-107">Autoopen 機能と作業ウィンドウの挿入の相違点</span><span class="sxs-lookup"><span data-stu-id="b2019-107">How is the autoopen feature different from inserting a task pane?</span></span>

<span data-ttu-id="b2019-p103">ユーザーがアドイン コマンドを使用しないアドイン (Office 2013 で実行するアドインなど) を起動すると、それらはドキュメントに挿入され、そのドキュメントに永続化されます。その結果として、別のユーザーがドキュメントを開くと、そのユーザーにアドインのインストールを求めるダイアログが表示され、作業ウィンドウが開きます。このモデルの問題点は、多くの場合、ユーザーの意に反してドキュメントにアドインが永続化することです。たとえば、Word ドキュメントで辞書アドインを使用する学生は、そのドキュメントを同級生や教師が開いたときに、アドインのインストールを求めるダイアログが表示されることを望まない場合もあります。</span><span class="sxs-lookup"><span data-stu-id="b2019-p103">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>

<span data-ttu-id="b2019-112">Autoopen 機能では、特定のドキュメントに特定の作業ウィンドウ アドインを永続化させるかどうかをユーザーが明示的に定義できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-112">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span>

## <a name="support-and-availability"></a><span data-ttu-id="b2019-113">サポートと可用性</span><span class="sxs-lookup"><span data-stu-id="b2019-113">Support and availability</span></span>

<span data-ttu-id="b2019-114">autoopen 機能は現在</span><span class="sxs-lookup"><span data-stu-id="b2019-114">The autoopen feature is currently  supported in the following products and platforms.</span></span> <!-- in **developer preview** and it is only --> <span data-ttu-id="b2019-115">次の製品およびプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b2019-115">The autoopen feature is currently  supported in the following products and platforms.</span></span>

|<span data-ttu-id="b2019-116">**製品**</span><span class="sxs-lookup"><span data-stu-id="b2019-116">**Products**</span></span>|<span data-ttu-id="b2019-117">**プラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="b2019-117">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="b2019-118">Word</span><span class="sxs-lookup"><span data-stu-id="b2019-118">Word</span></span></li><li><span data-ttu-id="b2019-119">Excel</span><span class="sxs-lookup"><span data-stu-id="b2019-119">Excel</span></span></li><li><span data-ttu-id="b2019-120">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b2019-120">PowerPoint</span></span></li></ul>|<span data-ttu-id="b2019-121">すべての製品でサポートされているプラットフォーム: </span><span class="sxs-lookup"><span data-stu-id="b2019-121">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="b2019-p104">Windows デスクトップ版 Office。ビルド 16.0.8121.1000+</span><span class="sxs-lookup"><span data-stu-id="b2019-p104">Office for Windows Desktop. Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="b2019-p105">Office for Mac。ビルド 15.34.17051500+</span><span class="sxs-lookup"><span data-stu-id="b2019-p105">Office for Mac. Build 15.34.17051500+</span></span></li><li><span data-ttu-id="b2019-126">Office Online</span><span class="sxs-lookup"><span data-stu-id="b2019-126">Office Online</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="b2019-127">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="b2019-127">Best practices</span></span>

<span data-ttu-id="b2019-128">Autoopen 機能を使用するときには、次に示すベスト プラクティスを適用してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-128">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="b2019-129">Autoopen 機能は、アドイン ユーザーの作業効率の向上に役立つ場合に使用します。たとえば、次の場合に使用します。</span><span class="sxs-lookup"><span data-stu-id="b2019-129">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
  - <span data-ttu-id="b2019-p106">適切に機能するには、ドキュメントにアドインが必要になる場合。たとえば、アドインで最新の情報に定期的に更新される在庫の値が含まれているスプレッドシート。最新の値が維持されるように、アドインはスプレッドシートが開かれたときに自動的に開かれる必要があります。</span><span class="sxs-lookup"><span data-stu-id="b2019-p106">When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span>
  - <span data-ttu-id="b2019-p107">特定のドキュメントでユーザーが常にアドインを使用する可能性が高い場合。たとえば、バックエンド システムから情報を取得して、ドキュメントのデータを設定または変更することでユーザーを支援するアドイン。</span><span class="sxs-lookup"><span data-stu-id="b2019-p107">When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span>
- <span data-ttu-id="b2019-p108">Autoopen 機能はユーザーがオン/オフできるようにします。アドインの作業ウィンドウが自動的に起動されないようにするオプションをユーザーの UI に含めます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p108">Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="b2019-137">要件セットの検出を使用して Autoopen 機能が利用可能かどうかを確認し、利用できない場合はフォールバック動作を提供します。</span><span class="sxs-lookup"><span data-stu-id="b2019-137">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.</span></span>
- <span data-ttu-id="b2019-p109">アドインの使用率を人為的に増やすために、Autoopen 機能を使用しないでください。特定のドキュメントでアドインが無意味に自動的に起動すると、ユーザーを不快にすることになります。</span><span class="sxs-lookup"><span data-stu-id="b2019-p109">Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b2019-140">Microsoft では、Autoopen 機能の乱用を見つけた場合は、そのアドインを AppSource から排除することがあります。</span><span class="sxs-lookup"><span data-stu-id="b2019-140">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span>

- <span data-ttu-id="b2019-p110">この機能は、複数の作業ウィンドウを固定するために使用しないでください。1 つのドキュメントで自動的に開くアドインのウィンドウは 1 つのみ設定できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p110">Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="b2019-143">実装</span><span class="sxs-lookup"><span data-stu-id="b2019-143">Implementation</span></span>

<span data-ttu-id="b2019-144">Autoopen 機能を実装するには: </span><span class="sxs-lookup"><span data-stu-id="b2019-144">To implement the autoopen feature:</span></span>

- <span data-ttu-id="b2019-145">自動的に開く作業ウィンドウを指定します。</span><span class="sxs-lookup"><span data-stu-id="b2019-145">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="b2019-146">作業ウィンドウを自動的に開くドキュメントにタグ設定します。</span><span class="sxs-lookup"><span data-stu-id="b2019-146">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b2019-p111">自動的に開くように指定したウィンドウは、アドインがユーザーのデバイスに既にインストールされている場合にのみ開きます。ユーザーがドキュメントを開いたときに、アドインがインストールされていない場合、Autoopen 機能は動作せずに、設定は無視されます。また、アドインをドキュメントと共に配布する必要がある場合は、可視性プロパティを 1 に設定する必要があります。これは、OpenXML を使用する場合にのみ実行できます。例については、この記事で後述します。</span><span class="sxs-lookup"><span data-stu-id="b2019-p111">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span>

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="b2019-150">手順 1: 開く作業ウィンドウを指定する</span><span class="sxs-lookup"><span data-stu-id="b2019-150">Step 1: Specify the task pane to open</span></span>

<span data-ttu-id="b2019-p112">自動的に開く作業ウィンドウを指定するには、[TaskpaneId](/office/dev/add-ins/reference/manifest/action#taskpaneid) の値を **Office.AutoShowTaskpaneWithDocument** に設定します。この値は 1 つの作業ウィンドウにのみ設定できます。この値を複数の作業ウィンドウに設定すると、最初に見つかった値が認識され、その他は無視されます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p112">To specify the task pane to open automatically, set the [TaskpaneId](/office/dev/add-ins/reference/manifest/action#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span>

<span data-ttu-id="b2019-154">次の例では、Office.AutoShowTaskpaneWithDocument に設定された TaskPaneId の値を示しています。</span><span class="sxs-lookup"><span data-stu-id="b2019-154">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="b2019-155">手順 2:作業ウィンドウを自動的に開くよう、ドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="b2019-155">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="b2019-p113">Autoopen 機能をトリガーするよう、2 つのうちどちらかの方法でドキュメントにタグを設定できます。 シナリオに最も適した方法を選択します。</span><span class="sxs-lookup"><span data-stu-id="b2019-p113">You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="b2019-158">クライアント側でドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="b2019-158">Tag the document on the client side</span></span>

<span data-ttu-id="b2019-159">Office.js の [settings.set](/javascript/api/office/office.settings) メソッドを使用して、**Office.AutoShowTaskpaneWithDocument** を **true** に設定します。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="b2019-159">Use the Office.js [settings.set](/javascript/api/office/office.settings) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="b2019-160">このメソッドは、アドインの対話式操作の一環としてドキュメントにタグを設定する必要がある場合に使用します (たとえば、ユーザーがバインディングを作成した直後に、または自動的にウィンドウを開くことを示すオプションを選択した直後に使用します)。</span><span class="sxs-lookup"><span data-stu-id="b2019-160">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="b2019-161">Open XML を使用してドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="b2019-161">Use Open XML to tag the document</span></span>

<span data-ttu-id="b2019-p114">Open XML を使用すると、Autoopen 機能をトリガーするために、ドキュメントを作成または変更して、適切な Open Office XML マークアップを追加できます。この方法を示すサンプルについては、「[Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-p114">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span>

<span data-ttu-id="b2019-164">次に示す 2 つの Open XML パートをドキュメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="b2019-164">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="b2019-165">`webextension` パート</span><span class="sxs-lookup"><span data-stu-id="b2019-165">A `webextension` part</span></span>
- <span data-ttu-id="b2019-166">`taskpane` パート</span><span class="sxs-lookup"><span data-stu-id="b2019-166">A `taskpane` part</span></span>

<span data-ttu-id="b2019-167">次の例は、`webextension` パートを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b2019-167">The following example shows how to add the `webextension` part.</span></span>

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

<span data-ttu-id="b2019-168">`webextension` パートには、プロパティ バッグと **Office.AutoShowTaskpaneWithDocument** という名前のプロパティが含まれています。このプロパティは、`true` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b2019-168">The `webextension` part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.</span></span>

<span data-ttu-id="b2019-169">また、`webextension` パートには、属性が `id`、`storeType`、`store`、および `version` のストアまたはカタログへの参照も含まれています。</span><span class="sxs-lookup"><span data-stu-id="b2019-169">The `webextension` part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`.</span></span> <span data-ttu-id="b2019-170">Autoopen 機能に関連する `storeType` の値は、4 つのみです。</span><span class="sxs-lookup"><span data-stu-id="b2019-170">Of the `storeType` values, only four are relevant to the autoopen feature.</span></span> <span data-ttu-id="b2019-171">その他の 3 つの属性の値は、次の表に示すように、`storeType` の値に応じて決まります。</span><span class="sxs-lookup"><span data-stu-id="b2019-171">The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span>

| <span data-ttu-id="b2019-172">**`storeType` 値**</span><span class="sxs-lookup"><span data-stu-id="b2019-172">**`storeType` value**</span></span> | <span data-ttu-id="b2019-173">**`id` 値**</span><span class="sxs-lookup"><span data-stu-id="b2019-173">**`id` value**</span></span>    |<span data-ttu-id="b2019-174">**`store` 値**</span><span class="sxs-lookup"><span data-stu-id="b2019-174">**`store` value**</span></span> | <span data-ttu-id="b2019-175">**`version` 値**</span><span class="sxs-lookup"><span data-stu-id="b2019-175">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="b2019-176">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="b2019-176">OMEX (AppSource)</span></span>|<span data-ttu-id="b2019-177">アドインの AppSource アセット ID (注を参照)</span><span class="sxs-lookup"><span data-stu-id="b2019-177">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="b2019-178">AppSource のロケール (たとえば、"en-us")。</span><span class="sxs-lookup"><span data-stu-id="b2019-178">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="b2019-179">AppSource カタログのバージョン (注を参照)</span><span class="sxs-lookup"><span data-stu-id="b2019-179">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="b2019-180">FileSystem (ネットワーク共有)</span><span class="sxs-lookup"><span data-stu-id="b2019-180">FileSystem (a network share)</span></span>|<span data-ttu-id="b2019-181">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="b2019-181">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="b2019-182">ネットワーク共有のパス。例: "\\\\MyComputer\\MySharedFolder"。</span><span class="sxs-lookup"><span data-stu-id="b2019-182">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="b2019-183">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="b2019-183">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="b2019-184">EXCatalog (Exchange サーバー経由の展開)</span><span class="sxs-lookup"><span data-stu-id="b2019-184">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="b2019-185">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="b2019-185">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="b2019-186">"EXCatalog"。</span><span class="sxs-lookup"><span data-stu-id="b2019-186">"EXCatalog".</span></span> <span data-ttu-id="b2019-187">EXCatalog 行は、Office 365 管理センターで一元展開を使用するアドインで使用する行です。</span><span class="sxs-lookup"><span data-stu-id="b2019-187">EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Office 365 admin center.</span></span>|<span data-ttu-id="b2019-188">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="b2019-188">The version in the add-in manifest.</span></span>
|<span data-ttu-id="b2019-189">Registry (システム レジストリ)</span><span class="sxs-lookup"><span data-stu-id="b2019-189">Registry (System registry)</span></span>|<span data-ttu-id="b2019-190">アドイン マニフェストでのアドインの GUID。</span><span class="sxs-lookup"><span data-stu-id="b2019-190">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="b2019-191">"developer"</span><span class="sxs-lookup"><span data-stu-id="b2019-191">"developer"</span></span>|<span data-ttu-id="b2019-192">アドイン マニフェストでのバージョン。</span><span class="sxs-lookup"><span data-stu-id="b2019-192">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="b2019-p117">AppSource でのアドインのアセット ID とバージョンを確認するには、そのアドインの AppSource ランディング ページに移動します。アセット ID は、ブラウザーのアドレス バーに表示されます。バージョンは、そのページの **[詳細]** セクションに示されます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p117">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="b2019-196">webextension マークアップの詳細については、「[[MS-OWEXML] 2.2.5.WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-196">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="b2019-197">次の例は、`taskpane` パートを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="b2019-197">The following example shows how to add the `taskpane` part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="b2019-198">この例では、`visibility` 属性が "0" に設定されている点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-198">Note that in this example, the `visibility` attribute is set to "0".</span></span> <span data-ttu-id="b2019-199">これは、webextension パートと `taskpane` パートの追加後に、初めてドキュメントを開いたときに、ユーザーはリボンの **[アドイン]** ボタンからアドインをインストールする必要があることを意味します。</span><span class="sxs-lookup"><span data-stu-id="b2019-199">This means that after the webextension and `taskpane` parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon.</span></span> <span data-ttu-id="b2019-200">それ以降は、ファイルを開いたときに、アドイン作業ウィンドが自動的に開きます。</span><span class="sxs-lookup"><span data-stu-id="b2019-200">Thereafter, the add-in task pane opens automatically when the file is opened.</span></span> <span data-ttu-id="b2019-201">また、`visibility` を "0" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用できるようにもなります。</span><span class="sxs-lookup"><span data-stu-id="b2019-201">Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> <span data-ttu-id="b2019-202">具体的には、スクリプトでドキュメント設定の **Office.AutoShowTaskpaneWithDocument** を `true` または `false` に設定します </span><span class="sxs-lookup"><span data-stu-id="b2019-202">Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`.</span></span> <span data-ttu-id="b2019-203">(詳細については、「[クライアント側でドキュメントにタグを設定する](#tag-the-document-on-the-client-side)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="b2019-203">(For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span>

<span data-ttu-id="b2019-p119">`visibility` が "1" に設定されていると、初めてドキュメントを開いたときに、自動的に作業ウィンドウが開きます。アドインを信頼するように求めるダイアログがユーザーに表示され、信頼が付与されるとアドインが開きます。それ以降は、ファイルを開いたときに、アドイン作業ウィンドが自動的に開きます。ただし、`visibility` を "1" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用することができなくなります。</span><span class="sxs-lookup"><span data-stu-id="b2019-p119">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span>

<span data-ttu-id="b2019-208">アドインとドキュメントのテンプレートまたはコンテンツが緊密に統合されていて、ユーザーが Autoopen 機能をオフにすることない場合は、`visibility` を "1" に設定することが適切な選択になります。</span><span class="sxs-lookup"><span data-stu-id="b2019-208">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span>

> [!NOTE]
> <span data-ttu-id="b2019-p120">ドキュメントとともにアドインを配布する場合は、ユーザーにアドインをインストールするように求めるために、visibility プロパティを 1 に設定する必要があります。これは、Open XML でのみ実行できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p120">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.</span></span>

<span data-ttu-id="b2019-p121">この XML を簡単に記述する 1 つの方法として、最初にアドインを実行し、値を書き込むために[クライアント側でドキュメントにタグを設定](#tag-the-document-on-the-client-side)して、ドキュメントを保存してから生成された XML を調べます。Office により、適切な属性値が検出されて設定されます。また、[Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) ツールを使用して生成した C# コードにより、生成する XML 基づいてプログラムでマークアップを追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="b2019-p121">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-task-panes"></a><span data-ttu-id="b2019-214">作業ウィンドウ表示のテストと検証</span><span class="sxs-lookup"><span data-stu-id="b2019-214">Test and verify opening task panes</span></span>

<span data-ttu-id="b2019-215">自動的に作業ウィンドウを開くアドインのテスト バージョンは、Office 365 管理センターを経由した一元展開を使用して展開できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-215">You can deploy a test version of your add-in that will automatically open a task pane using Centralized Deployment via the Office 365 admin center.</span></span> <span data-ttu-id="b2019-216">次の例では、EXCatalog のストア版を使用して一元展開カタログからアドインを挿入する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="b2019-216">The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

<span data-ttu-id="b2019-217">前の例をテストするには、[Office 365 開発者プログラム](/office/developer-program/office-365-developer-program)に参加することを検討してください。Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 開発者アカウント](https://developer.microsoft.com/office/dev-program)にサインアップすることを検討してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-217">To test the previous example, please consider joining the [Office 365 Developer Program](/office/developer-program/office-365-developer-program) and signing up for an [Office 365 developer account](https://developer.microsoft.com/office/dev-program) if you don't already own an Office 365 subscription.</span></span> <span data-ttu-id="b2019-218">実際に一元展開をテストし、アドインが期待どおりに動作することを確認できます。</span><span class="sxs-lookup"><span data-stu-id="b2019-218">You can actually test drive Centralized Deployment and verify that your add-in works as expected.</span></span>


## <a name="see-also"></a><span data-ttu-id="b2019-219">関連項目</span><span class="sxs-lookup"><span data-stu-id="b2019-219">See also</span></span>

<span data-ttu-id="b2019-220">Autoopen 機能の使用方法を示すサンプルについては、「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2019-220">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane).</span></span>
<span data-ttu-id="b2019-221">[Office 365 開発者プログラムに参加してください](/office/developer-program/office-365-developer-program)。</span><span class="sxs-lookup"><span data-stu-id="b2019-221">[Join the Office 365 developer program](/office/developer-program/office-365-developer-program).</span></span>
