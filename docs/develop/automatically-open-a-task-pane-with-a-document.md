---
title: ドキュメントで作業ウィンドウを自動的に開く
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 2ebce1ce8bd95ee7802b5509d375f1986bb2877e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505917"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a><span data-ttu-id="13acc-102">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="13acc-102">Automatically open a task pane with a document</span></span>

<span data-ttu-id="13acc-p101">Office アドインでアドイン コマンドを使用すると、Office リボンにボタンを追加して Office UI を拡張できます。ユーザーがコマンド ボタンをクリックすると、アクション (作業ウィンドウを開くなど) が実行されます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p101">You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane.</span></span> 

<span data-ttu-id="13acc-p102">一部のシナリオでは、ドキュメントを開いたときに、ユーザーが明示的に操作を行うことなく、自動的に作業ウィンドウを開く必要があります。AddInCommands 1.1 要件セットに導入されている作業ウィンドウの Autoopen 機能は、作業ウィンドウを自動的に開く必要があるシナリオで使用できます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p102">Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction. You can use the autoopen taskpane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it.</span></span> 


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a><span data-ttu-id="13acc-107">Autoopen 機能と作業ウィンドウの挿入の相違点</span><span class="sxs-lookup"><span data-stu-id="13acc-107">How is the autoopen feature different from inserting a task pane?</span></span> 

<span data-ttu-id="13acc-p103">ユーザーがアドイン コマンドを使用しないアドイン (Office 2013 で実行するアドインなど) を起動すると、そのアドインはドキュメントに挿入され保持されます。その結果、別のユーザーがドキュメントを開くと、そのユーザーにアドインのインストールを求めるダイアログが表示され作業ウィンドウが開きます。このモデルの問題点は、多くの場合においてユーザーの意に反してドキュメントにアドインが保持されてしまうことです。たとえば、Word ドキュメントで辞書アドインを使用する学生は、そのドキュメントを同級生や教師が開いたときにアドインのインストールを求めるダイアログが表示されることを望まない場合もあります。</span><span class="sxs-lookup"><span data-stu-id="13acc-p103">When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.</span></span>  

<span data-ttu-id="13acc-112">Autoopen 機能では、特定のドキュメントに特定の作業ウィンドウ アドインを保持させるかどうかをユーザーが明示的に定義できます。</span><span class="sxs-lookup"><span data-stu-id="13acc-112">With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document.</span></span> 

## <a name="support-and-availability"></a><span data-ttu-id="13acc-113">サポートと可用性</span><span class="sxs-lookup"><span data-stu-id="13acc-113">Support and availability</span></span>
<span data-ttu-id="13acc-114">現時点では、Autoopen 機能は次の製品およびプラットフォームで<!-- in **developer preview** and it is only -->サポートされています。</span><span class="sxs-lookup"><span data-stu-id="13acc-114">The autoopen feature is currently <!-- in **developer preview** and it is only --> supported in the following products and platforms.</span></span>

|<span data-ttu-id="13acc-115">**製品**</span><span class="sxs-lookup"><span data-stu-id="13acc-115">**Products**</span></span>|<span data-ttu-id="13acc-116">**プラットフォーム**</span><span class="sxs-lookup"><span data-stu-id="13acc-116">**Platforms**</span></span>|
|:-----------|:------------|
|<ul><li><span data-ttu-id="13acc-117">Word</span><span class="sxs-lookup"><span data-stu-id="13acc-117">Word</span></span></li><li><span data-ttu-id="13acc-118">Excel</span><span class="sxs-lookup"><span data-stu-id="13acc-118">Excel</span></span></li><li><span data-ttu-id="13acc-119">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="13acc-119">PowerPoint</span></span></li></ul>|<span data-ttu-id="13acc-120">すべての製品でサポートされているプラ​​ットフォーム。</span><span class="sxs-lookup"><span data-stu-id="13acc-120">Supported platforms for all products:</span></span><ul><li><span data-ttu-id="13acc-p104">Windows デスクトップ版 Office (ビルド 16.0.8121.1000 以降)</span><span class="sxs-lookup"><span data-stu-id="13acc-p104">Office for Windows Desktop. Build 16.0.8121.1000+</span></span></li><li><span data-ttu-id="13acc-p105">Office for Mac (ビルド 15.34.17051500 以降)</span><span class="sxs-lookup"><span data-stu-id="13acc-p105">Office for Mac. Build 15.34.17051500+</span></span></li><li><span data-ttu-id="13acc-125">Office Online</span><span class="sxs-lookup"><span data-stu-id="13acc-125">Office Online</span></span></li></ul>|


## <a name="best-practices"></a><span data-ttu-id="13acc-126">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="13acc-126">Best practices</span></span>

<span data-ttu-id="13acc-127">Autoopen 機能を使用する際には、次に示すベスト プラクティスを適用してください。</span><span class="sxs-lookup"><span data-stu-id="13acc-127">Apply the following best practices when you use the autoopen feature:</span></span>

- <span data-ttu-id="13acc-128">Autoopen 機能は、アドイン ユーザーの作業効率の向上に役立つ場合に使用します。いくつかの例を示します。</span><span class="sxs-lookup"><span data-stu-id="13acc-128">Use the autoopen feature when it will help make your add-in users more efficient, such as:</span></span>
    - <span data-ttu-id="13acc-p106">ドキュメントが適切に機能するためにアドインを必要とする場合。たとえば、アドインで定期的に最新の株価を更新するスプレッドシートでは、最新の値を維持するためにスプレッドシートが開かれたときにアドインが自動的に開かれる必要があります。</span><span class="sxs-lookup"><span data-stu-id="13acc-p106">When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.</span></span> 
    - <span data-ttu-id="13acc-p107">ユーザーが特定のドキュメントで常にアドインを使用する可能性が高い場合。たとえば、バックエンド システムから情報を取得してドキュメントのデータを入力または変更することでユーザーを支援するアドインです。</span><span class="sxs-lookup"><span data-stu-id="13acc-p107">When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.</span></span> 
- <span data-ttu-id="13acc-p108">Autoopen 機能をユーザーがオン/オフできるようにします。ユーザーの UI に、アドインの作業ウィンドウが自動的に起動されないようにするオプションを含めます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p108">Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.</span></span>  
- <span data-ttu-id="13acc-136">要件セットの検出を使用して Autoopen 機能が利用可能かどうかを確認し、利用できない場合のフォールバック処理を用意します。</span><span class="sxs-lookup"><span data-stu-id="13acc-136">Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.</span></span>
- <span data-ttu-id="13acc-p109">アドインの使用率を人為的に増やすために、Autoopen 機能を使用しないでください。特定のドキュメントでアドインが意味もなく自動的に起動することはユーザーの妨げになります。</span><span class="sxs-lookup"><span data-stu-id="13acc-p109">Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, this feature can annoy users.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="13acc-139">Microsoft が Autoopen 機能の乱用を見つけた場合には、そのアドインを AppSource から排除する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="13acc-139">If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource.</span></span> 

- <span data-ttu-id="13acc-p110">この機能は、複数の作業ウィンドウを固定するために使用しないでください。1 つのドキュメントで自動的に開くアドインのウィンドウは 1 つのみ設定できます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p110">Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.</span></span>  

## <a name="implementation"></a><span data-ttu-id="13acc-142">実装</span><span class="sxs-lookup"><span data-stu-id="13acc-142">Implementation</span></span>
<span data-ttu-id="13acc-143">Autoopen 機能は次のように実装します。</span><span class="sxs-lookup"><span data-stu-id="13acc-143">To implement the autoopen feature:</span></span>

- <span data-ttu-id="13acc-144">自動的に開く作業ウィンドウを指定します。</span><span class="sxs-lookup"><span data-stu-id="13acc-144">Specify the task pane to be opened automatically.</span></span>
- <span data-ttu-id="13acc-145">作業ウィンドウを自動的に開くドキュメントにタグを設定します。</span><span class="sxs-lookup"><span data-stu-id="13acc-145">Tag the document to automatically open the task pane.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="13acc-p111">自動的に開くように指定したウィンドウは、アドインがユーザーのデバイスに既にインストールされている場合にのみ開きます。ユーザーがドキュメントを開いたときに、アドインがインストールされていない場合、Autoopen 機能は動作せず、設定は無視されます。また、アドインをドキュメントと共に配布する必要がある場合は、可視性プロパティを 1 に設定する必要があります。これは、OpenXML を使用する場合にのみ実行できます。例については、この記事で後述します。</span><span class="sxs-lookup"><span data-stu-id="13acc-p111">The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.</span></span> 

### <a name="step-1-specify-the-task-pane-to-open"></a><span data-ttu-id="13acc-149">手順 1: 開く作業ウィンドウを指定する</span><span class="sxs-lookup"><span data-stu-id="13acc-149">Step 1: Specify the task pane to open</span></span>
<span data-ttu-id="13acc-p112">自動的に開く作業ウィンドウを指定するには、[TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) の値を **Office.AutoShowTaskpaneWithDocument** に設定します。この値は 1 つの作業ウィンドウにのみ設定できます。この値を複数の作業ウィンドウに設定すると、最初に見つかった値のみを認識し、その他は無視されます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p112">To specify the task pane to open automatically, set the [TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.</span></span> 

<span data-ttu-id="13acc-153">次の例は、TaskPaneId の値を Office.AutoShowTaskpaneWithDocument に設定しています。</span><span class="sxs-lookup"><span data-stu-id="13acc-153">The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.</span></span>
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a><span data-ttu-id="13acc-154">手順 2: 作業ウィンドウを自動的に開くドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="13acc-154">Step 2: Tag the document to automatically open the task pane</span></span>

<span data-ttu-id="13acc-p113">ドキュメントにタグを設定し Autoopen 機能をトリガーする方法は 2 つあります。シナリオに最も適した方法を選択してください。</span><span class="sxs-lookup"><span data-stu-id="13acc-p113">You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.</span></span>  


#### <a name="tag-the-document-on-the-client-side"></a><span data-ttu-id="13acc-157">クライアント側でドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="13acc-157">Tag the document on the client side</span></span>
<span data-ttu-id="13acc-158">Office.js の [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) メソッドを使用して、**Office.AutoShowTaskpaneWithDocument** を **true** に設定します。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="13acc-158">Use the Office.js [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.</span></span>   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

<span data-ttu-id="13acc-159">このメソッドは、アドインの対話型操作の一部としてドキュメントにタグを設定する必要がある場合に使用します (たとえば、ユーザーがバインディングを作成した直後、または自動的にウィンドウを開くことを示すオプションを選択した直後にウィンドウを開く場合に使用します)。</span><span class="sxs-lookup"><span data-stu-id="13acc-159">Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).</span></span>

#### <a name="use-open-xml-to-tag-the-document"></a><span data-ttu-id="13acc-160">Open XML を使用してドキュメントにタグを設定する</span><span class="sxs-lookup"><span data-stu-id="13acc-160">Use Open XML to tag the document</span></span>
<span data-ttu-id="13acc-p114">Open XML を使用して、ドキュメントを作成または変更し、Autoopen 機能をトリガーするために必要な Open Office XML マークアップを追加できます。この方法を示すサンプルについては、「[Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13acc-p114">You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).</span></span> 

<span data-ttu-id="13acc-163">次の 2 つの Open XML パートをドキュメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="13acc-163">Add two Open XML parts to the document:</span></span>

- <span data-ttu-id="13acc-164">webextension パート</span><span class="sxs-lookup"><span data-stu-id="13acc-164">A webextension part</span></span>
- <span data-ttu-id="13acc-165">taskpane パート</span><span class="sxs-lookup"><span data-stu-id="13acc-165">A task pane part</span></span>

<span data-ttu-id="13acc-166">次の例は、webextension パートを追加する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="13acc-166">The following example shows how to add the webextension part.</span></span>

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

<span data-ttu-id="13acc-167">webextension パートには、プロパティ バッグと **Office.AutoShowTaskpaneWithDocument** という名前のプロパティが含まれています。このプロパティは、`true` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="13acc-167">The webextension part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.</span></span>

<span data-ttu-id="13acc-p115">また、webextension パートには、ストアまたはカタログへの参照となる `id`、`storeType`、`store`、`version` の属性も含まれています。`storeType` 属性が Autoopen 機能に関連して持つ値は 4 つのみです。その他の 3 つの属性の値は、次の表に示すように、`storeType` の値に応じて決まります。</span><span class="sxs-lookup"><span data-stu-id="13acc-p115">The webextension part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`. Of the `storeType` values, only four are relevant to the autoopen feature. The values for the other three attributes depend on the value for `storeType`, as shown in the following table.</span></span> 

| <span data-ttu-id="13acc-171">**`storeType` 値**</span><span class="sxs-lookup"><span data-stu-id="13acc-171">**`storeType` value**</span></span> | <span data-ttu-id="13acc-172">**`id` 値**</span><span class="sxs-lookup"><span data-stu-id="13acc-172">**`id` value**</span></span>    |<span data-ttu-id="13acc-173">**`store` 値**</span><span class="sxs-lookup"><span data-stu-id="13acc-173">**`store` value**</span></span> | <span data-ttu-id="13acc-174">**`version` 値**</span><span class="sxs-lookup"><span data-stu-id="13acc-174">**`version` value**</span></span>|
|:---------------|:---------------|:---------------|:---------------|
|<span data-ttu-id="13acc-175">OMEX (AppSource)</span><span class="sxs-lookup"><span data-stu-id="13acc-175">OMEX (AppSource)</span></span>|<span data-ttu-id="13acc-176">アドインの AppSource 資産 ID (注を参照)</span><span class="sxs-lookup"><span data-stu-id="13acc-176">The AppSource asset ID of the add-in (see Note)</span></span>|<span data-ttu-id="13acc-177">AppSource のロケール (例: "en-us")</span><span class="sxs-lookup"><span data-stu-id="13acc-177">The locale of AppSource; for example, "en-us".</span></span>|<span data-ttu-id="13acc-178">AppSource カタログのバージョン (注を参照)</span><span class="sxs-lookup"><span data-stu-id="13acc-178">The version in the AppSource catalog (see Note)</span></span>|
|<span data-ttu-id="13acc-179">FileSystem (ネットワーク共有)</span><span class="sxs-lookup"><span data-stu-id="13acc-179">FileSystem (a network share)</span></span>|<span data-ttu-id="13acc-180">アドイン マニフェスト内のアドインの GUID</span><span class="sxs-lookup"><span data-stu-id="13acc-180">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="13acc-181">ネットワーク共有のパス。例: "\\\\MyComputer\\MySharedFolder"</span><span class="sxs-lookup"><span data-stu-id="13acc-181">The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span>|<span data-ttu-id="13acc-182">アドイン マニフェスト内のバージョン</span><span class="sxs-lookup"><span data-stu-id="13acc-182">The version in the add-in manifest.</span></span>|
|<span data-ttu-id="13acc-183">EXCatalog (Exchange Server 経由の展開)</span><span class="sxs-lookup"><span data-stu-id="13acc-183">EXCatalog (deployment via the Exchange server)</span></span> |<span data-ttu-id="13acc-184">アドイン マニフェスト内のアドインの GUID</span><span class="sxs-lookup"><span data-stu-id="13acc-184">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="13acc-p116">"EXCatalog"。EXCatalog 行は、Office 365 管理センターで一元展開を使用するアドインで使用する行です</span><span class="sxs-lookup"><span data-stu-id="13acc-p116">"EXCatalog". EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Office 365 admin center.</span></span>|<span data-ttu-id="13acc-187">アドイン マニフェスト内のバージョン</span><span class="sxs-lookup"><span data-stu-id="13acc-187">The version in the add-in manifest.</span></span>
|<span data-ttu-id="13acc-188">Registry (システム レジストリ)</span><span class="sxs-lookup"><span data-stu-id="13acc-188">Registry (System registry)</span></span>|<span data-ttu-id="13acc-189">アドイン マニフェスト内のアドインの GUID</span><span class="sxs-lookup"><span data-stu-id="13acc-189">The GUID of the add-in in the add-in manifest.</span></span>|<span data-ttu-id="13acc-190">"developer"</span><span class="sxs-lookup"><span data-stu-id="13acc-190">"developer"</span></span>|<span data-ttu-id="13acc-191">アドイン マニフェスト内のバージョン</span><span class="sxs-lookup"><span data-stu-id="13acc-191">The version in the add-in manifest.</span></span>|

> [!NOTE]
> <span data-ttu-id="13acc-p117">AppSource でのアドインの資産 ID とバージョンを確認するには、そのアドインの AppSource ランディング ページに移動します。資産 ID は、ブラウザのアドレス バーに表示されます。バージョンは、そのページの **[詳細]** セクションに示されます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p117">To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.</span></span>

<span data-ttu-id="13acc-195">webextension マークアップの詳細については、「[[MS-OWEXML] 2.2.5.WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13acc-195">For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).</span></span>

<span data-ttu-id="13acc-196">次の例は、taskpane パートを追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="13acc-196">The following example shows how to add the taskpane part.</span></span>

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

<span data-ttu-id="13acc-p118">この例では、`visibility` 属性が "0" に設定されている点に注目してください。これは、webextension パートと taskpane パートの追加後に、初めてドキュメントを開いたときに、ユーザーはリボンの **[アドイン]** ボタンからアドインをインストールする必要があることを意味します。それ以降は、ファイルを開くとアドイン作業ウィンドが自動的に開きます。また、`visibility` を "0" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用できるようにもなります。具体的には、スクリプトでドキュメント設定の **Office.AutoShowTaskpaneWithDocument** を `true` または `false` に設定します (詳細については、「[クライアント側でドキュメントにタグを設定する](#tag-the-document-on-the-client-side)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="13acc-p118">Note that in this example, the `visibility` attribute is set to "0". This means that after the webextension and taskpane parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon. Thereafter, the add-in task pane opens automatically when the file is opened. Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature. Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`. (For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).)</span></span> 

<span data-ttu-id="13acc-p119">`visibility` が "1" に設定されていると、初めてドキュメントを開いたときに、自動的に作業ウィンドウが開きます。アドインを信頼することを求めるダイアログがユーザーに表示され、信頼が付与されるとアドインが開きます。それ以降は、ファイルを開くとアドイン作業ウィンドが自動的に開きます。ただし、`visibility` を "1" に設定すると、ユーザーが Autoopen 機能をオン/オフできるようにするために Office.js を使用することができなくなります。</span><span class="sxs-lookup"><span data-stu-id="13acc-p119">If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.</span></span> 

<span data-ttu-id="13acc-207">アドインとドキュメントのテンプレートまたはコンテンツが緊密に統合されているために、ユーザーが Autoopen 機能をオフにすることがない場合には、`visibility` を "1" に設定することが適切な選択になります。</span><span class="sxs-lookup"><span data-stu-id="13acc-207">Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature.</span></span> 

> [!NOTE]
> <span data-ttu-id="13acc-p120">ドキュメントとともに配布するアドインのインストールをユーザーに求めるためには、visibility プロパティを 1 に設定する必要があります。これは、Open XML でのみ実行できます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p120">If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.</span></span>

<span data-ttu-id="13acc-p121">この XML を記述する簡単な方法は、最初にアドインを実行し、値を書き込むために[クライアント側でドキュメントにタグを設定](#tag-the-document-on-the-client-side)し、ドキュメントを保存してから生成された XML を調べる方法です。この方法では、Office は適切な属性値を検出し設定します。また、[Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) を使用して生成した C# コードにより、生成する XML  に基づくプログラムでマークアップを追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p121">An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.</span></span>

## <a name="test-and-verify-opening-taskpanes"></a><span data-ttu-id="13acc-213">作業ウィンドウ表示のテストと検証</span><span class="sxs-lookup"><span data-stu-id="13acc-213">Test and verify opening taskpanes</span></span>
<span data-ttu-id="13acc-p122">自動的に作業ウィンドウを開くアドインのテスト バージョンは、Office 365 管理センターによる一元展開を使用して展開できます。次の例は、EXCatalog のストア版を使用して一元展開カタログからアドインを挿入する方法を示すものです。</span><span class="sxs-lookup"><span data-stu-id="13acc-p122">You can deploy a test version of your add-in that will automatically open a taskpane using Centralized Deployment via the Office 365 admin center. The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.</span></span>

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
<span data-ttu-id="13acc-p123">前の例をテストするために、[Office 365 開発者プログラム](https://docs.microsoft.com/office/developer-program/office-365-developer-program) に参加し、Office 365 サブスクリプションを購入していない場合は、[Office 365 開発者アカウント](https://developer.microsoft.com/office/dev-program) にサインアップすることを検討してください。実際に一元展開をテストし、アドインが期待どおりに動作することを確認できます。</span><span class="sxs-lookup"><span data-stu-id="13acc-p123">To test the previous example, please consider joining the [Office 365 Developer Program](https://docs.microsoft.com/office/developer-program/office-365-developer-program) and signing up for an [Office 365 developer account](https://developer.microsoft.com/office/dev-program) if you don't already own an Office 365 subscription. You can actually test drive Centralized Deployment and verify that your add-in works as expected.</span></span>


## <a name="see-also"></a><span data-ttu-id="13acc-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="13acc-218">See also</span></span>

<span data-ttu-id="13acc-p124">Autoopen 機能の使用方法を示すサンプルについては、 [Office アドイン コマンドのサンプル](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)を参照してください。 [Office 365 開発者プログラムに参加](https://docs.microsoft.com/office/developer-program/office-365-developer-program)します。</span><span class="sxs-lookup"><span data-stu-id="13acc-p124">For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane). [Join the Office 365 developer program](https://docs.microsoft.com/office/developer-program/office-365-developer-program).</span></span> 

