---
title: Excel、Word、PowerPoint のマニフェストにアドイン コマンドを作成する
description: マニフェストに VersionOverrides を使用して、Excel、Word、PowerPoint のアドイン コマンドを定義します。 UI 要素を作成し、ボタンやリストを追加し、操作を実行するために、アドイン コマンドを使用します。
ms.date: 12/04/2017
ms.openlocfilehash: 3c3148a6a6cf3a6b0389b738253acc360f428eb2
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270986"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a><span data-ttu-id="801ea-104">Excel、Word、PowerPoint のマニフェストにアドイン コマンドを作成する</span><span class="sxs-lookup"><span data-stu-id="801ea-104">Create add-in commands in your manifest for Excel, Word, and PowerPoint</span></span>


<span data-ttu-id="801ea-105">マニフェストに **[VersionOverrides](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js)** を使用して、Excel、Word、PowerPoint のアドイン コマンドを定義します。</span><span class="sxs-lookup"><span data-stu-id="801ea-105">Use **[VersionOverrides](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js)** in your manifest to define add-in commands for Excel, Word, and PowerPoint.</span></span> <span data-ttu-id="801ea-106">アドイン コマンドは、アクションを実行する指定された UI 要素を使用して、既定の Office ユーザー インターフェイス (UI) をカスタマイズする簡単な方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="801ea-106">Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions.</span></span> <span data-ttu-id="801ea-107">アドイン コマンドを使用して、以下のことを行えます。</span><span class="sxs-lookup"><span data-stu-id="801ea-107">You can use add-in commands to:</span></span>
- <span data-ttu-id="801ea-108">アドインの機能を簡単に使用できる UI 要素またはエントリ ポイントを作成します。</span><span class="sxs-lookup"><span data-stu-id="801ea-108">Create UI elements or entry points that make your add-in's functionality easier to use.</span></span>  
  
- <span data-ttu-id="801ea-109">ボタン、またはボタンのドロップダウンリストをリボンに追加します。</span><span class="sxs-lookup"><span data-stu-id="801ea-109">Add buttons or a drop-down list of buttons to the ribbon.</span></span>    
  
- <span data-ttu-id="801ea-110">それぞれがオプションのサブメニューを含む個々のメニュー項目を、特定のコンテキスト (ショートカット) メニューに追加します。</span><span class="sxs-lookup"><span data-stu-id="801ea-110">Add individual menu items — each containing optional submenus — to specific context (shortcut) menus.</span></span>    
  
- <span data-ttu-id="801ea-p103">アドイン コマンドが選択されると、操作を実行します。次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p103">Perform actions when your add-in command is chosen. You can:</span></span>
    
  - <span data-ttu-id="801ea-p104">ユーザーが操作する 1 つ以上の作業ウィンドウ アドインを表示します。作業ウィンドウ アドイン内部で、Office の UI ファブリックを使用してカスタム UI を作成する HTML を表示できます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p104">Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.</span></span>
    
     <span data-ttu-id="801ea-115">*または*</span><span class="sxs-lookup"><span data-stu-id="801ea-115">*or*</span></span> 
      
  - <span data-ttu-id="801ea-116">通常はいずれの UI も表示しないで実行する JavaScript コードを実行します。</span><span class="sxs-lookup"><span data-stu-id="801ea-116">Run JavaScript code, which normally runs without displaying any UI.</span></span>
      
<span data-ttu-id="801ea-p105">この記事では、アドイン コマンドを定義するマニフェストの編集方法について説明します。次の図に、アドイン コマンドを定義するのに使用される要素の階層を示します。これらの要素は、この記事で詳細に説明します。</span><span class="sxs-lookup"><span data-stu-id="801ea-p105">This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.</span></span> 
      
<span data-ttu-id="801ea-120">次の画像は、マニフェスト内のアドイン コマンド要素の概要です。</span><span class="sxs-lookup"><span data-stu-id="801ea-120">The following image is an overview of add-in commands elements in the manifest.</span></span> 
<span data-ttu-id="801ea-121">![マニフェスト内のアドイン コマンド要素の概要](../images/version-overrides.png)</span><span class="sxs-lookup"><span data-stu-id="801ea-121">![Overview of add-in commands elements in the manifest](../images/version-overrides.png)</span></span>
 
## <a name="step-1-start-from-a-sample"></a><span data-ttu-id="801ea-122">手順 1: サンプルから始める</span><span class="sxs-lookup"><span data-stu-id="801ea-122">Step 1: Start from a sample</span></span>

<span data-ttu-id="801ea-p107">「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample)」にあるサンプルのいずれかから始めることを強くお勧めします。必要に応じて、このガイドの手順に従って独自のマニフェストを作成できます。「Office-Add-in-Commands-Samples」サイト内で XSD ファイルを使用してご使用のマニフェストを検証できます。アドイン コマンドを使用する前に、「[Excel、Word、および PowerPoint のアドイン コマンド](../design/add-in-commands.md)」をお読みください。</span><span class="sxs-lookup"><span data-stu-id="801ea-p107">We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.</span></span>

## <a name="step-2-create-a-task-pane-add-in"></a><span data-ttu-id="801ea-127">手順 2: 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="801ea-127">Step 2: Create a task pane add-in</span></span>

<span data-ttu-id="801ea-p108">アドイン コマンドの使用を開始するには、まず作業ウィンドウ アドインを作成し、次にアドインのマニフェストをこの記事で説明するように変更する必要があります。コンテンツ アドインではアドイン コマンドを使用できません。既存のマニフェストを更新している場合は、「[手順 3: VersionOverrides 要素を追加する](#step-3-add-versionoverrides-element)」で説明するように、**VersionOverrides** 要素をマニフェストに追加するだけでなく、適切な **XML 名前空間**も追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-p108">To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).</span></span>
   
<span data-ttu-id="801ea-p109">次の例は、Office 2013 アドインのマニフェストを示します。**VersionOverrides** 要素がないため、このマニフェストにはアドイン コマンドがありません。Office 2013 は、アドイン コマンドをサポートしていませんが、このマニフェストに **VersionOverrides** を追加することで、アドインは Office 2013 と Office 2016 の両方で動作します。Office 2013 では、アドインはアドイン コマンドを表示しません。また、**SourceLocation** の値を使用して、アドインを単一の作業ウィンドウ アドインとして実行します。Office 2016 では、**VersionOverrides** 要素が含まれない場合、アドインを実行するために **SourceLocation** が使用されます。ただし、**VersionOverrides** を含める場合は、アドインにアドイン コマンドのみが表示され、アドインは単一の作業ウィンドウ アドインとして表示されません。</span><span class="sxs-lookup"><span data-stu-id="801ea-p109">The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.</span></span>
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
 
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a><span data-ttu-id="801ea-136">手順 3: VersionOverrides 要素を追加する</span><span class="sxs-lookup"><span data-stu-id="801ea-136">Step 3: Add VersionOverrides element</span></span>
<span data-ttu-id="801ea-p110">**VersionOverrides** 要素は、アドイン コマンドの定義を含むルート要素です。**VersionOverrides** はマニフェスト内の **OfficeApp** 要素の子要素です。次の表に、**VersionOverrides** 要素の属性の一覧を示します。</span><span class="sxs-lookup"><span data-stu-id="801ea-p110">The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.</span></span>

|<span data-ttu-id="801ea-140">**属性**</span><span class="sxs-lookup"><span data-stu-id="801ea-140">**Attribute**</span></span>|<span data-ttu-id="801ea-141">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-141">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-142">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="801ea-142">**xmlns**</span></span> <br/> | <span data-ttu-id="801ea-143">必須です。</span><span class="sxs-lookup"><span data-stu-id="801ea-143">Required.</span></span> <span data-ttu-id="801ea-144">スキーマの場所。`http://schemas.microsoft.com/office/taskpaneappversionoverrides` にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-144">The schema location, which must be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span> <br/> |
|<span data-ttu-id="801ea-145">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="801ea-145">**xsi:type**</span></span> <br/> |<span data-ttu-id="801ea-p112">必須。スキーマのバージョン。この記事で説明されているスキーマのバージョンは "VersionOverridesV1_0" です。</span><span class="sxs-lookup"><span data-stu-id="801ea-p112">Required. The schema version. The version described in this article is "VersionOverridesV1_0".</span></span>  <br/> |
   
<span data-ttu-id="801ea-149">次の表は、**VersionOverrides** の子要素です。</span><span class="sxs-lookup"><span data-stu-id="801ea-149">The following table identifies the child elements of **VersionOverrides**.</span></span>
  
|<span data-ttu-id="801ea-150">**要素**</span><span class="sxs-lookup"><span data-stu-id="801ea-150">**Element**</span></span>|<span data-ttu-id="801ea-151">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-151">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-152">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-152">**Description**</span></span> <br/> |<span data-ttu-id="801ea-p113">省略可能。アドインについての説明。この子の **Description** 要素は、マニフェストの親部分の、元の **Description** 要素を上書きします。この **Description** 要素の **resid** 属性は、**String** 要素の **id** に設定されます。**String** 要素には、**Description** のテキストが含まれます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p113">Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. </span></span><br/> |
|<span data-ttu-id="801ea-158">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="801ea-158">**Requirements**</span></span> <br/> |<span data-ttu-id="801ea-p114">省略可能。アドインに必要な最小の Office.js のセットおよびバージョンを指定します。この子の **Requirements** 要素は、マニフェストの親部分の **Requirements** 要素を上書きします。詳細については、「[Office のホストと API の要件を指定する](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。  </span><span class="sxs-lookup"><span data-stu-id="801ea-p114">Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest. For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  </span></span><br/> |
|<span data-ttu-id="801ea-163">**Hosts**</span><span class="sxs-lookup"><span data-stu-id="801ea-163">**Hosts**</span></span> <br/> |<span data-ttu-id="801ea-p115">必須。Office ホストのコレクションを指定します。子の **Hosts** 要素は、マニフェストの親部分の **Hosts** 要素を上書きします。"Workbook" または "Document" に設定された **xsi:type** 属性を含める必要があります。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p115">Required. Specifies a collection of Office hosts. The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". </span></span><br/> |
|<span data-ttu-id="801ea-168">**Resources**</span><span class="sxs-lookup"><span data-stu-id="801ea-168">**Resources**</span></span> <br/> |<span data-ttu-id="801ea-p116">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。たとえば、**Description** 要素の値は、**Resources** の子要素を参照します。**Resources** 要素については、この記事の「[手順 7: Resources 要素を追加する](#step-7-add-the-resources-element)」で説明します。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p116">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. </span></span><br/> |
   
<span data-ttu-id="801ea-172">次の例に、**VersionOverrides** 要素と子要素を使用する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="801ea-172">The following example shows how to use the **VersionOverrides** element and its child elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a><span data-ttu-id="801ea-173">手順 4: Hosts、Host、DesktopFormFactor 要素を追加する</span><span class="sxs-lookup"><span data-stu-id="801ea-173">Step 4: Add Hosts, Host, and DesktopFormFactor elements</span></span>

<span data-ttu-id="801ea-p117">**Hosts** 要素には、1 つ以上の **Host** 要素が含まれます。**Host** 要素は、特定の Office ホストを指定します。**Host** 要素には、アドインが Office ホストにインストールされた後で表示するアドイン コマンドを指定する子要素が含まれます。同じアドイン コマンドを複数の異なる Office ホストで表示する場合は、各 **Host** で子要素を重複させる必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-p117">The **Hosts** element contains one or more **Host** elements. A **Host** element specifies a particular Office host. The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host. To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.</span></span>
       
<span data-ttu-id="801ea-178">**DesktopFormFactor** 要素では、Windows デスクトップ上の Office、および Office Online (ブラウザー内) で実行するアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="801ea-178">The **DesktopFormFactor** element specifies the settings for an add-in that runs in Office on Windows desktop, and Office Online (in browser).</span></span>
      
<span data-ttu-id="801ea-179">**Hosts** 要素、**Host** 要素、**DesktopFormFactor** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="801ea-179">The following is an example of **Hosts**, **Host**, and **DesktopFormFactor** elements.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a><span data-ttu-id="801ea-180">手順 5: FunctionFile 要素を追加する</span><span class="sxs-lookup"><span data-stu-id="801ea-180">Step 5: Add the FunctionFile element</span></span>

<span data-ttu-id="801ea-p118">**FunctionFile** 要素では、アドイン コマンドが **ExecuteFunction** 操作を使用するときに実行される JavaScript コードを含むファイルを指定します (「[ボタン コントロール](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control)」の説明を参照)。**FunctionFile** 要素の **resid** 属性は、アドイン コマンドに必要なすべての JavaScript ファイルを含む HTML ファイルに設定されます。JavaScript ファイルに直接リンクすることはできません。HTML ファイルにのみリンクできます。ファイル名は、**Resources** 要素の **Url** 要素として指定されます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p118">The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.</span></span>
        
<span data-ttu-id="801ea-186">**FunctionFile** 要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="801ea-186">The following is an example of the **FunctionFile** element.</span></span>
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint> 

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> <span data-ttu-id="801ea-187">JavaScript コードが `Office.initialize` を呼び出していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="801ea-187">Make sure your JavaScript code calls  `Office.initialize`.</span></span> 
   
<span data-ttu-id="801ea-p119">**FunctionFile** 要素によって参照される HTML ファイルの JavaScript は、`Office.initialize` を呼び出す必要があります。**FunctionName** 要素 (「[ボタン コントロール](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control)」の説明を参照) は、**FunctionFile** の関数を使用します。</span><span class="sxs-lookup"><span data-stu-id="801ea-p119">The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control) for a description) uses the functions in **FunctionFile**.</span></span>
     
<span data-ttu-id="801ea-190">次のコードは、**FunctionName** で使用される関数の実装方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="801ea-190">The following code shows how to implement the function used by **FunctionName**.</span></span>

```javascript

<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here. 
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message. 
                }
                else {
                    // Show success message.
                }
            });
        
        // Calling event.completed is required. event.completed lets the platform know that processing has completed. 
        event.completed();
    }
</script>
```

> [!IMPORTANT]
> <span data-ttu-id="801ea-p120">**event.completed** に対する呼び出しにより、イベントが正常に処理されたことが通知されます。同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。最初のイベントが自動的に実行され、その他のイベントはキューに残ります。関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。**event.completed** を実装する必要があります。実装しない場合、関数は実行されません。</span><span class="sxs-lookup"><span data-stu-id="801ea-p120">The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.</span></span>
 
## <a name="step-6-add-extensionpoint-elements"></a><span data-ttu-id="801ea-196">手順 6: ExtensionPoint 要素を追加する</span><span class="sxs-lookup"><span data-stu-id="801ea-196">Step 6: Add ExtensionPoint elements</span></span>

<span data-ttu-id="801ea-p121">**ExtensionPoint** 要素は、Office UI のどこにアドイン コマンドを表示するかを定義します。以下の **xsi:type** 値を使用して、**ExtensionPoint** 要素を定義できます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p121">The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:</span></span>
   
- <span data-ttu-id="801ea-199">**PrimaryCommandSurface**。Office のリボンを参照します。</span><span class="sxs-lookup"><span data-stu-id="801ea-199">**PrimaryCommandSurface**, which refers to the ribbon in Office.</span></span>
     
- <span data-ttu-id="801ea-200">**ContextMenu**。Office UI で右クリックしたときに表示されるショートカット メニューです。</span><span class="sxs-lookup"><span data-stu-id="801ea-200">**ContextMenu**, which is the shortcut menu that appears when you right-click in the Office UI.</span></span>
    
<span data-ttu-id="801ea-201">次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="801ea-201">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>
    
> [!IMPORTANT]
> <span data-ttu-id="801ea-p122">ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。`<CustomTab id="mycompanyname.mygroupname">`</span><span class="sxs-lookup"><span data-stu-id="801ea-p122">For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.</span></span> 
  
```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso Tab">
  <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
  <!-- <OfficeTab id="TabData"> -->
    <Label resid="residLabel4" />
    <Group id="Group1Id12">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Button1Id1">
        
        <!-- information about the control -->
      </Control>   
      <!-- other controls, as needed -->                                    
    </Group>
  </CustomTab>
</ExtensionPoint>
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
    </Control>   
    <!-- other controls, as needed -->         
  </OfficeMenu>
</ExtensionPoint>
```

|<span data-ttu-id="801ea-205">**要素**</span><span class="sxs-lookup"><span data-stu-id="801ea-205">**Element**</span></span>|<span data-ttu-id="801ea-206">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-206">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-207">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="801ea-207">**CustomTab**</span></span> <br/> |<span data-ttu-id="801ea-p123">カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p123">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. </span></span><br/> |
|<span data-ttu-id="801ea-211">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="801ea-211">**OfficeTab**</span></span> <br/> |<span data-ttu-id="801ea-p124">既定の Office リボン タブを拡張する場合は必須 (**PrimaryCommandSurface** を使用)。**OfficeTab** 要素を使用する場合、**CustomTab** 要素は使用できません。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p124">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the **OfficeTab** element, you can't use the **CustomTab** element. </span></span><br/> <span data-ttu-id="801ea-214">**id** 属性と共に使用するその他のタブの値については、「[既定の Office リボン タブ](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officetab?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="801ea-214">For more tab values to use with the **id** attribute, see [Tab values for default Office ribbon tabs](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officetab?view=office-js).</span></span>  <br/> |
|<span data-ttu-id="801ea-215">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="801ea-215">**OfficeMenu**</span></span> <br/> | <span data-ttu-id="801ea-p125">既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p125">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="801ea-p126">Excel または Word の場合は **ContextMenuText**。ユーザーがテキストを選択し、選択したテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p126">**ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="801ea-p127">Excel の場合は **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p127">**ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. </span></span><br/> |
|<span data-ttu-id="801ea-222">**グループ**</span><span class="sxs-lookup"><span data-stu-id="801ea-222">**Group**</span></span> <br/> |<span data-ttu-id="801ea-p128">タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p128">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="801ea-226">**Label**</span><span class="sxs-lookup"><span data-stu-id="801ea-226">**Label**</span></span> <br/> |<span data-ttu-id="801ea-p129">必須。グループのラベル。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p129">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-231">**Icon**</span><span class="sxs-lookup"><span data-stu-id="801ea-231">**Icon**</span></span> <br/> |<span data-ttu-id="801ea-p130">必須。小さいフォーム ファクターのデバイス、または多くのボタンが表示されるときに使用されるグループのアイコンを指定します。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で指定します。次の 3 つのイメージのサイズが必要です。16、32、および 80。次の 5 つのオプションのサイズもサポートされています。20、24、40、48、および 64。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p130">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="801ea-239">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="801ea-239">**Tooltip**</span></span> <br/> |<span data-ttu-id="801ea-p131">省略可能。グループのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p131">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-244">**Control**</span><span class="sxs-lookup"><span data-stu-id="801ea-244">**Control**</span></span> <br/> |<span data-ttu-id="801ea-p132">各グループには、1 つ以上のコントロールが必要です。**Control** 要素は、**Button** または **Menu** のいずれかにすることができます。ボタンのコントロールのドロップダウン リストを指定するには、**Menu** を使用します。現在、ボタンとメニューのみがサポートされています。詳しくは、「[ボタン コントロール](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control)」および「[メニュー コントロール](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#menu-dropdown-button-controls)」のセクションをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="801ea-p132">Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control) and [Menu controls](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#menu-dropdown-button-controls) sections for more information. </span></span><br/><span data-ttu-id="801ea-250">**注:** トラブルシューティングを容易にするために、**Control** 要素と関連する **Resources** 子要素を 1 つずつ追加することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="801ea-250">**Note:** To make troubleshooting easier, we recommend that you add a **Control** element and the related **Resources** child elements one at a time.</span></span>          |
   

### <a name="button-controls"></a><span data-ttu-id="801ea-251">Button コントロール</span><span class="sxs-lookup"><span data-stu-id="801ea-251">Button controls</span></span>
<span data-ttu-id="801ea-p133">ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。JavaScript 関数を実行するか、作業ウィンドウを表示することができます。次の例は、2 つのボタンを定義する方法を示しています。最初のボタンは UI を表示せずに JavaScript 関数を実行し、2 つ目のボタンは作業ウィンドウを表示します。**Control** 要素では、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="801ea-p133">A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:</span></span>        

- <span data-ttu-id="801ea-257">**type** 属性は必須であり、**Button** に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-257">The **type** attribute is required, and must be set to **Button**.</span></span>
    
- <span data-ttu-id="801ea-258">**Control** 要素の **id** 属性は、最大 125 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="801ea-258">The **id** attribute of the **Control** element is a string with a maximum of 125 characters.</span></span>
    
```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|<span data-ttu-id="801ea-259">**要素**</span><span class="sxs-lookup"><span data-stu-id="801ea-259">**Elements**</span></span>|<span data-ttu-id="801ea-260">**Description**</span><span class="sxs-lookup"><span data-stu-id="801ea-260">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-261">**Label**</span><span class="sxs-lookup"><span data-stu-id="801ea-261">**Label**</span></span> <br/> |<span data-ttu-id="801ea-p134">必須。ボタンのテキスト。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p134">Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-266">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="801ea-266">**Tooltip**</span></span> <br/> |<span data-ttu-id="801ea-p135">省略可能。ボタンのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p135">Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-271">**Supertip**</span><span class="sxs-lookup"><span data-stu-id="801ea-271">**Supertip**</span></span> <br/> | <span data-ttu-id="801ea-p136">必須。このボタンのヒントであり、次のものによって定義されます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p136">Required. The supertip for this button, which is defined by the following: </span></span><br/> <span data-ttu-id="801ea-274">**Title**</span><span class="sxs-lookup"><span data-stu-id="801ea-274">**Title**</span></span> <br/>  <span data-ttu-id="801ea-p137">必須。ヒントのテキスト。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p137">Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="801ea-279">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-279">**Description**</span></span> <br/>  <span data-ttu-id="801ea-p138">必須。ヒントの説明。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p138">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-284">**Icon**</span><span class="sxs-lookup"><span data-stu-id="801ea-284">**Icon**</span></span> <br/> | <span data-ttu-id="801ea-p139">必須。ボタンの **Image** 要素を含みます。画像ファイルは必ず .png 形式です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p139">Required. Contains the **Image** elements for the button. Image files must be .png format. </span></span><br/> <span data-ttu-id="801ea-288">**Image**</span><span class="sxs-lookup"><span data-stu-id="801ea-288">**Image**</span></span> <br/>  <span data-ttu-id="801ea-p140">ボタンに表示する画像を定義します。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で示します。次の 3 つのイメージのサイズが必要です。16、32、および 80。次の 5 つのオプションのサイズもサポートされています。20、24、40、48、および 64。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p140">Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="801ea-295">**操作**</span><span class="sxs-lookup"><span data-stu-id="801ea-295">**Action**</span></span> <br/> | <span data-ttu-id="801ea-p141">必須。ユーザーがボタンを選択したときに実行する操作を指定します。**xsi:type** 属性の値は、次のいずれかを指定できます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p141">Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: </span></span><br/> <span data-ttu-id="801ea-p142">**ExecuteFunction**。**FunctionFile** によって参照されるファイルにある JavaScript 関数を実行します。**ExecuteFunction** は UI を表示しません。**FunctionName** 子要素は、実行する関数の名前を指定します。</span><span class="sxs-lookup"><span data-stu-id="801ea-p142">**ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. </span></span><br/> <span data-ttu-id="801ea-p143">**ShowTaskPane**。作業ウィンドウ アドインを表示します。**SourceLocation** 子要素は、表示する作業ウィンドウ アドインのソース ファイルの位置を指定します。**resid** 属性は、**Url** 要素の **id** 属性の値に設定します。この要素は、**Resources** 要素の **Urls** 要素に含まれています。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p143">**ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. </span></span><br/> |
   

### <a name="menu-controls"></a><span data-ttu-id="801ea-305">Menu コントロール</span><span class="sxs-lookup"><span data-stu-id="801ea-305">Menu controls</span></span>
<span data-ttu-id="801ea-306">**Menu** コントロールは、**PrimaryCommandSurface** または **ContextMenu** のどちらかで使用できます。また、以下の項目を定義します。</span><span class="sxs-lookup"><span data-stu-id="801ea-306">A **Menu** control can be used with either **PrimaryCommandSurface** or **ContextMenu**, and defines:</span></span>
  
- <span data-ttu-id="801ea-307">ルートレベルのメニュー項目。</span><span class="sxs-lookup"><span data-stu-id="801ea-307">A root-level menu item.</span></span>
   
- <span data-ttu-id="801ea-308">サブメニュー項目のリスト。</span><span class="sxs-lookup"><span data-stu-id="801ea-308">A list of submenu items.</span></span>
 
<span data-ttu-id="801ea-p144">**PrimaryCommandSurface** と共に使用すると、ルートのメニュー項目がリボンのボタンとして表示されます。ボタンを選択すると、サブメニューがドロップダウン リストとして表示されます。**ContextMenu** と共に使用すると、サブメニューのあるメニュー項目がコンテキスト メニューに挿入されます。どちらの場合も、各サブメニュー項目は JavaScript 関数を実行するか、作業ウィンドウを表示することができます。現時点では、サブメニューの 1 つのレベルのみがサポートされます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p144">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>
       
<span data-ttu-id="801ea-p145">次の例では、2 つのサブメニュー項目があるメニュー項目を定義する方法を示します。最初のサブメニュー項目は作業ウィンドウを表示し、2 つ目のサブメニュー項目は、JavaScript 関数を実行します。**Control** 要素では、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="801ea-p145">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:</span></span>
    
- <span data-ttu-id="801ea-317">**xsi:type** 属性は必須であり、**Menu** に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-317">The **xsi:type** attribute is required, and must be set to **Menu**.</span></span>
  
- <span data-ttu-id="801ea-318">**id** 属性は、最大 125 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="801ea-318">The **id** attribute is a string with a maximum of 125 characters.</span></span>
    
```xml

<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|<span data-ttu-id="801ea-319">**要素**</span><span class="sxs-lookup"><span data-stu-id="801ea-319">**Elements**</span></span>|<span data-ttu-id="801ea-320">**Description**</span><span class="sxs-lookup"><span data-stu-id="801ea-320">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-321">**Label**</span><span class="sxs-lookup"><span data-stu-id="801ea-321">**Label**</span></span> <br/> |<span data-ttu-id="801ea-p146">必須。ルートのメニュー項目のテキスト。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p146">Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-326">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="801ea-326">**Tooltip**</span></span> <br/> |<span data-ttu-id="801ea-p147">省略可能。メニューのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p147">Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-331">**SuperTip**</span><span class="sxs-lookup"><span data-stu-id="801ea-331">**SuperTip**</span></span> <br/> | <span data-ttu-id="801ea-p148">必須。メニューのヒントであり、次のものによって定義されます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p148">Required. The supertip for the menu, which is defined by the following: </span></span><br/> <span data-ttu-id="801ea-334">**Title**</span><span class="sxs-lookup"><span data-stu-id="801ea-334">**Title**</span></span> <br/>  <span data-ttu-id="801ea-p149">必須。ヒントのテキスト。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p149">Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. </span></span><br/> <span data-ttu-id="801ea-339">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-339">**Description**</span></span> <br/>  <span data-ttu-id="801ea-p150">必須。ヒントの説明。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p150">Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. </span></span><br/> |
|<span data-ttu-id="801ea-344">**Icon**</span><span class="sxs-lookup"><span data-stu-id="801ea-344">**Icon**</span></span> <br/> | <span data-ttu-id="801ea-p151">必須。メニューの **Image** 要素を含みます。画像ファイルは必ず .png 形式です。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p151">Required. Contains the **Image** elements for the menu. Image files must be .png format. </span></span><br/> <span data-ttu-id="801ea-348">**Image**</span><span class="sxs-lookup"><span data-stu-id="801ea-348">**Image**</span></span> <br/>  <span data-ttu-id="801ea-p152">メニューの画像。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で示します。次の 3 つのイメージのサイズ (ピクセル単位) が必要です。16、32、および 80。次の 5 つのオプションのサイズ (ピクセル単位) もサポートされています。20、24、40、48、および 64。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p152">An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. </span></span><br/> |
|<span data-ttu-id="801ea-355">**Items**</span><span class="sxs-lookup"><span data-stu-id="801ea-355">**Items**</span></span> <br/> |<span data-ttu-id="801ea-p153">必須。各サブメニュー項目の **Item** 要素を含みます。各 **Item** 要素は、[ボタン コントロール](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control)と同じ子要素を含みます。  </span><span class="sxs-lookup"><span data-stu-id="801ea-p153">Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control?view=office-js#Button-control).  </span></span><br/> |
   
## <a name="step-7-add-the-resources-element"></a><span data-ttu-id="801ea-359">手順 7: Resources 要素を追加する</span><span class="sxs-lookup"><span data-stu-id="801ea-359">Step 7: Add the Resources element</span></span>

<span data-ttu-id="801ea-p154">**Resources** 要素は、**VersionOverrides** 要素の異なる子要素で使用されるリソースを含みます。リソースには、アイコン、文字列、および URL が含まれます。マニフェスト内の要素は、リソースの **id** を参照することでリソースを使用できます。**id** を使用するマニフェストの編成に有用です。特に、異なるロケールのリソースの異なるバージョンがある場合に役立ちます。**id** は 最大 32 文字まで使用できます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p154">The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.</span></span>
  
    
    
<span data-ttu-id="801ea-p155">次の表に、**Resources** 要素の使用法の例を示します。各リソースは、特定のロケールに異なるリソースを定義する 1 つ以上の **Override** 子要素を持つことができます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p155">The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>


```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>        
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>        
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>      
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|<span data-ttu-id="801ea-367">**Resource**</span><span class="sxs-lookup"><span data-stu-id="801ea-367">**Resource**</span></span>|<span data-ttu-id="801ea-368">**説明**</span><span class="sxs-lookup"><span data-stu-id="801ea-368">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-369">**Images**/ **Image**</span><span class="sxs-lookup"><span data-stu-id="801ea-369">**Images**/ **Image**</span></span> <br/> | <span data-ttu-id="801ea-p156">イメージ ファイルへの HTTPS URL を指定します。各イメージは、次の 3 つの必須のイメージ サイズを定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-p156">Provides the HTTPS URL to an image file. Each image must define the three required image sizes:</span></span> <br/>  <span data-ttu-id="801ea-372">16×16</span><span class="sxs-lookup"><span data-stu-id="801ea-372">16×16</span></span> <br/>  <span data-ttu-id="801ea-373">32×32</span><span class="sxs-lookup"><span data-stu-id="801ea-373">32×32</span></span> <br/>  <span data-ttu-id="801ea-374">80×80</span><span class="sxs-lookup"><span data-stu-id="801ea-374">80×80</span></span> <br/>  <span data-ttu-id="801ea-375">次のイメージ サイズもサポートされますが、必須ではありません。</span><span class="sxs-lookup"><span data-stu-id="801ea-375">The following image sizes are also supported, but not required:</span></span> <br/>  <span data-ttu-id="801ea-376">20×20</span><span class="sxs-lookup"><span data-stu-id="801ea-376">20×20</span></span> <br/>  <span data-ttu-id="801ea-377">24×24</span><span class="sxs-lookup"><span data-stu-id="801ea-377">24×24</span></span> <br/>  <span data-ttu-id="801ea-378">40×40</span><span class="sxs-lookup"><span data-stu-id="801ea-378">40×40</span></span> <br/>  <span data-ttu-id="801ea-379">48×48</span><span class="sxs-lookup"><span data-stu-id="801ea-379">48×48</span></span> <br/>  <span data-ttu-id="801ea-380">64×64</span><span class="sxs-lookup"><span data-stu-id="801ea-380">64×64</span></span> <br/> |
|<span data-ttu-id="801ea-381">**Urls**/ **Url**</span><span class="sxs-lookup"><span data-stu-id="801ea-381">**Urls**/ **Url**</span></span> <br/> |<span data-ttu-id="801ea-p157">HTTPS URL の場所を指定します。URL には最大 2048 文字まで指定できます。</span><span class="sxs-lookup"><span data-stu-id="801ea-p157">Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.</span></span>  <br/> |
|<span data-ttu-id="801ea-384">**ShortStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="801ea-384">**ShortStrings**/ **String**</span></span> <br/> |<span data-ttu-id="801ea-p158">**Label** 要素と **Title** 要素のテキスト。各 **String** には、最大 125 文字を使用できます。 </span><span class="sxs-lookup"><span data-stu-id="801ea-p158">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. </span></span><br/> |
|<span data-ttu-id="801ea-387">**LongStrings**/ **String**</span><span class="sxs-lookup"><span data-stu-id="801ea-387">**LongStrings**/ **String**</span></span> <br/> |<span data-ttu-id="801ea-p159">**Tooltip** と **Description** 要素のテキスト。各 **String** は最大 250 文字です。</span><span class="sxs-lookup"><span data-stu-id="801ea-p159">The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. </span></span><br/> |
   
> [!NOTE] 
> <span data-ttu-id="801ea-390">**Image** 要素と **Url** 要素のすべての URL で Secure Sockets Layer (SSL) を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="801ea-390">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="tab-values-for-default-office-ribbon-tabs"></a><span data-ttu-id="801ea-391">既定の Office リボン タブの値</span><span class="sxs-lookup"><span data-stu-id="801ea-391">Tab values for default Office ribbon tabs</span></span>
<span data-ttu-id="801ea-p160">Excel および Word で、既定の Office UI タブを使用することで、リボンにアドイン コマンドを追加できます。次の表に、**OfficeTab** 要素の **id** 属性で使用できる値を示します。タブの値は大文字と小文字を区別します。</span><span class="sxs-lookup"><span data-stu-id="801ea-p160">In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.</span></span>

|<span data-ttu-id="801ea-395">**Office ホスト アプリケーション**</span><span class="sxs-lookup"><span data-stu-id="801ea-395">**Office host application**</span></span>|<span data-ttu-id="801ea-396">**タブの値**</span><span class="sxs-lookup"><span data-stu-id="801ea-396">**Tab values**</span></span>|
|:-----|:-----|
|<span data-ttu-id="801ea-397">Excel</span><span class="sxs-lookup"><span data-stu-id="801ea-397">Excel</span></span>  <br/> |<span data-ttu-id="801ea-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span><span class="sxs-lookup"><span data-stu-id="801ea-398">**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval**</span></span> <br/> |
|<span data-ttu-id="801ea-399">Word</span><span class="sxs-lookup"><span data-stu-id="801ea-399">Word</span></span>  <br/> |<span data-ttu-id="801ea-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span><span class="sxs-lookup"><span data-stu-id="801ea-400">**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation**</span></span> <br/> |
|<span data-ttu-id="801ea-401">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="801ea-401">PowerPoint</span></span>  <br/> |<span data-ttu-id="801ea-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span><span class="sxs-lookup"><span data-stu-id="801ea-402">**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**</span></span>          <br/> |
   
## <a name="see-also"></a><span data-ttu-id="801ea-403">関連項目</span><span class="sxs-lookup"><span data-stu-id="801ea-403">See also</span></span>

-  [<span data-ttu-id="801ea-404">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="801ea-404">Add-in commands for Excel, Word and PowerPoint</span></span>](../design/add-in-commands.md)      
