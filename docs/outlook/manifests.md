---
title: Outlook アドイン マニフェスト
description: このマニフェストでは、 Outlook アドインが Outlook クラインアント間でどのように統合されるかを、例を交えて説明します。
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: 0135db8b6ff2b9fbcb3b6370979d8013aa21155a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839825"
---
# <a name="outlook-add-in-manifests"></a><span data-ttu-id="8f1cd-103">Outlook アドイン マニフェスト</span><span class="sxs-lookup"><span data-stu-id="8f1cd-103">Outlook add-in manifests</span></span>

<span data-ttu-id="8f1cd-p101">Outlook アドインは XML アドイン マニフェストと Web ページの 2 つのコンポーネントで構成されています。これらは Office アドイン (office.js) の JavaScript ライブラリでサポートされます。マニフェストは、アドインが Outlook クライアント間でどのように統合されるかを説明します。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p101">An Outlook add-in consists of two components: the XML add-in manifest and a web page supported by the JavaScript library for Office Add-ins (office.js). The manifest describes how the add-in integrates across Outlook clients. The following is an example.</span></span>

 > [!NOTE]
 > <span data-ttu-id="8f1cd-p102">次のサンプルの URL 値はすべて "https://appdemo.contoso.com" で始まります。この値はプレースホルダーであり、有効な実際のマニフェストでは、この部分には有効な HTTPS Web URL が入ります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p102">All URL values in the following sample begin with "https://appdemo.contoso.com". This value is a placeholder. In an actual valid manifest, these values would contain valid https web URLs.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="https://appdemo.contoso.com/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="https://appdemo.contoso.com/images/blue-128.png" />
  <SupportUrl DefaultValue="https://appdemo.contoso.com"/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read task pane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />

          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="https://appdemo.contoso.com/images/blue-16.png" />
        <bt:Image id="blue-icon-32" DefaultValue="https://appdemo.contoso.com/images/blue-32.png" />
        <bt:Image id="blue-icon-80" DefaultValue="https://appdemo.contoso.com/images/blue-80.png" />
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="https://appdemo.contoso.com/images/red-16.png" />
        <bt:Image id="red-icon-32" DefaultValue="https://appdemo.contoso.com/images/red-32.png" />
        <bt:Image id="red-icon-80" DefaultValue="https://appdemo.contoso.com/images/red-80.png" />
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="https://appdemo.contoso.com/images/green-16.png" />
        <bt:Image id="green-icon-32" DefaultValue="https://appdemo.contoso.com/images/green-32.png" />
        <bt:Image id="green-icon-80" DefaultValue="https://appdemo.contoso.com/images/green-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://appdemo.contoso.com/FunctionFile/Functions.html" />
        <bt:Url id="readTaskPaneUrl" DefaultValue="https://appdemo.contoso.com/AppRead/TaskPane/TaskPane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo" />
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject" />
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property" />
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties" />

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment" />
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get" />
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties" />

        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class" />
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created" />
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button." />
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button." />
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane." />

        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

## <a name="schema-versions"></a><span data-ttu-id="8f1cd-110">スキーマのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f1cd-110">Schema versions</span></span>

<span data-ttu-id="8f1cd-p103">すべての Outlook クライアントで最新の機能がサポートされているわけではありません。一部の Outlook ユーザーは前のバージョンの Outlook を使用していることがあります。スキーマのバージョンにより、開発者は下位互換性のあるアドインを作成することができます。その際、使用可能な場合には最新の機能を使用しますが、前のバージョンも引き続き機能します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p103">Not all Outlook clients support the latest features, and some Outlook users will have an older version of Outlook. Having schema versions lets developers build add-ins that are backwards compatible, using the newest features where they are available but still functioning on older versions.</span></span>

<span data-ttu-id="8f1cd-p104">マニフェストの **VersionOverrides** 要素が、この一例です。**VersionOverrides** 内で定義されたすべての要素は、マニフェストの他の部分にある同じ要素をオーバーライドします。つまり、Outlook は、可能な場合は常に、**VersionOverrides** セクションにあるものを使用してアドインをセットアップします。ただし、Outlook のバージョンが特定のバージョンの **VersionOverrides** をサポートしていない場合、Outlook はこれを無視して、マニフェストの残りの部分の情報のみを使用します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p104">The **VersionOverrides** element in the manifest is an example of this. All elements defined inside **VersionOverrides** will override the same element in the other part of the manifest. This means that, whenever possible, Outlook will use what is in the **VersionOverrides** section to set up the add-in. However, if the version of Outlook doesn't support a certain version of **VersionOverrides**, Outlook will ignore it and depend on the information in the rest of the manifest.</span></span> 

<span data-ttu-id="8f1cd-117">このアプローチでは、開発者は個別のマニフェストを複数作成する必要がなく、すべてを 1 つのファイルで定義することになります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-117">This approach means that developers don't have to create multiple individual manifests, but rather keep everything defined in one file.</span></span>

<span data-ttu-id="8f1cd-118">現在のスキーマのバージョンは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-118">The current versions of the schema are:</span></span>


|<span data-ttu-id="8f1cd-119">バージョン</span><span class="sxs-lookup"><span data-stu-id="8f1cd-119">Version</span></span>|<span data-ttu-id="8f1cd-120">説明</span><span class="sxs-lookup"><span data-stu-id="8f1cd-120">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="8f1cd-121">v1.0</span><span class="sxs-lookup"><span data-stu-id="8f1cd-121">v1.0</span></span>|<span data-ttu-id="8f1cd-p105">Office JavaScript API バージョン 1.0 をサポートします。Outlook アドインであれば、閲覧フォームがサポートされることになります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p105">Supports version 1.0 of the Office JavaScript API. For Outlook add-ins, this supports read form.</span></span> |
|<span data-ttu-id="8f1cd-124">v1.1</span><span class="sxs-lookup"><span data-stu-id="8f1cd-124">v1.1</span></span>|<span data-ttu-id="8f1cd-p106">Office JavaScript API バージョン 1.1 と **VersionOverrides** をサポートします。Outlook アドインで、新規作成フォームもサポートされることになります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p106">Supports version 1.1 of the Office JavaScript API and **VersionOverrides**. For Outlook add-ins, this adds support for compose form.</span></span>|
|<span data-ttu-id="8f1cd-127">**VersionOverrides** 1.0</span><span class="sxs-lookup"><span data-stu-id="8f1cd-127">**VersionOverrides** 1.0</span></span>|<span data-ttu-id="8f1cd-p107">Office JavaScript API の最新バージョンをサポートします。これは、アドイン コマンドをサポートします。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p107">Supports later versions of the Office JavaScript API. This supports add-in commands.</span></span>|
|<span data-ttu-id="8f1cd-130">**VersionOverrides** 1.1</span><span class="sxs-lookup"><span data-stu-id="8f1cd-130">**VersionOverrides** 1.1</span></span>|<span data-ttu-id="8f1cd-p108">Office JavaScript API の最新バージョンをサポートします。これは、アドイン コマンドをサポートし、[ピン留め可能な作業ウィンドウ](pinnable-taskpane.md)やモバイル アドインなどの、より新しい機能のサポートを追加します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p108">Supports later versions of the Office JavaScript API. This supports add-in commands and adds support for newer features, such as [pinnable task panes](pinnable-taskpane.md) and mobile add-ins.</span></span>|

<span data-ttu-id="8f1cd-p109">この記事では、v1.1 マニフェストの要件を取り上げます。アドイン マニフェストで **VersionOverrides** 要素を使用するとしても、**VersionOverrides** をサポートしない以前のクライアントでアドインが機能できるように 1.1 マニフェスト要素を組み込むことは重要です。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p109">This article will cover the requirements for a v1.1 manifest. Even if your add-in manifest uses the **VersionOverrides** element, it is still important to include the v1.1 manifest elements to allow your add-in to work with older clients that do not support **VersionOverrides**.</span></span>

> [!NOTE]
> <span data-ttu-id="8f1cd-p110">Outlook では、マニフェストの検証にスキーマを使用します。スキーマは、マニフェスト内の要素が特定の順序に従うことを要求します。要求されている順序に従わない要素が含まれていると、アドインをサイドロードするときにエラーが発生することがあります。[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) をダウンロードすると、要求されている順序に要素を並べてマニフェストを作成するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p110">Outlook uses a schema to validate manifests. The schema requires that elements in the manifest appear in a specific order. If you include elements out of the required order, you may get errors when sideloading your add-in. You can download the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) to help create your manifest with elements in the required order.</span></span>

## <a name="root-element"></a><span data-ttu-id="8f1cd-139">ルート要素</span><span class="sxs-lookup"><span data-stu-id="8f1cd-139">Root element</span></span>

<span data-ttu-id="8f1cd-p111">Outlook アドイン マニフェストのルート要素は **OfficeApp** です。この要素はまた、既定の名前空間、スキーマのバージョン、およびアドインの種類を宣言します。開始タグと終了タグの間にマニフェストのその他すべての要素を配置します。ルート要素の例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p111">The root element for the Outlook add-in manifest is **OfficeApp**. This element also declares the default namespace, schema version and the type of add-in. Place all other elements in the manifest within its open and close tags. The following is an example of the root element:</span></span>


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest -->

</OfficeApp>
```

## <a name="version"></a><span data-ttu-id="8f1cd-144">バージョン</span><span class="sxs-lookup"><span data-stu-id="8f1cd-144">Version</span></span>

<span data-ttu-id="8f1cd-p112">これは、特定のアドインのバージョンです。開発者がマニフェストの一部を更新する場合、バージョンの番号も増やす必要があります。このようにすることで、新しいマニフェストがインストールされると、既存のマニフェストが上書きされ、ユーザーは新機能を使用できるようになります。このアドインがストアに送信されている場合は、新しいマニフェストをもう一度送信して検証する必要があります。承認されると、数時間以内に、このアドインのユーザーは更新された新しいマニフェストを自動的に使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p112">This is the version of the specific add-in. If a developer updates something in the manifest, the version must be incremented as well. This way, when the new manifest is installed, it will overwrite the existing one and the user will get the new functionality. If this add-in was submitted to the store, the new manifest will have to be re-submitted and re-validated. Then, users of this add-in will get the new updated manifest automatically in a few hours, after it is approved.</span></span>

<span data-ttu-id="8f1cd-p113">アドインに必要なアクセス許可が変更された場合、ユーザーは、アップグレードを行いアドインに再同意するように求められます。管理者が組織全体にこのアドインをインストール済みである場合、管理者がまず再同意する必要があります。それまでの間、ユーザーには引き続き古い機能が表示されます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p113">If the add-in's requested permissions change, users will be prompted to upgrade and re-consent to the add-in. If the admin installed this add-in for the entire organization, the admin will have to re-consent first. Users will continue to see old functionality in the meantime.</span></span>

## <a name="versionoverrides"></a><span data-ttu-id="8f1cd-153">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8f1cd-153">VersionOverrides</span></span>

<span data-ttu-id="8f1cd-154">**VersionOverrides** 要素は、[アドイン コマンド](add-in-commands-for-outlook.md)の情報の場所です。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-154">The **VersionOverrides** element is the location of information for [add-in commands](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="8f1cd-155">この要素は、アドインによって[モバイル アドイン](add-mobile-support.md)のサポートが定義される場所でもあります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-155">This element is also where add-ins define support for [mobile add-ins](add-mobile-support.md).</span></span>

<span data-ttu-id="8f1cd-156">この要素の説明については、「[Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する](../develop/create-addin-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-156">For a discussion on this element, see [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="localization"></a><span data-ttu-id="8f1cd-157">ローカライズ</span><span class="sxs-lookup"><span data-stu-id="8f1cd-157">Localization</span></span>

<span data-ttu-id="8f1cd-p114">名前、説明、および読み込む URL など、アドインのいくつかの側面は、各種のロケール用にローカライズする必要があります。これらの要素は、既定値を指定してから、**VersionOverrides** 要素内の **Resources** 要素でロケールのオーバーライドを指定することによって簡単にローカライズできます。画像、URL、および文字列をオーバーライドする方法を次に示します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p114">Some aspects of the add-in need to be localized for different locales, such as the name, description and the URL that's loaded. These elements can easily be localized by specifying the default value and then locale overrides in the **Resources** element within the **VersionOverrides** element. The following shows how to override an image, a URL, and a string:</span></span>


```XML
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
      <!-- add information for other locales -->
    </bt:Image>
  </bt:Images>

  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
      <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
      <!-- add information for other locales -->
    </bt:Url>
  </bt:Urls>

  <bt:ShortStrings> 
    <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
      <bt:Override Locale="ar-sa" Value="<add localized value here>" />
      <!-- add information for other locales -->
    </bt:String>
  </bt:ShortStrings>
</Resources>
```

<span data-ttu-id="8f1cd-161">スキーマ リファレンスには、ローカライズできる要素に関する詳しい情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-161">The schema reference contains full information on which elements can be localized.</span></span>

## <a name="hosts"></a><span data-ttu-id="8f1cd-162">Hosts</span><span class="sxs-lookup"><span data-stu-id="8f1cd-162">Hosts</span></span>

<span data-ttu-id="8f1cd-163">Outlook アドインでは、次のように **Hosts** 要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-163">Outlook add-ins specify the **Hosts** element like the following.</span></span>

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

<span data-ttu-id="8f1cd-164">これは、「[Excel、PowerPoint、および Word のマニフェストでのアドイン コマンドの作成](../develop/create-addin-commands.md)」で説明されている **VersionOverrides** 要素内の **Hosts** 要素とは別のものです。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-164">This is separate from the **Hosts** element inside the **VersionOverrides** element, which is discussed in [Create add-in commands in your manifest for Excel, PowerPoint, and Word](../develop/create-addin-commands.md).</span></span>

## <a name="requirements"></a><span data-ttu-id="8f1cd-165">要件</span><span class="sxs-lookup"><span data-stu-id="8f1cd-165">Requirements</span></span>

<span data-ttu-id="8f1cd-p115">**Requirements** 要素は、アドインで使用できる API のセットを指定します。Outlook アドインの場合、要件セットは Mailbox、値は 1.1 以上になっている必要があります。最新の要件セットのバージョンについては、API リファレンスを参照してください。要件セットの詳細については、「[Outlook アドインの API](apis.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p115">The **Requirements** element specifies the set of APIs available to the add-in. For an Outlook add-in, the requirement set must be Mailbox and a value of 1.1 or above. Please refer to the API reference for the latest requirement set version. Refer to the [Outlook add-in APIs](apis.md) for more information on requirement sets.</span></span>

<span data-ttu-id="8f1cd-170">**Requirements** 要素を **VersionOverrides** 要素に表示することもできます。これにより、**VersionOverrides** をサポートするクライアントでアドインが読み込まれたときに、アドインの別の要件を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-170">The **Requirements** element can also appear in the **VersionOverrides** element, allowing the add-in to specify a different requirement when loaded in clients that support **VersionOverrides**.</span></span>

<span data-ttu-id="8f1cd-171">次の例では、**Sets** 要素の **DefaultMinVersion** 属性を使用して office.js バージョン 1.1 以降を要求し、**Set** 要素の **MinVersion** 属性を使用してMailbox 要件セットのバージョン 1.1 を要求しています。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-171">The following example uses the **DefaultMinVersion** attribute of the **Sets** element to require office.js version 1.1 or higher, and the **MinVersion** attribute of the **Set** element to require the Mailbox requirement set version 1.1.</span></span>

```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```

## <a name="form-settings"></a><span data-ttu-id="8f1cd-172">Form settings</span><span class="sxs-lookup"><span data-stu-id="8f1cd-172">Form settings</span></span>

<span data-ttu-id="8f1cd-p116">**FormSettings** 要素は古い Outlook クライアント (スキーマ 1.1 のみをサポートし、**VersionOverrides** はサポートしない) によって使用されます。この要素を使用して、開発者はこのようなクライアントでアドインがどのように表示されるかを定義します。**ItemRead** と **ItemEdit** の 2 つの部分があります。**ItemRead** を使用すると、ユーザーがメッセージと予定を読み込むときに、アドインがどのように表示されるかを指定できます。**ItemEdit** を使用すると、ユーザーが返信や新しいメッセージ、または予定を作成したり (ユーザーが開催者の場合)、予定を編集したりするときに、アドインがどのように表示されるかについて記述できます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p116">The **FormSettings** element is used by older Outlook clients, which only support schema 1.1 and not **VersionOverrides**. Using this element, developers define how the add-in will appear in such clients. There are two parts - **ItemRead** and **ItemEdit**. **ItemRead** is used to specify how the add-in appears when the user reads messages and appointments. **ItemEdit** describes how the add-in appears while the user is composing a reply, new message, new appointment or editing an appointment where they are the organizer.</span></span>

<span data-ttu-id="8f1cd-p117">これらの設定は、**Rule** 要素のアクティブ化ルールと直接関連します。アドインにおいてそのアドインが作成モードのメッセージ上に表示されるように指定する場合は、**ItemEdit** フォームを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p117">These settings are directly related to the activation rules in the **Rule** element. For example, if an add-in specifies that it should appear on a message in compose mode, an **ItemEdit** form must be specified.</span></span>

<span data-ttu-id="8f1cd-180">詳細は、「Schema reference for Office Add-ins manifests (v1.1)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-180">For more details, please refer to the [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>

## <a name="app-domains"></a><span data-ttu-id="8f1cd-181">アプリ ドメイン</span><span class="sxs-lookup"><span data-stu-id="8f1cd-181">App domains</span></span>

<span data-ttu-id="8f1cd-p118">**SourceLocation** 要素に指定するアドインの開始ページのドメインは、そのアドインの既定のドメインです。**AppDomains** 要素と **AppDomain** 要素を使用しない場合は、アドインが別のドメインに移動しようとすると、ブラウザーがそのアドイン ウィンドウの外に新しいウィンドウを開きます。アドインがアドイン ウィンドウ内の別のドメインに移動できるようにするには、アドインのマニフェストに **AppDomains** 要素を追加し、その **AppDomain** サブ要素に各追加ドメインを含めます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p118">The domain of the add-in start page that you specify in the **SourceLocation** element is the default domain for the add-in. Without using the **AppDomains** and **AppDomain** elements, if your add-in attempts to navigate to another domain, the browser will open a new window outside of the add-in pane. In order to allow the add-in to navigate to another domain within the add-in pane, add an **AppDomains** element and include each additional domain in its own **AppDomain** sub-element in the add-in manifest.</span></span>

<span data-ttu-id="8f1cd-185">次のサンプルでは、アドインがアドイン ウィンドウ内で移動できる 2 番目のドメインとして  `https://www.contoso2.com` を指定しています。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-185">The following example specifies a domain  `https://www.contoso2.com` as a second domain that the add-in can navigate to within the add-in pane:</span></span>

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

<span data-ttu-id="8f1cd-186">アプリ ドメインは、ポップアップ ウィンドウと、リッチ クライアントで実行するアドインとの間での Cookie の共有を有効にするためにも必要です。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-186">App domains are also necessary to enable cookie sharing between the pop-out window and the add-in running in the rich client.</span></span>

<span data-ttu-id="8f1cd-187">次の表では、アドインが既定のドメイン外の URL に移動しようとした場合のブラウザーの動作について説明します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-187">The following table describes browser behavior when your add-in attempts to navigate to a URL outside of the add-in's default domain.</span></span>

|<span data-ttu-id="8f1cd-188">Outlook クライアント</span><span class="sxs-lookup"><span data-stu-id="8f1cd-188">Outlook client</span></span>|<span data-ttu-id="8f1cd-189">定義されたドメイン</span><span class="sxs-lookup"><span data-stu-id="8f1cd-189">Domain defined</span></span><br><span data-ttu-id="8f1cd-190">AppDomainsで?</span><span class="sxs-lookup"><span data-stu-id="8f1cd-190">in AppDomains?</span></span>|<span data-ttu-id="8f1cd-191">ブラウザーの動作</span><span class="sxs-lookup"><span data-stu-id="8f1cd-191">Browser behavior</span></span>|
|---|---|---|
|<span data-ttu-id="8f1cd-192">すべてのクライアント</span><span class="sxs-lookup"><span data-stu-id="8f1cd-192">All clients</span></span>|<span data-ttu-id="8f1cd-193">はい</span><span class="sxs-lookup"><span data-stu-id="8f1cd-193">Yes</span></span>|<span data-ttu-id="8f1cd-194">リンクがアドインの作業ウィンドウで開きます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-194">Link opens in add-in task pane.</span></span>|
|<span data-ttu-id="8f1cd-195">Windows 用 Outlook 2016 (1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="8f1cd-195">Outlook 2016 on Windows (one-time purchase)</span></span><br><span data-ttu-id="8f1cd-196">Windows 用 Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="8f1cd-196">Outlook 2013 on Windows</span></span>|<span data-ttu-id="8f1cd-197">いいえ</span><span class="sxs-lookup"><span data-stu-id="8f1cd-197">No</span></span>|<span data-ttu-id="8f1cd-198">リンクが Internet Explorer 11 で開きます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-198">Link opens in Internet Explorer 11.</span></span>|
|<span data-ttu-id="8f1cd-199">その他のクライアント</span><span class="sxs-lookup"><span data-stu-id="8f1cd-199">Other clients</span></span>|<span data-ttu-id="8f1cd-200">いいえ</span><span class="sxs-lookup"><span data-stu-id="8f1cd-200">No</span></span>|<span data-ttu-id="8f1cd-201">リンクがユーザーの既定のブラウザーで開きます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-201">Link opens in user's default browser.</span></span>|

<span data-ttu-id="8f1cd-202">詳細については、「[アドイン ウィンドウで開くドメインの指定](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-202">For more details, see the [Specify domains you want to open in the add-in window](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window).</span></span>

## <a name="permissions"></a><span data-ttu-id="8f1cd-203">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="8f1cd-203">Permissions</span></span>

<span data-ttu-id="8f1cd-p119">**Permissions** 要素には、アドインに必要なアクセス許可が含まれます。通常は、使用する予定の実際のメソッドに応じて、そのアドインに必要な最小限のアクセス許可を指定します。たとえば、新規作成フォームでアクティブ化され、[item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) などのアイテム プロパティを読み取るだけで書き込みはせず、かつ [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を呼び出して Exchange Web サービスの操作にアクセスすることのないメール アドインでは、**ReadItem** アクセス許可を指定する必要があります。利用できるアクセス許可について詳しくは、「[Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p119">The **Permissions** element contains the required permissions for the add-in. In general, you should specify the minimum necessary permission that your add-in needs, depending on the exact methods that you plan to use. For example, a mail add-in that activates in compose forms and only reads but does not write to item properties like [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and does not call [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to access any Exchange Web Services operations should specify **ReadItem** permission. For details on the available permissions, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<span data-ttu-id="8f1cd-208">**メール アドインの 4 層アクセス許可モデル**</span><span class="sxs-lookup"><span data-stu-id="8f1cd-208">**Four-tier permissions model for mail add-ins**</span></span>

![メール アプリ スキーマ v1.1 の 4 層アクセス許可モデル](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a><span data-ttu-id="8f1cd-210">アクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="8f1cd-210">Activation rules</span></span>

<span data-ttu-id="8f1cd-p120">アクティブ化ルールは、**Rule** 要素で指定されます。**Rule** 要素は、1.1 マニフェストの **OfficeApp** 要素の子として表示される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p120">Activation rules are specified in the **Rule** element. The **Rule** element can appear as a child of the **OfficeApp** element in 1.1 manifests.</span></span>

<span data-ttu-id="8f1cd-213">アクティブ化ルールを使用すると、現在選択されているアイテムについての以下の 1 つ以上の条件に基づいてアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-213">Activation rules can be used to activate an add-in based on one or more of the following conditions on the currently selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8f1cd-214">アクティブ化ルールは、**VersionOverrides** 要素をサポートしないクライアントにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-214">Activation rules only apply to clients that do not support the **VersionOverrides** element.</span></span>

- <span data-ttu-id="8f1cd-215">アイテムの種類またはメッセージ クラス、あるいはその両方</span><span class="sxs-lookup"><span data-stu-id="8f1cd-215">The item type and/or message class</span></span>

- <span data-ttu-id="8f1cd-216">特定の種類の既知のリソース (住所または電話番号など) が存在すること</span><span class="sxs-lookup"><span data-stu-id="8f1cd-216">The presence of a specific type of known entity, such as an address or phone number</span></span>

- <span data-ttu-id="8f1cd-217">本文、件名、送信者のメール アドレスにおける正規表現の一致</span><span class="sxs-lookup"><span data-stu-id="8f1cd-217">A regular expression match in the body, subject, or sender email address</span></span>

- <span data-ttu-id="8f1cd-218">添付ファイルが存在すること</span><span class="sxs-lookup"><span data-stu-id="8f1cd-218">The presence of an attachment</span></span>

<span data-ttu-id="8f1cd-219">アクティブ化ルールの詳細とサンプルについては、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-219">For details and samples of activation rules, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>


## <a name="next-steps-add-in-commands"></a><span data-ttu-id="8f1cd-220">次の手順: アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="8f1cd-220">Next steps: Add-in commands</span></span>

<span data-ttu-id="8f1cd-221">基本的なマニフェストを定義した後、アドインのアドイン コマンドを定義します。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-221">After defining a basic manifest, define add-in commands for your add-in.</span></span> <span data-ttu-id="8f1cd-222">アドイン コマンドは、リボン内にボタンを表示して、ユーザーがアドインを簡単かつ直感的な方法でアクティブ化できるようにします。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-222">Add-in commands present a button in the ribbon so users can activate your add-in in a simple, intuitive way.</span></span> <span data-ttu-id="8f1cd-223">詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-223">For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="8f1cd-224">アドイン コマンドを定義するアドインの例については、[command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-224">For an example add-in that defines add-in commands, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span></span>

## <a name="next-steps-add-mobile-support"></a><span data-ttu-id="8f1cd-225">次の手順：モバイル サポートの追加</span><span class="sxs-lookup"><span data-stu-id="8f1cd-225">Next steps: Add mobile support</span></span>

<span data-ttu-id="8f1cd-p122">オプションで、アドインは Outlook モバイルのサポートを追加できます。Outlook モバイルは、Windows と Mac において、Outlook と同じ方法でアドイン コマンドをサポートします。詳しくは、「[Outlook Mobile 用のアドイン コマンドのサポートを追加する](add-mobile-support.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8f1cd-p122">Add-ins can optionally add support for Outlook mobile. Outlook mobile supports add-in commands in a similar fashion to Outlook on Windows and Mac. For more information, see [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8f1cd-229">関連項目</span><span class="sxs-lookup"><span data-stu-id="8f1cd-229">See also</span></span>

- [<span data-ttu-id="8f1cd-230">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="8f1cd-230">Localization for Office Add-ins</span></span>](../develop/localization.md)
- [<span data-ttu-id="8f1cd-231">Outlook アドインに関するプライバシー、アクセス許可、セキュリティ</span><span class="sxs-lookup"><span data-stu-id="8f1cd-231">Privacy, permissions, and security for Outlook add-ins</span></span>](privacy-and-security.md)
- [<span data-ttu-id="8f1cd-232">Outlook アドインの API</span><span class="sxs-lookup"><span data-stu-id="8f1cd-232">Outlook add-in APIs</span></span>](apis.md)
- [<span data-ttu-id="8f1cd-233">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="8f1cd-233">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="8f1cd-234">Office アドイン マニフェストのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="8f1cd-234">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="8f1cd-235">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="8f1cd-235">Design your Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="8f1cd-236">Outlook アドインのアクセス許可を理解する</span><span class="sxs-lookup"><span data-stu-id="8f1cd-236">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="8f1cd-237">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="8f1cd-237">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="8f1cd-238">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="8f1cd-238">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)