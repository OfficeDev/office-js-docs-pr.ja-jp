---
title: Outlook アドイン マニフェスト
description: このマニフェストでは、 Outlook アドインが Outlook クラインアント間でどのように統合されるかを、例を交えて説明します。
ms.date: 05/27/2020
localization_priority: Priority
ms.openlocfilehash: 8c5a31248f68e8f8b5b6ab4b2cf12c9bb969e062f0dccd68c8f5d7c3f5262452
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094194"
---
# <a name="outlook-add-in-manifests"></a>Outlook アドイン マニフェスト

Outlook アドインは XML アドイン マニフェストと Web ページの 2 つのコンポーネントで構成されています。これらは Office アドイン (office.js) の JavaScript ライブラリでサポートされます。マニフェストは、アドインが Outlook クライアント間でどのように統合されるかを説明します。次に例を示します。

 > [!NOTE]
 > 次のサンプルの URL 値はすべて "https://appdemo.contoso.com" で始まります。この値はプレースホルダーであり、有効な実際のマニフェストでは、この部分には有効な HTTPS Web URL が入ります。

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

## <a name="schema-versions"></a>スキーマのバージョン

すべての Outlook クライアントで最新の機能がサポートされているわけではありません。一部の Outlook ユーザーは前のバージョンの Outlook を使用していることがあります。スキーマのバージョンにより、開発者は下位互換性のあるアドインを作成することができます。その際、使用可能な場合には最新の機能を使用しますが、前のバージョンも引き続き機能します。

マニフェストの **VersionOverrides** 要素が、この一例です。**VersionOverrides** 内で定義されたすべての要素は、マニフェストの他の部分にある同じ要素をオーバーライドします。つまり、Outlook は、可能な場合は常に、**VersionOverrides** セクションにあるものを使用してアドインをセットアップします。ただし、Outlook のバージョンが特定のバージョンの **VersionOverrides** をサポートしていない場合、Outlook はこれを無視して、マニフェストの残りの部分の情報のみを使用します。 

このアプローチでは、開発者は個別のマニフェストを複数作成する必要がなく、すべてを 1 つのファイルで定義することになります。

現在のスキーマのバージョンは次のとおりです。


|バージョン|説明|
|:-----|:-----|
|v1.0|Office JavaScript API バージョン 1.0 をサポートします。Outlook アドインであれば、閲覧フォームがサポートされることになります。 |
|v1.1|Office JavaScript API バージョン 1.1 と **VersionOverrides** をサポートします。Outlook アドインで、新規作成フォームもサポートされることになります。|
|**VersionOverrides** 1.0|Office JavaScript API の最新バージョンをサポートします。これは、アドイン コマンドをサポートします。|
|**VersionOverrides** 1.1|Office JavaScript API の最新バージョンをサポートします。これは、アドイン コマンドをサポートし、[ピン留め可能な作業ウィンドウ](pinnable-taskpane.md)やモバイル アドインなどの、より新しい機能のサポートを追加します。|

この記事では、v1.1 マニフェストの要件を取り上げます。アドイン マニフェストで **VersionOverrides** 要素を使用するとしても、**VersionOverrides** をサポートしない以前のクライアントでアドインが機能できるように 1.1 マニフェスト要素を組み込むことは重要です。

> [!NOTE]
> Outlook では、マニフェストの検証にスキーマを使用します。スキーマは、マニフェスト内の要素が特定の順序に従うことを要求します。要求されている順序に従わない要素が含まれていると、アドインをサイドロードするときにエラーが発生することがあります。[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) をダウンロードすると、要求されている順序に要素を並べてマニフェストを作成するのに役立ちます。

## <a name="root-element"></a>ルート要素

Outlook アドイン マニフェストのルート要素は **OfficeApp** です。この要素はまた、既定の名前空間、スキーマのバージョン、アドインの種類を宣言します。開始タグと終了タグの間にマニフェストのその他すべての要素を配置します。ルート要素の例を以下に示します。


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

## <a name="version"></a>バージョン

これは、特定のアドインのバージョンです。開発者がマニフェストの一部を更新する場合、バージョンの番号も増やす必要があります。このようにすることで、新しいマニフェストがインストールされると、既存のマニフェストが上書きされ、ユーザーは新機能を使用できるようになります。このアドインがストアに送信されている場合は、新しいマニフェストをもう一度送信して検証する必要があります。承認されると、数時間以内に、このアドインのユーザーは更新された新しいマニフェストを自動的に使用できるようになります。

アドインに必要なアクセス許可が変更された場合、ユーザーは、アップグレードを行いアドインに再同意するように求められます。管理者が組織全体にこのアドインをインストール済みである場合、管理者がまず再同意する必要があります。それまでの間、ユーザーには引き続き古い機能が表示されます。

## <a name="versionoverrides"></a>VersionOverrides

**VersionOverrides** 要素は、[アドイン コマンド](add-in-commands-for-outlook.md)の情報の場所です。

この要素は、アドインによって[モバイル アドイン](add-mobile-support.md)のサポートが定義される場所でもあります。

この要素の説明については、「[Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する](../develop/create-addin-commands.md)」を参照してください。

## <a name="localization"></a>ローカライズ

名前、説明、読み込む URL など、アドインのいくつかの側面は、各種のロケール用にローカライズする必要があります。これらの要素は、既定値を指定してから、**VersionOverrides** 要素内の **Resources** 要素でロケールのオーバーライドを指定することによって簡単にローカライズできます。画像、URL、文字列をオーバーライドする方法を次に示します。


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

スキーマ リファレンスには、ローカライズできる要素に関する詳しい情報が含まれています。

## <a name="hosts"></a>Hosts

Outlook アドインでは、次のように **Hosts** 要素を指定します。

```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

これは、「[Excel、PowerPoint、および Word のマニフェストでのアドイン コマンドの作成](../develop/create-addin-commands.md)」で説明されている **VersionOverrides** 要素内の **Hosts** 要素とは別のものです。

## <a name="requirements"></a>要件

**Requirements** 要素は、アドインで使用できる API のセットを指定します。Outlook アドインの場合、要件セットは Mailbox、値は 1.1 以上になっている必要があります。最新の要件セットのバージョンについては、API リファレンスを参照してください。要件セットの詳細については、「[Outlook アドインの API](apis.md)」を参照してください。

**Requirements** 要素を **VersionOverrides** 要素に表示することもできます。これにより、**VersionOverrides** をサポートするクライアントでアドインが読み込まれたときに、アドインの別の要件を指定できます。

次の例では、**Sets** 要素の **DefaultMinVersion** 属性を使用して office.js バージョン 1.1 以降を要求し、**Set** 要素の **MinVersion** 属性を使用してMailbox 要件セットのバージョン 1.1 を要求しています。

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

## <a name="form-settings"></a>Form settings

**FormSettings** 要素は古い Outlook クライアント (スキーマ 1.1 のみをサポートし、**VersionOverrides** はサポートしない) によって使用されます。この要素を使用して、開発者はこのようなクライアントでアドインがどのように表示されるかを定義します。**ItemRead** と **ItemEdit** の 2 つの部分があります。**ItemRead** を使用すると、ユーザーがメッセージと予定を読み込むときに、アドインがどのように表示されるかを指定できます。**ItemEdit** を使用すると、ユーザーが返信や新しいメッセージ、または予定を作成したり (ユーザーが開催者の場合)、予定を編集したりするときに、アドインがどのように表示されるかについて記述できます。

これらの設定は、**Rule** 要素のアクティブ化ルールと直接関連します。アドインにおいてそのアドインが作成モードのメッセージ上に表示されるように指定する場合は、**ItemEdit** フォームを指定する必要があります。

詳細は、「Schema reference for Office Add-ins manifests (v1.1)」をご覧ください。

## <a name="app-domains"></a>アプリ ドメイン

**SourceLocation** 要素に指定するアドインの開始ページのドメインは、そのアドインの既定のドメインです。**AppDomains** 要素と **AppDomain** 要素を使用しない場合は、アドインが別のドメインに移動しようとすると、ブラウザーがそのアドイン ウィンドウの外に新しいウィンドウを開きます。アドインがアドイン ウィンドウ内の別のドメインに移動できるようにするには、アドインのマニフェストに **AppDomains** 要素を追加し、その **AppDomain** サブ要素に各追加ドメインを含めます。

次の例では、アドインがアドイン ウィンドウ内で移動できる 2 番目のドメインとして `https://www.contoso2.com` を指定しています。

```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

アプリ ドメインは、ポップアップ ウィンドウと、リッチ クライアントで実行するアドインとの間での Cookie の共有を有効にするためにも必要です。

次の表では、アドインが既定のドメイン外の URL に移動しようとした場合のブラウザーの動作について説明します。

|Outlook クライアント|定義されたドメイン<br>AppDomainsで?|ブラウザーの動作|
|---|---|---|
|すべてのクライアント|はい|リンクがアドインの作業ウィンドウで開きます。|
|Windows 用 Outlook 2016 (1 回限りの購入)<br>Windows 用 Outlook 2013|いいえ|リンクが Internet Explorer 11 で開きます。|
|その他のクライアント|いいえ|リンクがユーザーの既定のブラウザーで開きます。|

詳細については、「[アドイン ウィンドウで開くドメインの指定](../develop/add-in-manifests.md?tabs=tabid-1#specify-domains-you-want-to-open-in-the-add-in-window)」を参照してください。

## <a name="permissions"></a>アクセス許可

**Permissions** 要素には、アドインに必要なアクセス許可が含まれます。通常は、使用する予定の実際のメソッドに応じて、そのアドインに必要な最小限のアクセス許可を指定します。たとえば、新規作成フォームでアクティブ化され、[item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) などのアイテム プロパティを読み取るだけで書き込みはせず、かつ [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を呼び出して Exchange Web サービスの操作にアクセスすることのないメール アドインでは、**ReadItem** アクセス許可を指定する必要があります。利用できるアクセス許可について詳しくは、「[Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)」を参照してください。

**メール アドインの 4 層アクセス許可モデル**

![メール アプリ スキーマ v1.1 の 4 層アクセス許可モデル。](../images/add-in-permission-tiers.png)

```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>
```

## <a name="activation-rules"></a>アクティブ化ルール

アクティブ化ルールは、**Rule** 要素で指定されます。**Rule** 要素は、1.1 マニフェストの **OfficeApp** 要素の子として表示される可能性があります。

アクティブ化ルールを使用すると、現在選択されているアイテムについての以下の 1 つ以上の条件に基づいてアドインをアクティブ化できます。

> [!NOTE]
> アクティブ化ルールは、**VersionOverrides** 要素をサポートしないクライアントにのみ適用されます。

- アイテムの種類またはメッセージ クラス、あるいはその両方

- 特定の種類の既知のリソース (住所または電話番号など) が存在すること

- 本文、件名、送信者のメール アドレスにおける正規表現の一致

- 添付ファイルが存在すること

アクティブ化ルールの詳細とサンプルについては、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。


## <a name="next-steps-add-in-commands"></a>次の手順: アドイン コマンド

基本のマニフェストを定義したら、 アドインのアドイン コマンドを定義します。アドイン コマンドは、リボン内にボタンを表示して、ユーザーがアドインを簡単かつ直感的な方法でアクティブ化できるようにします。詳細は、「[Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。

アドイン コマンドを定義するアドインの例については、[command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。

## <a name="next-steps-add-mobile-support"></a>次の手順：モバイル サポートの追加

オプションで、アドインは Outlook モバイルのサポートを追加できます。Outlook モバイルは、Windows と Mac において、Outlook と同じ方法でアドイン コマンドをサポートします。詳しくは、「[Outlook Mobile 用のアドイン コマンドのサポートを追加する](add-mobile-support.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office アドインのローカライズ](../develop/localization.md)
- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](privacy-and-security.md)
- [Outlook アドインの API](apis.md)
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドイン マニフェストのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)