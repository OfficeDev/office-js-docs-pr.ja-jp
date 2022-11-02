---
title: Office アドインの XML マニフェスト
description: Office アドインのマニフェストとその使用方法の概要について説明します。
ms.date: 05/24/2022
ms.localizationpriority: high
ms.openlocfilehash: 60368d74cad0d1b8c0562888613d960f52b21a74
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810226"
---
# <a name="office-add-ins-xml-manifest"></a>Office アドインの XML マニフェスト

Office アドインの XML マニフェスト ファイルでは、エンド ユーザーが Office ドキュメントや Office アプリケーションにアドインをインストールして使用するときにアドインをアクティブ化する方法が記述されています。

> [!TIP]
> この記事では、現在の XML 形式のマニフェストについて説明します。 プレビューで使用できる JSON 形式のTeams マニフェストもあります。 詳細については、「[Office アドイン用のTeams マニフェスト (プレビュー)](json-manifest-overview.md)」を参照してください。

XML マニフェスト ファイルを使用すると、Office アドインで次のことができます。

- ID、バージョン、説明、表示名、および既定のロケールを指定することで、アプリ自体について説明する。

- アドインのブランド化に使用するイメージと、Office アプリ リボンで[アドイン コマンド](create-addin-commands.md)に使用する画像を指定する。

- アドインを Office に統合する方法を指定する。アドインによって作成されるカスタム UI (リボンのボタンなど) の統合も含む。

- コンテンツ アドインに必要な既定のサイズ、および Outlook アドインに必要な高さを指定する。

- ドキュメントの読み取り、書き込みなど、Office アドインに必要なアクセス許可を宣言する。

- Outlook アドインでは、アプリがアクティブ化されてメッセージ、予定、または会議出席依頼アイテムを操作するコンテキストを指定するルールを定義する。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a>必要な要素

次の表に、3 種類の Office アドインに必要な要素を示します。

> [!NOTE]
> 親要素内で要素を表示する順序も決まっています。 詳細については、[マニフェスト要素の正しい順序を確認する方法](manifest-element-ordering.md)を参照してください。

### <a name="required-elements-by-office-add-in-type"></a>Office アドインの種類ごとの必要な要素

| 要素                                                                                      | コンテンツ    | 作業ウィンドウ    | メール<br>(Outlook)      |
| :------------------------------------------------------------------------------------------- | :--------: | :----------: | :--------:   |
| [OfficeApp][]                                                                                | 必須   | 必須     | 必須     |
| [Id][]                                                                                       | 必須   | 必須     | 必須     |
| [バージョン][]                                                                                  | 必須   | 必須     | 必須     |
| [ProviderName][]                                                                             | 必須   | 必須     | 必須     |
| [DefaultLocale][]                                                                            | 必須   | 必須     | 必須     |
| [DisplayName][]                                                                              | 必須   | 必須     | 必須     |
| [Description][]                                                                              | 必須   | 必須     | 必須     |
| [IconUrl][]                                                                                  | 必須   | 必須     | 必須     |
| [SupportUrl][]\*\*                                                                           | 必須   | 必須     | 必須     |
| [DefaultSettings (ContentApp)][]<br/>[DefaultSettings (TaskPaneApp)][]                       | 必須   | 必須     | 使用不可|
| [SourceLocation (ContentApp)][]<br/>[SourceLocation (TaskPaneApp)][]<br/>[SourceLocation (MailApp)][]| 必須 | 必須 | 必須   |
| [DesktopSettings][]                                                                          | 使用不可 | 使用不可 | 必須 |
| [Permissions (ContentApp)][]<br/>[Permissions (TaskPaneApp)][]<br/>[Permissions (MailApp)][] | 必須   | 必須     | 必須     |
| [Rule (RuleCollection)][]<br/>[Rule (MailApp)][]                                             | 使用不可 | 使用不可 | 必須 |
| [Requirements (MailApp)][]\*                                                                 | 該当なし| 使用不可 | 必須 |
| [Set][]\*<br/>[Sets (Requirements)][]\*<br/>[Sets (MailAppRequirements)][]\*                 | 必須   | 必須     | 必須     |
| [Form][]\*<br/>[FormSettings][]\*                                                            | 使用不可 | 使用不可 | 必須 |
| [Hosts][]\*                                                                                  | 必須   | 必須     | 省略可能     |

_\*Office アドイン マニフェスト スキーマ バージョン 1.1 で追加されました。_

_\*\* SupportUrl は、AppSource 経由で配布されたアドインに対してのみ必要です。_

<!-- Links for above table -->

[officeapp]: /javascript/api/manifest/officeapp
[id]: /javascript/api/manifest/id
[version]: /javascript/api/manifest/version
[providername]: /javascript/api/manifest/providername
[defaultlocale]: /javascript/api/manifest/defaultlocale
[displayname]: /javascript/api/manifest/displayname
[description]: /javascript/api/manifest/description
[iconurl]: /javascript/api/manifest/iconurl
[supporturl]: /javascript/api/manifest/supporturl
[defaultsettings (contentapp)]: /javascript/api/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /javascript/api/manifest/defaultsettings
[sourcelocation (contentapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /javascript/api/manifest/sourcelocation
[sourcelocation (mailapp)]: /javascript/api/manifest/sourcelocation
[desktopsettings]: /javascript/api/manifest/desktopsettings
[permissions (contentapp)]: /javascript/api/manifest/permissions
[permissions (taskpaneapp)]: /javascript/api/manifest/permissions
[permissions (mailapp)]: /javascript/api/manifest/permissions
[rule (rulecollection)]: /javascript/api/manifest/rule
[rule (mailapp)]: /javascript/api/manifest/rule
[requirements (mailapp)]: /javascript/api/manifest/requirements
[set]: /javascript/api/manifest/set
[sets (mailapprequirements)]: /javascript/api/manifest/sets
[form]: /javascript/api/manifest/form
[formsettings]: /javascript/api/manifest/formsettings
[sets (requirements)]: /javascript/api/manifest/sets
[hosts]: /javascript/api/manifest/hosts

## <a name="hosting-requirements"></a>ホストするための要件

[アドイン コマンド](create-addin-commands.md)などで使用されるすべてのイメージ URI はキャッシュをサポートしている必要があります。 イメージをホストしているサーバーは、HTTP 応答で `no-cache`、`no-store`、または同様のオプションを指定する `Cache-Control` ヘッダーを返しません。

[SourceLocation](/javascript/api/manifest/sourcelocation) 要素で指定されるソース ファイルの場所など、すべての URL は **SSL (HTTPS) でセキュリティ保護されている** べきです。 [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a>AppSource に提出するためのベスト プラクティス

Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.

AppSource に提出するアドインには、[SupportUrl](/javascript/api/manifest/supporturl) 要素も含める必要があります。 詳細については、「[AppSource に提出されたアプリとアドインの検証ポリシー](/legal/marketplace/certification-policies)」をご覧ください。

必ず [AppDomains](/javascript/api/manifest/appdomains) 要素を使い、認証シナリオのために [SourceLocation](/javascript/api/manifest/sourcelocation) 要素で指定されたもの以外のドメインを指定してください。

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a>アドイン ウィンドウで開くドメインの指定

When running in Office on the web, your task pane can be navigated to any URL. However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/javascript/api/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.

このデスクトップの Office の動作を変更するには、マニフェスト ファイルの [AppDomains](/javascript/api/manifest/appdomains) 要素で指定するドメインの一覧で、アドイン ウィンドウで開く各ドメインを指定します。 アドインがこの一覧にあるドメインの URL に移動しようとすると、Office on the web とデスクトップの Office の両方の作業ウィンドウで開きます。 この一覧にない URL に移動しようとすると、その URL はデスクトップの Office 新しいブラウザー ウィンドウ (アドイン ウィンドウとは別のウィンドウ) で開きます。

> [!NOTE]
> この動作に対する例外は 2 つあります。
>
> - これは、アドインのルート ウィンドウに対してのみ適用されます。 アドインページに iframe が埋め込まれている場合、Office デスクトップの場合でも、**\<AppDomains\>** の一覧にあるかどうかにかかわらず、その iframe を任意の URL に転送できます。
> - [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#office-office-ui-displaydialogasync-member(1)) API でダイアログを開く場合、メソッドに渡される URL はアドインと同じドメインにある必要がありますが、ダイアログはデスクトップ Office であっても **\<AppDomains\>** にリストされているかどうかに関係なく、任意の URL にリダイレクトできます。

次に示す XML マニフェストの例では、**\<SourceLocation\>** 要素に指定された `https://www.contoso.com` ドメインでメイン アドイン ページをホストします。 また、この例では、**\<AppDomains\>** 要素リスト内の [AppDomain](/javascript/api/manifest/appdomain) 要素の `https://www.northwindtraders.com` ドメインも指定しています。 アドインが `www.northwindtraders.com` ドメイン内のページに移動すると、Office デスクトップの場合でも、そのページはアドイン ウィンドウで開きます。

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="version-overrides-in-the-manifest"></a>マニフェストでのバージョンの上書き

オプションの [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素は特筆すべきものです。 追加のアドイン機能を有効にする子マークアップが含まれます。 その一部を次に示します。

- Office リボンやメニューをカスタマイズします。
- アドインを実行する埋め込みランタイムでの Office の動作をカスタマイズする。
- アドインが Azure Active Directory やシングル サインオン用 Microsoft Graph と対話する方法を構成します。

子要素 `VersionOverrides` の中には、親要素 `OfficeApp` の値を上書きする値があります。 たとえば、`VersionOverrides` 内の `Hosts` 要素は `OfficeApp` 内の `Hosts` 要素よりも優先されます。

The `VersionOverrides` element has its own schema, actually four of them, depending on the type of add-in and the features it uses. The schemas are:

- [作業ウィンドウ 1.0](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)
- [コンテンツ 1.0](/openspecs/office_file_formats/ms-owemxml/c9cb8dca-e9e7-45a7-86b7-f1f0833ce2c7)
- [メール 1.0](/openspecs/office_file_formats/ms-owemxml/578d8214-2657-4e6a-8485-25899e772fac)
- [メール 1.1](/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)

`VersionOverrides` 要素を使用する場合、`OfficeApp` 要素には適切なスキーマを識別する `xmlns` 属性を含む必要があります。 この属性に設定できる値は以下のとおりです。

- `http://schemas.microsoft.com/office/taskpaneappversionoverrides`
- `http://schemas.microsoft.com/office/contentappversionoverrides`
- `http://schemas.microsoft.com/office/mailappversionoverrides`

`VersionOverrides` 要素自体にもスキーマを指定する `xmlns` 属性が必要です。 設定可能な値は、上記の 3 つと以下に示す値です。

- `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`

`VersionOverrides` 要素には、スキーマ バージョンを指定する `xsi:type` 属性も必要です。 指定できる値は以下のとおりです。

- `VersionOverridesV1_0`
- `VersionOverridesV1_1`

作業ウィンドウ アドインとメール アドインにそれぞれ `VersionOverrides` を使用した例を以下に示します。 バージョン 1.1 のメール `VersionOverrides` を使用する場合は、タイプ 1.0 の親要素 `VersionOverrides` の最後の子要素である必要があることに注意してください。 内側の子要素 `VersionOverrides` の値は、親要素 `VersionOverrides` と祖父母要素 `OfficeApp` の同名要素の値を上書きします。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- child elements omitted -->
</VersionOverrides>
```

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <!-- other child elements omitted -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- child elements omitted -->
  </VersionOverrides>
</VersionOverrides>
```

`VersionOverrides` 要素を含むマニフェストの例については、「[マニフェスト v1.1 XML ファイルの例とスキーマ](#manifest-v11-xml-file-examples-and-schemas)」を参照してください。

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a>Office.js API 呼び出しが行われるドメインを指定する

アドインは、マニフェスト ファイルの [SourceLocation](/javascript/api/manifest/sourcelocation) 要素で参照されているドメインから Office.js API 呼び出しを行うことができます。 アドイン内に、Office.js API にアクセスする必要がある他の IFrame がある場合は、マニフェスト ファイルの [AppDomains](/javascript/api/manifest/appdomains) 要素で指定されているリストにそのソース URL のドメインを追加します。 `AppDomains` リストに含まれていないソースを持つ IFrame が Office.js API 呼び出しを行おうとすると、アドインには[アクセス許可の拒否エラー](../reference/javascript-api-for-office-error-codes.md)が返されます。

## <a name="manifest-v11-xml-file-examples-and-schemas"></a>マニフェスト v1.1 XML ファイルの例とスキーマ

次のセクションでは、コンテンツ、作業ウィンドウ、メール (Outlook) アドイン用のマニフェスト v1.1 XML ファイルの例を示します。

# <a name="task-pane"></a>[作業ウィンドウ](#tab/tabid-1)

[アドイン マニフェストのスキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[Content](#tab/tabid-2)

[アドイン マニフェストのスキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[メール](#tab/tabid-3)

[アドイン マニフェストのスキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a>Office アドインのマニフェストを検証する

[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) に対してマニフェストを検証する方法については、「[Office アドインのマニフェストを検証する](../testing/troubleshoot-manifest.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [マニフェスト要素の正しい順序を確認する方法](manifest-element-ordering.md)
- [マニフェストでアドイン コマンドを作成する](create-addin-commands.md)
- [Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md)
- [Office アドインのローカライズ](localization.md)
- [Office アドイン マニフェストのスキーマ参照](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
- [API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)
- [同等な COM アドインを特定する](make-office-add-in-compatible-with-existing-com-add-in.md)
- [アドインでの API 使用についてアクセス許可を要求する](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [Office アドインのマニフェストを検証する](../testing/troubleshoot-manifest.md)
