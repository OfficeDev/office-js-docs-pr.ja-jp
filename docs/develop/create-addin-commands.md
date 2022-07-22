---
title: Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する
description: マニフェストで VersionOverrides を使用して、Excel、PowerPoint、Word のアドイン コマンドを定義します。 UI 要素を作成し、ボタンやリストを追加し、操作を実行するために、アドイン コマンドを使用します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4df14158d6a9fde9d18e75632c44e40fca235b8d
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958707"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する

> [!NOTE]
> アドイン コマンドは、Outlook でもサポートされています。 詳細については、「[Outlook のアドイン コマンド」を](../outlook/add-in-commands-for-outlook.md)参照してください。

マニフェストで **[VersionOverrides](/javascript/api/manifest/versionoverrides)** を使用して、Excel、PowerPoint、Word のアドイン コマンドを定義します。 アドイン コマンドは、アクションを実行する指定された UI 要素を使用して、既定の Office ユーザー インターフェイス (UI) をカスタマイズする簡単な方法を提供します。 アドイン コマンドの概要については、「 [Excel、PowerPoint、Word のアドイン コマンド](../design/add-in-commands.md)」を参照してください。

この記事では、マニフェストを編集してアドイン コマンドを定義する方法と [、関数](../design/add-in-commands.md#types-of-add-in-commands)コマンドのコードを作成する方法について説明します。 次の図に、アドイン コマンドを定義するのに使用される要素の階層を示します。 これらの要素は、この記事で詳細に説明します。

![マニフェスト内のアドイン コマンド要素の概要。 ここでの最上位ノードは、子ホストとリソースを含む VersionOverrides です。 [ホスト]、[DesktopFormFactor] の順に選択します。 DesktopFormFactor の下には FunctionFile と ExtensionPoint があります。 ExtensionPoint の下には、CustomTab または OfficeTab と Office メニューがあります。 [CustomTab] または [Office] タブの下には、[グループ] と [コントロール] 、[アクション] の順に選択します。 [Office] メニューの [コントロール] と [アクション] の順に選択します。 リソース (VersionOverrides の子) の下には、イメージ、URL、ShortStrings、LongStrings があります。](../images/version-overrides.png)

## <a name="step-1-create-the-project"></a>手順 1: プロジェクトを作成する

[Excel 作業ウィンドウ アドインの作成](../quickstarts/excel-quickstart-jquery.md)などのクイック スタートの 1 つに従ってプロジェクトを作成することをお勧めします。 Excel、Word、および PowerPoint の各クイック スタートでは、作業ウィンドウを表示するアドイン コマンド (ボタン) が既に含まれているプロジェクトが生成されます。 アドイン コマンドを使用する前 [に、Excel、Word、PowerPoint の](../design/add-in-commands.md) アドイン コマンドを読み取っていることを確認します。

## <a name="step-2-create-a-task-pane-add-in"></a>手順 2: 作業ウィンドウ アドインを作成する

アドイン コマンドの使用を開始するには、まず作業ウィンドウ アドインを作成してから、この記事で説明するようにアドインのマニフェストを変更する必要があります。 コンテンツ アドインでアドイン コマンドを使用することはできません。既存のマニフェストを更新する場合は、「[手順 3: VersionOverrides](#step-3-add-versionoverrides-element) 要素の追加」の説明に従って、適切な **XML 名前空間** を追加し、マニフェストに要素を追加 **\<VersionOverrides\>** する必要があります。

次の例は、Office 2013 アドインのマニフェストを示します。VersionOverrides 要素がないため、このマニフェストにはアドイン コマンドがありません。 このマニフェストには、要素がないため、アドイン コマンドはありません **\<VersionOverrides\>** 。 Office 2013 ではアドイン コマンドはサポートされていませんが、このマニフェストに追加 **\<VersionOverrides\>** することで、アドインは Office 2013 と Office 2016 の両方で実行されます。 Office 2013 では、アドインはアドイン コマンドを表示せず、アドインを 1 つの作業ウィンドウ アドインとして実行する値 **\<SourceLocation\>** を使用します。 Office 2016 では、要素が含まれていない **\<VersionOverrides\>** 場合は、アドインの作業ウィンドウが自動的に開き **\<SourceLocation\>**、 . ただし、アドインを含める **\<VersionOverrides\>** 場合、アドインにはアドイン コマンドのみが表示され、アドインの作業ウィンドウは最初に表示されません。
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
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

## <a name="step-3-add-versionoverrides-element"></a>手順 3: VersionOverrides 要素を追加する

要素は **\<VersionOverrides\>** 、アドイン コマンドの定義を含むルート要素です。 **\<VersionOverrides\>** は、マニフェスト内の要素の **\<OfficeApp\>** 子要素です。 次の表に、要素の属性を **\<VersionOverrides\>** 示します。

|属性|説明|
|:-----|:-----|
|**xmlns** <br/> | 必須です。 スキーマの場所。`http://schemas.microsoft.com/office/taskpaneappversionoverrides` にする必要があります。 <br/> |
|**xsi:type** <br/> |必須。スキーマのバージョン。この記事で説明されているスキーマのバージョンは "VersionOverridesV1_0" です。  <br/> |

次の表は、 **\<VersionOverrides\>**.
  
|要素|説明|
|:-----|:-----|
|**\<Description\>** <br/> |省略可能。 アドインの説明です。 この子 **\<Description\>** 要素は、マニフェストの親部分の前 **\<Description\>** の要素をオーバーライドします。 この **\<Description\>** 要素 **の resid** 属性は、要素の **\<String\>** **ID** に設定されます。 要素には **\<String\>** 、 **\<Description\>**. <br/> |
|**\<Requirements\>** <br/> |省略可能。 アドインに必要な最小の Office.js のセットおよびバージョンを指定します。 この子 **\<Requirements\>** 要素は、マニフェストの **\<Requirements\>** 親部分の要素をオーバーライドします。 詳細については、「 [Office アプリケーションと API 要件の指定](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。  <br/> |
|**\<Hosts\>** <br/> |必須です。 Office アプリケーションのコレクションを指定します。 子 **\<Hosts\>** 要素は、マニフェストの **\<Hosts\>** 親部分の要素をオーバーライドします。 "Workbook" または "Document" に設定された **xsi:type** 属性を含める必要があります。 <br/> |
|**\<Resources\>** <br/> |マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。 たとえば、要素の **\<Description\>** 値 **\<Resources\>** は 、 . この要素については **\<Resources\>** 、この記事の [後半の「手順 7: Resources 要素を追加する](#step-7-add-the-resources-element) 」で説明します。 <br/> |

次の例では、要素とその子要素を **\<VersionOverrides\>** 使用する方法を示します。

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

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>手順 4: Hosts、Host、DesktopFormFactor 要素を追加する

要素には **\<Hosts\>** 、1 つ以上の要素が **\<Host\>** 含まれています。 要素は **\<Host\>** 、特定の Office アプリケーションを指定します。 この **\<Host\>** 要素には、アドインをその Office アプリケーションにインストールした後に表示するアドイン コマンドを指定する子要素が含まれています。 2 つ以上の異なる Office アプリケーションで同じアドイン コマンドを表示するには、それぞれの **\<Host\>** 子要素を複製する必要があります。

この要素は **\<DesktopFormFactor\>**、Office on the web (ブラウザー) と Windows で実行されるアドインの設定を指定します。

次に、、**\<Host\>** および要素の **\<Hosts\>** 例を **\<DesktopFormFactor\>** 示します。

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

## <a name="step-5-add-the-functionfile-element"></a>手順 5: FunctionFile 要素を追加する

この要素は **\<FunctionFile\>** 、アドイン コマンドが **ExecuteFunction** アクションを使用するときに実行する JavaScript コードを含むファイルを指定します (説明については [、ボタン コントロール](/javascript/api/manifest/control-button) を参照してください)。 **\<FunctionFile\>** 要素の **resid** 属性は、アドイン コマンドで必要なすべての JavaScript ファイルを含む HTML ファイルに設定されます。 You can't link directly to a JavaScript file. You can only link to an HTML file. ファイル名は、要素内の **\<Url\>****\<Resources\>** 要素として指定されます。

要素の例を次に示します **\<FunctionFile\>** 。
  
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
> JavaScript コードが `Office.initialize` を呼び出していることを確認します。

要素によって参照される HTML ファイル内の **\<FunctionFile\>** JavaScript は、呼び出す `Office.initialize`必要があります。 要素 ( **\<FunctionName\>** 説明については [Button コントロール](/javascript/api/manifest/control-button) を参照) では、次の **\<FunctionFile\>** 関数を使用します。

次のコードは、によって使用される関数を実装する方法を **\<FunctionName\>** 示しています。

```html
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
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
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> **event.completed** に対する呼び出しにより、イベントが正常に処理されたことが通知されます。同一のアドイン コマンドを複数回クリックするなど、関数を複数回呼び出すと、すべてのイベントが自動的にキューに入れられます。最初のイベントが自動的に実行され、その他のイベントはキューに残ります。関数により **event.completed** が呼び出されると、キューに入れられている、その関数に対する次の呼び出しが実行されます。**event.completed** を実装する必要があります。実装しない場合、関数は実行されません。

## <a name="step-6-add-extensionpoint-elements"></a>手順 6: ExtensionPoint 要素を追加する

この要素は **\<ExtensionPoint\>** 、Office UI でアドイン コマンドを表示する場所を定義します。 これらの **xsi:type** 値を使用して要素を定義 **\<ExtensionPoint\>** できます。

- **PrimaryCommandSurface**。Office のリボンを参照します。

- **ContextMenu**。Office UI で右クリックしたときに表示されるショートカット メニューです。

次の例では、**PrimaryCommandSurface** および **ContextMenu** 属性値を持つ要素と、それぞれに使用する必要がある子要素を使用する方法 **\<ExtensionPoint\>** を示します。

> [!IMPORTANT]
> ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。`<CustomTab id="mycompanyname.mygroupname">`
  
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

|要素|説明|
|:-----|:-----|
|**\<CustomTab\>** <br/> |リボンに ( **PrimaryCommandSurface** を使用して) カスタム タブを追加する場合は必須です。 要素を使用する **\<CustomTab\>** 場合は、要素を **\<OfficeTab\>** 使用できません。 **id** 属性が必要です。 <br/> |
|**\<OfficeTab\>** <br/> |既定の Office アプリ リボン タブを拡張する場合に必要です ( **PrimaryCommandSurface** を使用)。 要素を使用する **\<OfficeTab\>** 場合は、要素を **\<CustomTab\>** 使用できません。 <br/> **id** 属性で使用するその他のタブ値については、[Office アプリの既定のリボン タブのタブ値に関するページを参照してください](/javascript/api/manifest/officetab)。  <br/> |
|**\<OfficeMenu\>** <br/> | 既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 <br/> Excel または Word の場合は **ContextMenuText**。ユーザーがテキストを選択し、選択したテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。<br/> Excel の場合は **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。 <br/> |
|**\<Group\>** <br/> |タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 <br/> |
|**\<Label\>** <br/> |必須。 グループのラベルです。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<ShortStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Icon\>** <br/> |必須。 小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。 **resid** 属性は、要素の **id** 属性の値に設定する **\<Image\>** 必要があります。 要素は **\<Image\>** 要素の **\<Images\>** 子要素であり、要素の **\<Resources\>** 子要素です。 **size** 属性は、イメージのサイズをピクセル単位で指定します。 3 つのイメージのサイズ (16、32、80) が必要です。 5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。 <br/> |
|**\<Tooltip\>** <br/> |省略可能。 グループのツールヒント。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<LongStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Control\>** <br/> |各グループには、少なくとも 1 つのコントロールが必要です。 要素には **\<Control\>****、ボタン** または **メニュー** を指定できます。 **メニューを** 使用して、ボタン コントロールのドロップダウン リストを指定します。 現在は、ボタンとメニューのみがサポートされています。 詳細については、 [ボタン コントロール](/javascript/api/manifest/control-button) と [メニュー コントロール](/javascript/api/manifest/control-menu) に関するページを参照してください。 <br/>**メモ：** トラブルシューティングを簡単にするために、要素と関連 **\<Resources\>** する子要素を **\<Control\>** 一度に 1 つずつ追加することをお勧めします。          |

### <a name="button-controls"></a>ボタン コントロール

ボタンは、ユーザーが選択したときに 1 つの操作を実行します。 JavaScript 関数を実行するか、作業ウィンドウを表示することができます。 次の例は、2 つのボタンを定義する方法を示しています。 最初のボタンは UI を表示しないで JavaScript 関数を実行し、2 つ目のボタンは作業ウィンドウを表示します。 要素内:**\<Control\>**

- **type** 属性は必須であり、**Button** に設定する必要があります。

- 要素の **\<Control\>** **id** 属性は、最大 125 文字の文字列です。

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

|要素|説明|
|:-----|:-----|
|**\<Label\>** <br/> |必須。 ボタンのテキストです。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<ShortStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Tooltip\>** <br/> |省略可能。 ボタンのツールヒント。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<LongStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Supertip\>** <br/> | 必須。このボタンのヒントであり、次のものによって定義されます。 <br/> **Title** <br/>  必ず指定します。 ヒントのテキストです。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<ShortStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> **\<Description\>** <br/>  必ず指定します。 ヒントの記述です。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<LongStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Icon\>** <br/> | 必須です。 ボタンの **\<Image\>** 要素を格納します。 画像ファイルは必ず .png 形式です。 <br/> **\<Image\>** <br/>  ボタンに表示する画像を定義します。 **resid** 属性は、要素の **id** 属性の値に設定する **\<Image\>** 必要があります。 要素は **\<Image\>** 要素の **\<Images\>** 子要素であり、要素の **\<Resources\>** 子要素です。 **size** 属性は、イメージのサイズをピクセル単位で示します。 3 つのイメージのサイズ (16、32、80) が必要です。 5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。 <br/> |
|**\<Action\>** <br/> | 必須。ユーザーがボタンを選択したときに実行する操作を指定します。**xsi:type** 属性の値は、次のいずれかを指定できます。 <br/> **ExecuteFunction**: によって参照される **\<FunctionFile\>** ファイル内にある JavaScript 関数を実行します。 子要素は **\<FunctionName\>** 、実行する関数の名前を指定します。 <br/> アドインの作業ウィンドウを表示する **ShowTaskPane**。 子要素は **\<SourceLocation\>** 、表示するページのソース ファイルの場所を指定します。 **resid** 属性は、要素内の要素内 **\<Urls\>** の要素の **id** 属性の **\<Url\>** 値に設定する **\<Resources\>** 必要があります。 <br/> |

### <a name="menu-controls"></a>Menu コントロール

**Menu** コントロールは、**PrimaryCommandSurface** または **ContextMenu** のどちらかで使用できます。また、以下の項目を定義します。
  
- ルートレベルのメニュー項目。
- サブメニュー項目のリスト。

**PrimaryCommandSurface** と共に使用すると、ルートのメニュー項目がリボンのボタンとして表示されます。ボタンを選択すると、サブメニューがドロップダウン リストとして表示されます。**ContextMenu** と共に使用すると、サブメニューのあるメニュー項目がコンテキスト メニューに挿入されます。どちらの場合も、各サブメニュー項目は JavaScript 関数を実行するか、作業ウィンドウを表示することができます。現時点では、サブメニューの 1 つのレベルのみがサポートされます。

次の例では、2 つのサブメニュー項目を持つメニュー項目を定義する方法を示します。 最初のサブメニュー項目は作業ウィンドウを示し、2 番目のサブメニュー項目は JavaScript 関数を実行します。 要素内:**\<Control\>**

- **xsi:type** 属性は必須であり、**Menu** に設定する必要があります。
- **id** 属性は、最大 125 文字の文字列です。

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

|要素|説明|
|:-----|:-----|
|**\<Label\>** <br/> |必須。 ルートのメニュー項目のテキスト。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<ShortStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Tooltip\>** <br/> |省略可能。 メニューのツールヒント。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<LongStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<SuperTip\>** <br/> | 必須。 メニューのヒントであり、次のものによって定義されます。 <br/> **\<Title\>** <br/>  必須です。 ヒントのテキスト。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<ShortStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> **\<Description\>** <br/>  必ず指定します。 ヒントの記述です。 **resid** 属性は、要素の **id** 属性の値に設定する **\<String\>** 必要があります。 要素は **\<String\>** 要素の **\<LongStrings\>** 子要素であり、要素の **\<Resources\>** 子要素です。 <br/> |
|**\<Icon\>** <br/> | 必須です。 メニューの **\<Image\>** 要素を含みます。 画像ファイルは必ず .png 形式です。 <br/> **\<Image\>** <br/>  メニューのの画像。 **resid** 属性は、要素の **id** 属性の値に設定する **\<Image\>** 必要があります。 要素は **\<Image\>** 要素の **\<Images\>** 子要素であり、要素の **\<Resources\>** 子要素です。 **size** 属性は、イメージのサイズをピクセル単位で示します。 3 つのイメージのサイズ (16、32、80) が必要です。 20、24、40、48、64 の 5 つのオプション サイズ (ピクセル単位) もサポートされています。 <br/> |
|**\<Items\>** <br/> |必須です。 各サブメニュー項目の **\<Item\>** 要素を格納します。 各 **\<Item\>** 要素には、 [Button コントロール](/javascript/api/manifest/control-button)と同じ子要素が含まれています。  <br/> |

## <a name="step-7-add-the-resources-element"></a>手順 7: Resources 要素を追加する

要素には **\<Resources\>** 、要素のさまざまな子要素によって使用されるリソースが **\<VersionOverrides\>** 含まれています。 リソースには、アイコン、文字列、URL が含まれます。 An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.
  
要素の使用方法の例を次に **\<Resources\>** 示します。 各リソースには、特定のロケールに対して異なるリソースを定義するための 1 つ以上 **\<Override\>** の子要素を含めることができます。

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

|関連情報|説明|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | イメージ ファイルへの HTTPS URL を指定します。各イメージは、次の 3 つの必須のイメージ サイズを定義する必要があります。 <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  次のイメージ サイズもサポートされますが、必須ではありません。 <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |HTTPS URL の場所を指定します。 URL には最大 2048 文字まで指定できます。  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |テキストと **\<Label\>****\<Title\>** 要素。 それぞれに **\<String\>** 最大 125 文字が含まれます。 <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |テキストと **\<Tooltip\>****\<Description\>** 要素。 それぞれに **\<String\>** 最大 250 文字が含まれます。 <br/> |

> [!NOTE]
> および要素内のすべての URL に対して Secure Sockets Layer (SSL) を使用する **\<Image\>****\<Url\>** 必要があります。

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>Office アプリの既定のリボン タブのタブ値

Excel および Word で、既定の Office UI タブを使用することで、リボンにアドイン コマンドを追加できます。 次の表に、要素の **id** 属性に使用できる値を **\<OfficeTab\>** 示します。 タブの値は大文字と小文字を区別します。

|Office クライアント アプリケーション|タブの値|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>関連項目

- [Excel、PowerPoint、Word のアドイン コマンド](../design/add-in-commands.md)
- [サンプル: コマンド ボタンを使用して Excel アドインを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [サンプル: コマンド ボタンを使用して Word アドインを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [サンプル: コマンド ボタンを使用して PowerPoint アドインを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
