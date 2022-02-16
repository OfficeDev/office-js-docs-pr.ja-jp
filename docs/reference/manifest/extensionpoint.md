---
title: マニフェスト ファイルの ExtensionPoint 要素
description: Office UI でアドインが機能を公開する場所を定義します。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8ccc08a9c0d42edf89c904b8809a530239be4c
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855633"
---
# <a name="extensionpoint-element"></a>ExtensionPoint 要素

 Office UI でアドインが機能を公開する場所を定義します。 **ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  はい  | 定義される拡張点の種類。 使用できる値は、Office Host 要素の値で定義されているホスト アプリケーション **によって** 異なっています。|

## <a name="extension-points-for-excel-onenote-powerpoint-and-word-add-in-commands"></a>アドイン、Excel、OneNote、PowerPoint、および Word アドイン コマンドの拡張ポイント

これらのホストの一部またはすべてで使用できる拡張ポイントは 3 種類あります。

- [PrimaryCommandSurface](#primarycommandsurface) (Word、Excel、PowerPoint、OneNote に有効) - Office のリボン。
- [ContextMenu](#contextmenu) (Word、Excel、PowerPoint、および OneNote で有効) - Office UI で長押し (または右クリック) すると表示されるショートカット メニュー。
- [CustomFunctions](#customfunctions) (Excel の場合のみ有効) - JavaScript で作成されたカスタム関数で、Excel。

これらの種類の拡張ポイントの子要素と例については、次のサブセクションを参照してください。

### <a name="primarycommandsurface"></a>PrimaryCommandSurface

Word、Excel、PowerPoint、OneNoteの主なコマンド サーフェスがリボンです。

#### <a name="child-elements"></a>子要素

|要素|説明|
|:-----|:-----|
|[CustomTab](customtab.md|カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 |
|[OfficeTab](officetab.md)|既定のリボン タブ (**PrimaryCommandSurface** をOffice アプリする場合は必須です。 **OfficeTab 要素を使用する** 場合は、**CustomTab 要素を使用** することはできません。|

#### <a name="example"></a>例

次の例は、**PrimaryCommandSurface で ExtensionPoint 要素を使用する方法を示しています**。 リボンにカスタム タブを追加します。

> [!IMPORTANT]
> ID 属性を含む要素では、一意の ID を指定してください。

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### <a name="contextmenu"></a>ContextMenu

コンテキスト メニューは、UI で右クリックすると表示されるショートカット Officeです。

#### <a name="child-elements"></a>子要素
 
|要素|説明|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|アドイン コマンドを既定のコンテキスト メニュー (ContextMenu を使用) に追加する場合 **は必須です**。 **id 属性は**、次のいずれかの文字列に設定する必要があります。 <br/> - **コンテキスト メニューを開** く必要がある場合、ユーザーが選択したテキストを右クリックすると ContextMenuText。 <br/> - **ContextMenuCell** を使用すると、ユーザーがスプレッドシート上のセルを右クリックすると、コンテキスト メニューがExcelされます。|

#### <a name="example"></a>例

次に、ユーザー設定のコンテキスト メニューをスプレッドシート内のセルExcelします。

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### <a name="customfunctions"></a>CustomFunctions

JavaScript または TypeScript で作成されたカスタム関数で、Excel。

#### <a name="child-elements"></a>子要素

|要素|説明|
|:-----|:-----|
|[Script](script.md)|必須です。 カスタム関数の定義と登録コードを含む JavaScript ファイルへのリンク。|
|[Page](page.md)|必須です。 カスタム関数についての HTML ページにリンクします。|
|[MetaData](metadata.md)|必須です。 Excel でカスタム関数によって使用されるメタデータの設定を定義します。|
|[Namespace](namespace.md)|省略可能。 Excel でカスタム関数によって使用される名前空間を定義します。|

#### <a name="example"></a>例

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## <a name="extension-points-for-outlook"></a>Outlook のみの拡張点

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

この拡張点により、メールの閲覧ビューのコマンド サーフェスにボタンが配置されます。Outlook デスクトップでは、これはリボンに表示されます。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### <a name="officetab-example"></a>OfficeTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="Contoso.TabCustom2">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface

この拡張点は、メールの新規作成フォームを使用してアドイン用のリボンにボタンを配置します。 

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### <a name="officetab-example"></a>OfficeTab の例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="Contoso.TabCustom3">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

この拡張点は、会議の開催者に表示されるフォームのリボンにボタンを配置します。 

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### <a name="officetab-example"></a>OfficeTab の例

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="Contoso.TabCustom4">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

この拡張点は、会議の出席者に表示されるフォームのリボンにボタンを配置します。 

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

#### <a name="officetab-example"></a>OfficeTab の例

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。

> [!IMPORTANT]
> メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  コマンドを既定のリボン タブに追加します。  |
|  [CustomTab](customtab.md) |  コマンドをカスタム リボン タブに追加します。  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface

この拡張点により、モバイル フォーム ファクターのメールの閲覧ビューのコマンド領域にボタンが配置されます。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [Group](group.md) |  コマンド領域にボタンのグループを追加します。  |

この種類の **ExtensionPoint** 要素には子要素を 1 つだけ含めることができます (**Group** 要素)。

この拡張点に含まれる **Control** 要素の **xsi:type** 属性を `MobileButton` に設定する必要があります。

#### <a name="example"></a>例

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

この拡張ポイントは、モバイル フォーム ファクターの予定のコマンド 画面にモードに適したトグルを設定します。 会議の開催者は、オンライン会議を作成できます。 その後、出席者はオンライン会議に参加できます。 このシナリオの詳細については、「オンライン会議プロバイダー Outlookモバイル アドインを作成する[」をご覧](../../outlook/online-meeting.md)ください。

> [!NOTE]
> この拡張ポイントは、Android と iOS でのみサポートされ、サブスクリプションMicrosoft 365されます。
>
> メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [Control](control.md) |  コマンド 画面にボタンを追加します。  |

`ExtensionPoint` この型の要素は、要素という 1 つの子要素のみを持 `Control` つ場合があります。

この `Control` 拡張ポイントに含まれる要素には、属性が `xsi:type` に設定されている必要があります `MobileButton`。

画像 `Icon` は、16 進数コードまたは `#919191` 他の色形式で同等の値を使用してグレー [スケールに設定する必要があります](https://convertingcolors.com/hex-color-919191.html)。

#### <a name="example"></a>例

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent"></a>LaunchEvent

この拡張ポイントを使用すると、デスクトップ フォーム ファクターでサポートされているイベントに基づいてアドインをアクティブ化できます。 このシナリオの詳細と、サポートされているイベントの完全な一覧については、「イベント ベースのアクティブ化用に Outlookアドインを構成する」[を参照](../../outlook/autolaunch.md)してください。

> [!IMPORTANT]
> メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  イベント ベース [のアクティブ化の LaunchEvent](launchevent.md) の一覧。  |
| [SourceLocation](sourcelocation.md) |  ソース JavaScript ファイルの場所。  |

#### <a name="example"></a>例

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a>Events

この拡張点は、指定したイベントのイベント ハンドラーを追加します。 この拡張ポイントの使用の詳細については、「[On-send feature for Outlookアドイン」を参照してください](../../outlook/outlook-on-send-addins.md)。

> [!IMPORTANT]
> メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。

| 要素 | 説明  |
|:-----|:-----|
|  [Event](event.md) |  イベントとイベント ハンドラーの関数を指定します。  |

#### <a name="itemsend-event-example"></a>ItemSend イベントの例

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a>DetectedEntity

この拡張点は、指定したエンティティの種類に対するコンテキスト アドインのアクティブ化を追加します。

> [!IMPORTANT]
> メールボックスイベント [とアイテム](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) イベント [の](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 登録は、この拡張ポイントでは使用できません。

これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

> [!NOTE]
> この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。

|  要素 |  説明  |
|:-----|:-----|
|  [Label](#label) |  アドインのコンテキスト ウィンドウのラベルを指定します。  |
|  [SourceLocation](sourcelocation.md) |  コンテキスト ウィンドウの URL を指定します。  |
|  [Rule](rule.md) |  アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。  |

#### <a name="label"></a>Label

必ず指定します。 グループのラベルです。 **resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

#### <a name="highlight-requirements"></a>強調表示の要件

ユーザーは、強調表示されたエンティティに対話型の操作を実行する方法でのみコンテキスト アドインを有効化できます。開発者は、`ItemHasKnownEntity` および `ItemHasRegularExpressionMatch` のルールの種類に対応する `Rule` 要素の `Highlight` 属性を使用して、強調表示にするエンティティを制御します。

ただし、注意する必要のある制限があります。これらの制限は、ユーザーにアドインをアクティブ化する方法を提供するために、適用可能なメッセージや予定で強調表示されたエンティティが常に存在するようにするために実施されます。

- `EmailAddress` および `Url` のエンティティの種類は、強調表示できません。そのため、アドインをアクティブ化するためには使用できません。
- 単一のルールを使用する場合、`Highlight` は `all` に設定されている必要があります。
- 複数のルールを組み合わせるために `Mode="AND"` で `RuleCollection` のルールの種類を使用する場合は、少なくとも 1 つのルールの `Highlight` が `all` に設定されている必要があります。
- 複数のルールを組み合わせるために `Mode="OR"` で `RuleCollection` のルールの種類を使用する場合は、すべてのルールの `Highlight` が `all` に設定されている必要があります。

#### <a name="detectedentity-event-example"></a>DetectedEntity イベントの例

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
