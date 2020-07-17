---
title: マニフェスト ファイルの ExtensionPoint 要素
description: Office UI でアドインが機能を公開する場所を定義します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 44824e0c74b35105833f1f05cdda87bc873a4427
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094457"
---
# <a name="extensionpoint-element"></a>ExtensionPoint 要素

 Office UI でアドインが機能を公開する場所を定義します。 **ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  はい  | 定義される拡張点の種類。|

## <a name="extension-points-for-excel-only"></a>Excel のみの拡張点

- **CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。

[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点

- **PrimaryCommandSurface** - Office のリボン。
- **ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。

次の例は、**PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。

> [!IMPORTANT]
> ID 属性を含む要素では、一意の ID を指定してください。会社の名前と ID を使用することをお勧めします。たとえば、次の形式にします。<CustomTab id="mycompanyname.mygroupname">

```XML
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

#### <a name="child-elements"></a>子要素
 
|**Element**|**説明**|
|:-----|:-----|
|**CustomTab**|カスタム タブをリボンに追加する必要がある場合は必須 (**PrimaryCommandSurface** を使用)。**CustomTab** 要素を使用する場合、**OfficeTab** 要素は使用できません。**id** 属性が必要です。 |
|**OfficeTab**|既定の Office アプリリボンタブ ( **Primarycommandsurface**を使用) を拡張する場合に必要です。 **Officetab**要素を使用する場合、 **customtab**要素は使用できません。 詳細については、「[OfficeTab](officetab.md)」を参照してください。|
|**OfficeMenu**|既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須 (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 <br/> Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 <br/> Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。|
|**グループ**|タブのユーザー インターフェイスの拡張点のグループ。1 つのグループに、最大 6 個のコントロールを指定できます。**id** 属性が必要です。最大 125 文字の文字列です。 |
|**Label**|必須。グループのラベル。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**ShortStrings** 要素 (**Resources** 要素の子要素) の子要素です。 |
|**Icon**|必須。小さいフォーム ファクターのデバイス、または多くのボタンが表示されるときに使用されるグループのアイコンを指定します。**resid** 属性は、**Image** 要素の **id** 属性の値に設定する必要があります。**Image** 要素は、**Images** 要素 (**Resources** 要素の子要素) の子要素です。**size** 属性は、イメージのサイズをピクセル単位で指定します。次の 3 つのイメージのサイズが必要です。16、32、および 80。次の 5 つのオプションのサイズもサポートされています。20、24、40、48、および 64。 |
|**Tooltip**|省略可能。グループのヒント。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 (**Resources** 要素の子要素) の子要素です。 |
|**Control**|各グループには、少なくとも 1 つのコントロールが必要です。 **Control**要素には、**ボタン**または**メニュー**のいずれかを指定できます。 **メニュー**を使用して、ボタンコントロールのドロップダウンリストを指定します。 現在は、ボタンとメニューのみがサポートされています。 詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。<br/>**注:** トラブルシューティングを簡単にするために、 **Control**要素と関連する**Resources**子要素を一度に1つずつ追加することをお勧めします。|
|**スクリプト**|カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。 Developer Preview では、この要素は使用しません。 代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。|
|**Page**|カスタム関数についての HTML ページにリンクします。|

## <a name="extension-points-for-outlook"></a>Outlook のみの拡張点

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface-preview)
- [LaunchEvent](#launchevent-preview)
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

この拡張点は、モジュール拡張機能用のリボンにボタンを配置します。

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
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface-preview"></a>MobileOnlineMeetingCommandSurface (プレビュー)

> [!NOTE]
> この拡張点は、Microsoft 365 サブスクリプションを使用した Android の[プレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。

この拡張点は、モバイルフォームファクターの予定に対して、適切なモードのトグルをコマンド画面に配置します。 会議の開催者は、オンライン会議を作成できます。 出席者は、後でオンライン会議に参加できます。 このシナリオの詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](../../outlook/online-meeting.md)する」の記事を参照してください。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
|  [Control](control.md) |  コマンド画面にボタンを追加します。  |

`ExtensionPoint`この型の要素には、1つの子要素 (要素) のみを含めることができます `Control` 。

`Control`この拡張点に含まれる要素の `xsi:type` 属性をに設定する必要があり `MobileButton` ます。

画像は、 `Icon` 16 進コード `#919191` または[他のカラー形式](https://convertingcolors.com/hex-color-919191.html)の同等の画像を使用したグレースケールである必要があります。

#### <a name="example"></a>例

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
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

### <a name="launchevent-preview"></a>LaunchEvent (プレビュー)

> [!NOTE]
> この拡張点は、Microsoft 365 サブスクリプションを使用する web 上の Outlook の[プレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。

この拡張点により、デスクトップフォームファクターでサポートされているイベントに基づいてアドインをアクティブにすることができます。 現時点でサポートされているイベントは、とだけです `OnNewMessageCompose` `OnNewAppointmentOrganizer` 。 このシナリオの詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成](../../outlook/autolaunch.md)する」の記事を参照してください。

#### <a name="child-elements"></a>子要素

|  要素 |  説明  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  イベントベースのライセンス認証のための[Launchevent](launchevent.md)のリスト。  |
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

この拡張点は、指定したイベントのイベント ハンドラーを追加します。 この拡張点の使用の詳細については、「 [Outlook アドインの送信時機能](../../outlook/outlook-on-send-addins.md)」を参照してください。

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

これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

> [!NOTE]
> この要素の種類は、[要件セット 1.6 以降をサポートする Outlook クライアント ](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)が利用できます。

|  要素 |  説明  |
|:-----|:-----|
|  [Label](#label) |  アドインのコンテキスト ウィンドウのラベルを指定します。  |
|  [SourceLocation](sourcelocation.md) |  コンテキスト ウィンドウの URL を指定します。  |
|  [Rule](rule.md) |  アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。  |

#### <a name="label"></a>Label

必ず指定します。 グループのラベルです。 **Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。

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
