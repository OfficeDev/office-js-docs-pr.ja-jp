---
title: マニフェスト ファイルの ExtensionPoint 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 21def2de7168ff06c1eda26add07c33d366ec296
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433965"
---
# <a name="extensionpoint-element"></a>ExtensionPoint 要素

 Office UI でアドインが機能を公開する場所を定義します。 **ExtensionPoint** 要素は、[AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md)、[MobileFormFactor](mobileformfactor.md) の子要素です。 

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  はい  | 定義される拡張点の種類。|

## <a name="extension-points-for-excel-only"></a>Excel のみの拡張点

- **CustomFunctions** - Excel 向けの JavaScript で記述されたカスタム関数。

[この XML コード サンプル](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml)は、**CustomFunctions** 属性の値を持つ **ExtensionPoint** 要素を使用する方法と、使用する子要素を示しています。

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Word、Excel、PowerPoint、OneNote アドイン コマンドの拡張点

- **PrimaryCommandSurface** - Office のリボン。
- **ContextMenu**Office UI で右クリックしたときに表示されるショートカット メニュー。

次の例は、 **PrimaryCommandSurface** と **ContextMenu** の属性値を持つ **ExtensionPoint** 要素を使用する方法と、各要素と併用する必要がある子要素を示しています。

> [!IMPORTANT] 
> ID 属性を含む要素では、一意の ID を指定してください。 会社の名前と ID を使用することをお勧めします。 たとえば、次の形式にします。 <CustomTab id="mycompanyname.mygroupname">

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
|**CustomTab**|カスタム タブをリボンに追加する場合は必須です (  **PrimaryCommandSurface** を使用)。 **CustomTab** 要素を使用する場合、 **OfficeTab** 要素は使用できません。 **id** 属性が必要です。|
|**OfficeTab**|既定の Office リボン タブを拡張する場合は必須です (**PrimaryCommandSurface** を使用)。**OfficeTab** 要素を使用する場合、**CustomTab** 要素は使用できません。詳細については、「[OfficeTab](officetab.md)」を参照してください。|
|**OfficeMenu**|既定のコンテキスト メニューにアドイン コマンドを追加する場合は必須です (**ContextMenu** を使用)。**id** 属性は以下に設定する必要があります。 <br/> Excel または Word の場合は - **ContextMenuText**。テキストが選択され、ユーザーが選択されたテキストを右クリックしたときに、コンテキスト メニューに項目が表示されます。 <br/> Excel の場合は - **ContextMenuCell**。ユーザーがスプレッドシートのセルを右クリックすると、コンテキスト メニューに項目が表示されます。|
|**Group**|タブのユーザー インターフェイスの拡張点のグループ。グループには、最大 6 個のコントロールを指定できます。 **id** 属性が必要です。id は最大 125 文字の文字列です。|
|**Label**|必須。グループのラベル。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **ShortStrings** 要素の子要素です。|
|**Icon**|必須。小さいフォーム ファクターのデバイス、または表示されるボタンが多すぎるときに使用されるグループのアイコンを指定します。 **resid** 属性は、 **Image** 要素の **id** 属性の値に設定する必要があります。 **Image** 要素は、 **Resources** 要素の子要素である **Images** 要素の子要素です。 **size** 属性は、イメージのサイズをピクセル単位で指定します。3 つのイメージのサイズ (16、32、80) が必要です。5 つのオプションのサイズ (20、24、40、48、64) もサポートされています。|
|**Tooltip**|省略可能。グループのツールヒント。 **resid** 属性は、 **String** 要素の **id** 属性の値に設定する必要があります。 **String** 要素は、 **Resources** 要素の子要素である **LongStrings** 要素の子要素です。|
|**Control**|各グループには、少なくとも 1 つのコントロールが必要です。 **Control** 要素は、**Button** または **Menu** のどちらかになります。 ボタン コントロールのドロップダウン リストを指定する場合は、**Menu** を使用します。 現在は、ボタンとメニューのみがサポートされています。 詳細については、「[Button コントロール](control.md#button-control)」および「[Menu コントロール](control.md#menu-dropdown-button-controls)」のセクションを参照してください。<br/>**注:** トラブルシューティングを簡単にするために、**Control** 要素と関連する **Resources** 子要素を一度に 1 つずつ追加することをお勧めします。|
|**スクリプト**|カスタム関数の定義と登録コードを含む JavaScript ファイルにリンクします。 Developer Preview では、この要素は使用しません。 代わりに、HTML ページはすべての JavaScript ファイルを読み込みます。|
|**Page**|カスタム関数についての HTML ページにリンクします。|

## <a name="extension-points-for-outlook"></a>Outlook のみの拡張点

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) ([DesktopFormFactor](desktopformfactor.md) でのみ使用できます。)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
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

### <a name="events"></a>Events

この拡張点は、指定したイベントのイベント ハンドラーを追加します。

> [!NOTE]
> この要素は、Office 365 の Outlook on the web でのみサポートされています。

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
> この要素は、Office 365 の Outlook on the web でのみサポートされています。

|  要素 |  説明  |
|:-----|:-----|
|  [Label](#label) |  アドインのコンテキスト ウィンドウのラベルを指定します。  |
|  [SourceLocation](sourcelocation.md) |  コンテキスト ウィンドウの URL を指定します。  |
|  [Rule](rule.md) |  アドインをアクティブ化するタイミングを決定する 1 つ以上のルールを指定します。  |

#### <a name="label"></a>Label

必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。

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
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```