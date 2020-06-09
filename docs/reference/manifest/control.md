---
title: マニフェスト ファイルの Control 要素
description: アクションを実行したり、作業ウィンドウを起動したりする JavaScript 関数を定義します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 55faa52e5020691967e65b33c2d975535405b2a4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612320"
---
# <a name="control-element"></a>Control 要素

アクションを実行したり、作業ウィンドウを起動する JavaScript 関数を定義します。**Control** 要素は、[ボタン] または [メニュー] オプションのどちらかになります。少なくとも 1 つの **Control** に 1 つの [Group](group.md) 要素を含む必要があります。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|**xsi:type**|はい|定義されているコントロールの型。`Button`、`Menu`、または `MobileButton` です。 |
|**id**|いいえ|コントロール要素の ID です。最大で 125 文字です。|

> [!NOTE]
> **xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。 これは、[MobileFormFactor](mobileformfactor.md) 要素内に含まれる **Control** 要素にのみ当てはまります。

## <a name="button-control"></a>ボタン コントロール

ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。 

### <a name="child-elements"></a>子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Label**     | はい |  ボタンのテキストです。 **Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。        |
|  **ToolTip**    |いいえ|ボタンのヒントです。 **Resid**属性は、 **String**要素の**id**属性の値に設定する必要があります。 **String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|        
|  [Supertip](supertip.md)  | はい |  このボタンのヒントです。    |
|  [Icon](icon.md)      | はい |  ボタンの画像。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。  |
|  [Enabled](enabled.md)    | いいえ |  アドインを起動するときにコントロールを有効にするかどうかを指定します。  |

### <a name="executefunction-button-example"></a>ExecuteFunction ボタンの例

次の例では、アドインが起動するとボタンが無効になります。 プログラムを使用して有効にすることができます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

```xml
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
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a>ShowTaskpane ボタンの例

```xml
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
```

## <a name="menu-dropdown-button-controls"></a>メニュー (ドロップダウン ボタン) コントロール

メニューは、静的なオプションの一覧を定義します。各メニュー項目は、関数を実行したり、作業ウィンドウを表示したりします。サブメニューはサポートされません。 

[**PrimaryCommandSurface**] または [**ContextMenu**] [の拡張点](extensionpoint.md)が使用されている場合、メニュー コントロールによって以下が定義されます。

- ルートレベルのメニュー項目。

- サブメニュー項目のリスト。

**PrimaryCommandSurface** と共に使用すると、ルートのメニュー項目がリボンのボタンとして表示されます。ボタンを選択すると、サブメニューがドロップダウン リストとして表示されます。**ContextMenu** と共に使用すると、サブメニューのあるメニュー項目がコンテキスト メニューに挿入されます。どちらの場合も、各サブメニュー項目は JavaScript 関数を実行するか、作業ウィンドウを表示することができます。現時点では、サブメニューの 1 つのレベルのみがサポートされます。

次の例では、2 つのサブメニュー項目を持つメニュー項目を定義する方法を示します。最初のサブメニュー項目は作業ウィンドウを示し、2 番目のサブメニュー項目は JavaScript 関数を実行します。

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

### <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Label**     | はい |  ボタンのテキストです。 **Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。      |
|  **ToolTip**    |いいえ|ボタンのヒントです。 **Resid**属性は、 **String**要素の**id**属性の値に設定する必要があります。 **String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|        
|  [Supertip](supertip.md)  | はい |  このボタンのヒント。    |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  **Items**     | はい |  メニュー内で表示するボタンのコレクションです。 各サブメニュー項目の **Item** 要素を含みます。 各 **Item** 要素は、[ボタン コントロール](#button-control)の子要素を含みます。|

### <a name="menu-control-examples"></a>メニュー コントロールの例

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

```xml
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
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a>MobileButton コントロール

モバイル ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各モバイル ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。

**xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

### <a name="child-elements"></a>子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Label**     | はい |  ボタンのテキストです。 **Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。        |
|  [Icon](icon.md)      | はい |  ボタンの画像。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。  |

### <a name="executefunction-mobile-button-example"></a>ExecuteFunction モバイル ボタンの例

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a>ShowTaskpane モバイル ボタンの例

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
