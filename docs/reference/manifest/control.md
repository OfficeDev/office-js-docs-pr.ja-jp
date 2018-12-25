---
title: マニフェスト ファイルの Control 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: e5d8574e322c21e768fb9f66fe9bbb0c12a34ed4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433937"
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
|  **Label**     | はい |  ボタンのテキストです。**resid** 属性には、**Resources** 属性の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。        |
|  **ToolTip**  |いいえ|ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|        
|  [Supertip](supertip.md)  | はい |  このボタンのヒントです。    |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  [Action](action.md)    | はい |  実行するアクションを指定します。  |

### <a name="executefunction-button-example"></a>ExecuteFunction ボタンの例

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

- ルートレベルのメニュー項目です。

- サブメニュー項目のリスト。

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.

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
|  **Label**     | はい |  ボタンのテキストです。**resid** 属性は、[Resources](resources.md) 要素の **ShortStrings** 要素にある **String** 要素の **id** 属性の値に設定する必要があります。      |
|  **ToolTip**  |いいえ|ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|        
|  [Supertip](supertip.md)  | はい |  このボタンのヒント。    |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
|  **Items**     | はい |  メニュー内で表示するボタンのコレクションです。各サブメニュー項目の **Item** 要素を含みます。各 **Item** 要素は、[ボタン コントロール](#button-control)の子要素を含みます。|

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
|  **Label**     | はい |  ボタンのテキストです。**resid** 属性には、**Resources** 属性の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。        |
|  [Icon](icon.md)      | はい |  ボタンの画像です。         |
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