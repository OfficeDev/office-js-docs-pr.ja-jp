---
title: マニフェスト ファイルの種類 Menu のコントロール要素
description: アイテムがアクションを実行したり、作業ウィンドウを起動したりできるメニューを定義します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7287b8e2cdf2378140ef50a41306820a0fd4002f
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467921"
---
# <a name="control-element-of-type-menu"></a>Menu 型のコントロール要素

メニューは、オプションの一覧を定義します。 各メニュー項目は、関数を実行したり、作業ウィンドウを表示したりします。

> [!NOTE]
> この記事では、要素の属性に関する重要な情報を含む基本的な [Control](control.md) リファレンス記事に精通している必要があります。

メニュー コントロールは、次の項目を定義します。

- ルート レベルのメニュー コントロール。
- メニュー項目の一覧。

**PrimaryCommandSurface** [](extensionpoint.md)拡張ポイントと一緒に使用すると、ルート メニュー項目がリボンのボタンとして表示されます。 ボタンを選択すると、メニューがドロップダウン リストとして表示されます。 サブメニューはサポートされません。

**ContextMenu 拡張ポイントと**[一](extensionpoint.md)緒に使用すると、コンテキスト メニューにルート メニュー項目が表示されます。 ルート アイテムが選択されている場合、メニュー項目はサブメニューとして表示されます。 サブメニューは 1 つのレベルしかサポートされませんので、どのアイテムもサブメニューになじめはありません。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)     | はい |  メニューのテキストです。 |
|  **ToolTip**    |いいえ|メニューのツールヒント。 **resid 属性** は 32 文字以内で、String 要素の **id** 属性の値に設定する **必要** があります。 **String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。|
|  [Supertip](supertip.md)  | はい |  このメニューのスーパーヒント。    |
|  [Icon](icon.md)      | はい |  メニューのの画像。         |
|  **Items**     | はい |  メニュー内に表示するアイテムのコレクション。 各アイテムの **Item** 要素を格納します。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにメニューを表示するかどうかを指定します。 使用する場合は、最初の子 *要素である* 必要があります。 |

### <a name="label"></a>ラベル

メニュー名のテキストを、32 文字以内で指定できる唯一の属性を使用して指定し、Resources 要素の **ShortStrings** 子の **String** 要素の **id** 属性の値に設定する [必要](resources.md)があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="examples"></a>例

次の例では、メニューには 2 つの項目があります。 1 つ目は作業ウィンドウを表示します。 2 つ目は関数を実行します。 コンテキスト タブをサポート *する* プラットフォームでアドインが実行されている場合は、メニューが表示されなく設定されています。 詳細については、「カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する [」を参照してください](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
    <Item id="ShowMainTaskPane">
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
    <Item id="GetData">
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
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

次の例では、コンテキスト タブをサポートするプラットフォームでアドインが実行されている場合、メニューの 2 番目のアイテムが表示されない構成です。 詳細については、「カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する [」を参照してください](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
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
    <Item id="ShowMainTaskPane">
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
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
