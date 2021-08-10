---
title: マニフェスト ファイル内の OverriddenByRibbonApi 要素
description: カスタム タブ、グループ、コントロール、またはメニュー アイテムがカスタム コンテキスト タブの一部である場合に表示してはならないことを指定する方法について説明します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 0f314761f686ca559caea4e04ec5d5a66fab9618ea21a221a6cf2affde897578
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092212"
---
# <a name="overriddenbyribbonapi-element"></a>OverriddenByRibbonApi 要素

カスタム コンテキスト タブをリボンにインストールする API [](control.md#menu-dropdown-button-controls) [(Office.ribbon.requestCreateControls)](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)をサポートするアプリケーションとプラットフォームの組み合わせで、CustomTab、Group、Button[コントロール、](control.md#button-control)メニュー コントロール、またはメニュー項目を非表示にするかどうかを指定します。 [](customtab.md) [](group.md)

省略すると、既定値は `false` . 使用する場合は、親要素の *最初* の子要素である必要があります。

> [!NOTE]
> この要素の詳細については、「カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する」 [を参照してください](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

この要素の目的は、カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するフォールバック エクスペリエンスをアドインに作成します。 重要な戦略は、カスタム コンテキスト タブの一部またはすべてのグループとコントロールを 1 つ以上のカスタム コア タブ (つまり、コンテキストに依存しないカスタム タブ) *に複製する* 方法です。 次に、カスタム コンテキスト タブがサポートされていないときにこれらのグループとコントロールが表示されますが、カスタム コンテキスト タブがサポートされている場合は表示されない場合は `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **、CustomTab、Group、Control、** またはmenu  **Item** 要素の最初の子要素として追加します。 その効果は次のとおりです。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインが実行されている場合、重複したタブ、グループ、およびコントロールはリボンに表示されません。 代わりに、アドインがメソッドを呼び出す際に、カスタム コンテキスト タブがインストール `requestCreateControls` されます。
- カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインを実行すると、複製されたタブ、グループ、およびコントロールがリボンに表示されます。

## <a name="examples"></a>例

### <a name="overriding-an-entire-tab"></a>タブ全体のオーバーライド

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a>グループのオーバーライド

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a>コントロールのオーバーライド

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a>メニュー項目の上書き


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
