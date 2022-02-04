---
title: マニフェスト ファイル内の OverriddenByRibbonApi 要素
description: カスタム タブ、グループ、コントロール、またはメニュー アイテムがカスタム コンテキスト タブの一部である場合に表示してはならないことを指定する方法について説明します。
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="overriddenbyribbonapi-element"></a>OverriddenByRibbonApi 要素

リボンにカスタム コンテキスト タブ[](group.md)をインストール[](control.md#button-control)する API ([](control.md#menu-dropdown-button-controls)[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1))) をサポートするアプリケーションとプラットフォームの組み合わせで、グループ、ボタン コントロール、メニュー コントロール、またはメニュー項目を非表示にするかどうかを指定します。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [リボン 1.2](../requirement-sets/add-in-commands-requirement-sets.md) (Excel、PowerPoint、および Word に必須)

この要素を省略すると、既定値は .`false` 使用する場合は、親要素の *最初の子* 要素である必要があります。

> [!NOTE]
> この要素の詳細については、「カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する」 [を参照してください](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

この要素の目的は、カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するフォールバック エクスペリエンスをアドインに作成します。 重要な戦略は、カスタム コンテキスト タブの一部またはすべてのグループとコントロールを 1 つ以上のカスタム コア タブ (つまり、コンテキストに依存しないカスタム タブ) *に複製する* 方法です。  `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`次に、カスタム コンテキスト タブがサポートされていないときにこれらのグループとコントロールが表示されますが、カスタム コンテキスト タブがサポートされている場合は表示されない場合は、Group、**Control**、または menu **Item** 要素の最初の子要素として追加します。 その効果は次のとおりです。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインが実行されている場合、重複したグループとコントロールはリボンに表示されません。 代わりに、アドインがメソッドを呼び出す際に、カスタム コンテキスト タブがインストール `requestCreateControls` されます。
- カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインを実行すると、重複したグループとコントロールがリボンに表示されます。

## <a name="examples"></a>例

### <a name="overriding-a-group"></a>グループのオーバーライド

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.CustomTab.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="Contoso.MyButton1">
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
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.CustomTab2.group2">
      <Control  xsi:type="Button" id="Contoso.MyButton2">
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
  <CustomTab id="Contoso.TabCustom3">
    <Group id="Contoso.CustomTab3.group3.">
      <Control  xsi:type="Menu" id="Contoso.MyMenu">
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
