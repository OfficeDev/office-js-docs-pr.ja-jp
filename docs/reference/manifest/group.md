---
title: マニフェスト ファイルの Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173963"
---
# <a name="group-element"></a>Group 要素

タブ内の UI コントロールのグループを定義します。カスタム タブでは、アドインは複数のグループを作成できます。 アドインは、カスタム タブ 1 つに制限されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  はい  | グループの一意の ID。|

### <a name="id-attribute"></a>id 属性

必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)      | はい |  CustomTab またはグループのラベル。  |
|  [Icon](icon.md)      | はい |  グループのイメージ。  |
|  [Control](#control)    | いいえ |  Control オブジェクトを表します。 0 以上を指定できます。  |
|  [OfficeControl](#officecontrol)  | いいえ | 組み込みのコントロールコントロールの 1 つOfficeします。 0 以上を指定できます。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにグループを表示するかどうかを指定します。  |

### <a name="label"></a>Label

必ず指定します。 グループのラベルです。 **resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="icon"></a>Icon

必ず指定します。 タブに多くのグループが含まれている場合、プログラム ウィンドウのサイズが変更された場合、指定した画像が代わりに表示される可能性があります。

### <a name="control"></a>コントロール

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeControl が必要です**。 サポートされるコントロールの種類の詳細については [、Control](control.md) 要素を参照してください。 マニフェスト内 **のコントロール** と **OfficeControl** の順序は同じであり、複数の要素がある場合は、これらの順序が異なる可能性がありますが、すべてが **Icon** 要素の下にある必要があります。

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

省略可能ですが、存在しない場合は、少なくとも 1 つのコントロールが必要 **です**。 1 つ以上の組み込みのOffice要素を含むグループ内のコントロールを含 `<OfficeControl>` める。 この `id` 属性は、組み込みのコントロールコントロールの ID Officeします。 コントロールの ID を検索するには、「コントロールとコントロール グループの [ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内 **のコントロール** と **OfficeControl** の順序は同じであり、複数の要素がある場合は、これらの順序が異なる可能性がありますが、すべてが **Icon** 要素の下にある必要があります。

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

省略可能 (ブール値)。 実行時にリボンにカスタムコンテキスト タブをインストールする API をサポートするアプリケーションとプラットフォームの組み合わせでグループを非表示にするかどうかを指定します。 既定値 (存在しない場合) は次の値です `false` 。 使用する場合 **、OverriddenByRibbonApi は** Group の最初 *の* 子である必要 **があります**。 詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
