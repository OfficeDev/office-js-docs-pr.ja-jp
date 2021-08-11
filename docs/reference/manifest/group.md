---
title: マニフェスト ファイル内の Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 236746d6f6ae5e04612aade7e7d29564b064f384d65b6c0be582117faf6cecf6
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098750"
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
|  [Icon](icon.md)      | はい |  グループのイメージ。 このアドインではOutlookサポートされていません。 |
|  [Control](#control)    | いいえ |  Control オブジェクトを表します。 0 以上の値を指定できます。  |
|  [OfficeControl](#officecontrol)  | いいえ | 組み込みのコントロールの 1 Officeします。 0 以上の値を指定できます。 このアドインではOutlookサポートされていません。|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | いいえ |  カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにグループを表示するかどうかを指定します。 このアドインではOutlookサポートされていません。 |

### <a name="label"></a>Label

必ず指定します。 グループのラベルです。 **resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

### <a name="icon"></a>Icon

必ず指定します。 タブに多くのグループが含まれている場合、プログラム ウィンドウのサイズが変更された場合、指定したイメージが代わりに表示される場合があります。

> [!NOTE]
> この子要素は、アドインOutlookサポートされていません。

### <a name="control"></a>コントロール

省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeControl が必要です**。 サポートされるコントロールの種類の詳細については [、Control](control.md) 要素を参照してください。 マニフェスト内の **Control** と **OfficeControl** の順序は交換可能で、複数の要素がある場合は相互に混同できますが、すべてが Icon 要素の下にある **必要** があります。

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

オプションですが、存在しない場合は少なくとも 1 つの Control が必要 **です**。 1 つ以上の組み込みOffice要素を含むコントロールをグループに含 `<OfficeControl>` める。 属性 `id` は、組み込みのコントロールの ID をOfficeします。 コントロールの ID を見つけるには、「コントロールとコントロール グループの [ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 マニフェスト内の **Control** と **OfficeControl** の順序は交換可能で、複数の要素がある場合は相互に混同できますが、すべてが Icon 要素の下にある **必要** があります。

> [!NOTE]
> この子要素は、アドインOutlookサポートされていません。

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

省略可能 (ブール型)。 実行時にリボンにカスタムコンテキスト タブをインストールする API をサポートするアプリケーションとプラットフォームの組み合わせでグループを非表示にするかどうかを指定します。 既定値 (存在しない場合) は、 です `false` 。 使用する場合 **、OverriddenByRibbonApi は Group** の *最初の* 子である **必要があります**。 詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。

> [!NOTE]
> この子要素は、アドインOutlookサポートされていません。

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
