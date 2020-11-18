---
title: マニフェストファイルの Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087946"
---
# <a name="group-element"></a>Group 要素

タブ内の UI コントロールのグループを定義します。カスタムタブでは、アドインは複数のグループを作成できます。 アドインは、カスタム タブ 1 つに制限されています。

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
|  [Control](#control)    | いいえ |  Control オブジェクトを表します。 0個以上の値を指定できます。  |
|  [Officeecontrol](#officecontrol)  | いいえ | 組み込みの Office コントロールの1つを表します。 0個以上の値を指定できます。 |

### <a name="label"></a>Label

必ず指定します。 グループのラベルです。 **Resid** 属性は、 [Resources](resources.md)要素の Short **strings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="icon"></a>Icon

必ず指定します。 タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。

### <a name="control"></a>コントロール

省略可能。ただし、存在しない場合は、少なくとも1つの **Officeecontrol** が必要です。 サポートされているコントロールの種類の詳細については、 [Control](control.md) 要素を参照してください。 マニフェストでは、 **Control** と are **econtrol** の順序は相互に置き換え可能で、複数の要素がある場合は混在させることができますが、すべてが **Icon** 要素の下になければなりません。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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

### <a name="officecontrol"></a>Officeecontrol

省略可能。ただし、存在しない場合は、少なくとも1つの **コントロール** が必要です。 1つ以上の組み込みの Office コントロールを要素を含むグループに含め `<OfficeControl>` ます。 属性は、 `id` 組み込みの Office コントロールの ID を指定します。 コントロールの ID を検索するには、「 [コントロールおよびコントロールグループの id を検索](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)する」を参照してください。 マニフェストでは、 **Control** と are **econtrol** の順序は相互に置き換え可能で、複数の要素がある場合は混在させることができますが、すべてが **Icon** 要素の下になければなりません。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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
