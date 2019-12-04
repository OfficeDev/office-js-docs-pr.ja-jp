---
title: マニフェストファイルの Group 要素
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 35db4829b40078e97fbfc007e2fb552e00875f9c
ms.sourcegitcommit: 164b11b1e9d2ae20b3d816092025b32a9070450f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2019
ms.locfileid: "39818728"
---
# <a name="group-element"></a>Group 要素

タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  はい  | グループの一意の ID。|

### <a name="id-attribute"></a>id 属性

必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。

## <a name="child-elements"></a>子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)      | ○ |  CustomTab またはグループのラベル。  |
|  [Icon](icon.md)      | はい |  グループのイメージ。  |
|  [Control](#control)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

### <a name="label"></a>Label 

必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。

### <a name="icon"></a>Icon

必ず指定します。 タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。

### <a name="control"></a>Control
1 つのグループに少なくとも 1 つのコントロールが必要です。 サポートされているコントロールの種類の詳細については、 [Control](control.md)要素を参照してください。

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
