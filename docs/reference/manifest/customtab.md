---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8db29f166b5a5238a7ecf121ba5e5adca66ebe94
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718427"
---
# <a name="customtab-element"></a>CustomTab 要素

リボン上で、アドイン コマンドに使用するタブとグループを指定します。 これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。

カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。

**Id**属性はマニフェスト内で一意である必要があります。

> [!IMPORTANT]
> Mac 上の Outlook では`CustomTab` 、要素を使用できないため、代わりに[[officetab タブ](officetab.md)を使用する必要があります。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | はい |  コマンドのグループを定義します。  |
|  [Label](#label-tab)      | はい |  CustomTab または Group のラベル。  |

### <a name="group"></a>Group

必須です。 [Group 要素](group.md)を参照してください。

### <a name="label-tab"></a>Label (タブ)

必須です。 ユーザー設定のタブのラベルを示します。**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。


## <a name="customtab-example"></a>CustomTab の例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
