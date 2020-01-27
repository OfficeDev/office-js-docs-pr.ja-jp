---
title: マニフェスト ファイルの CustomTab 要素
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554014"
---
# <a name="customtab-element"></a>CustomTab 要素

リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。

カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。

**id** 属性はマニフェスト内で一意でなければなりません。

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

必須。カスタム タブのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。


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
