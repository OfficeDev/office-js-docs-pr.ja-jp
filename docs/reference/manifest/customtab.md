---
title: マニフェスト ファイルの CustomTab 要素
description: ''
ms.date: 04/29/2019
localization_priority: Normal
ms.openlocfilehash: 4fa7dd86736b5ab421be5653f2e256a6b84fb480
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2019
ms.locfileid: "33517395"
---
# <a name="customtab-element"></a>CustomTab 要素

リボン上で、アドイン コマンドに使用するタブとグループを指定します。これは既定のタブ (**[ホーム]**、**[メッセージ]**、または **[会議]** のいずれか) か、アドインで定義されたカスタム タブになります。

カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。

**id** 属性はマニフェスト内で一意でなければなりません。

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
