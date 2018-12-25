---
title: マニフェスト ファイルの CustomTab 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 7d609ad216ba5e8e7358bbc741f7b6c992bc97e2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433608"
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
|  [Control](control.md)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

### <a name="group"></a>Group

必須です。[Group 要素](group.md)を参照してください。

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