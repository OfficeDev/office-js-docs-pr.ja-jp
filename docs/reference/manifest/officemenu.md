---
title: マニフェスト ファイルの OfficeMenu 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d243612c9b78c362bed9d90dcb539b0dbacfa6f3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432488"
---
# <a name="officemenu-element"></a>OfficeMenu 要素

Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。Word、Excel、PowerPoint、OneNote アドインに適用されます。

## <a name="attributes"></a>属性

| 属性            | 必須 | 説明                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | はい      | 定義する OfficeMenu の種類。|

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Control](#control)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

## <a name="xsitype"></a>xsi:type

この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。

- `ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。Word、Excel、PowerPoint、OneNote に適用されます。
- `ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。Excel に適用されます。 

## <a name="control"></a>コントロール

各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。 

## <a name="example"></a>例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
