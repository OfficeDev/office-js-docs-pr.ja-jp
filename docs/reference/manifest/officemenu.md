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
# <a name="officemenu-element"></a><span data-ttu-id="217b8-102">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="217b8-102">OfficeMenu element</span></span>

<span data-ttu-id="217b8-p101">Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="217b8-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="217b8-105">属性</span><span class="sxs-lookup"><span data-stu-id="217b8-105">Attributes</span></span>

| <span data-ttu-id="217b8-106">属性</span><span class="sxs-lookup"><span data-stu-id="217b8-106">Attribute</span></span>            | <span data-ttu-id="217b8-107">必須</span><span class="sxs-lookup"><span data-stu-id="217b8-107">Required</span></span> | <span data-ttu-id="217b8-108">説明</span><span class="sxs-lookup"><span data-stu-id="217b8-108">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="217b8-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="217b8-109">xsi:type</span></span>](#xsitype) | <span data-ttu-id="217b8-110">はい</span><span class="sxs-lookup"><span data-stu-id="217b8-110">Yes</span></span>      | <span data-ttu-id="217b8-111">定義する OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="217b8-111">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="217b8-112">子要素</span><span class="sxs-lookup"><span data-stu-id="217b8-112">Child elements</span></span>

|  <span data-ttu-id="217b8-113">要素</span><span class="sxs-lookup"><span data-stu-id="217b8-113">Element</span></span> |  <span data-ttu-id="217b8-114">必須</span><span class="sxs-lookup"><span data-stu-id="217b8-114">Required</span></span>  |  <span data-ttu-id="217b8-115">説明</span><span class="sxs-lookup"><span data-stu-id="217b8-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="217b8-116">Control</span><span class="sxs-lookup"><span data-stu-id="217b8-116">Control</span></span>](#control)    | <span data-ttu-id="217b8-117">はい</span><span class="sxs-lookup"><span data-stu-id="217b8-117">Yes</span></span> |  <span data-ttu-id="217b8-118">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="217b8-118">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="217b8-119">xsi:type</span><span class="sxs-lookup"><span data-stu-id="217b8-119">xsi:type</span></span>

<span data-ttu-id="217b8-120">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="217b8-120">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="217b8-p102">`ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="217b8-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="217b8-p103">`ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="217b8-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="217b8-125">コントロール</span><span class="sxs-lookup"><span data-stu-id="217b8-125">Control</span></span>

<span data-ttu-id="217b8-126">各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="217b8-126">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="217b8-127">例</span><span class="sxs-lookup"><span data-stu-id="217b8-127">Example</span></span>

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
