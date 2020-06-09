---
title: マニフェスト ファイルの OfficeMenu 要素
description: Office のコンテキストメニューに追加するコントロールのコレクションを定義するのは、OfficeMenu 要素です。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f5aac4e3454e1aa18021c10bfb2f06df90805980
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611520"
---
# <a name="officemenu-element"></a><span data-ttu-id="a5b10-103">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="a5b10-103">OfficeMenu element</span></span>

<span data-ttu-id="a5b10-p101">Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="a5b10-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="a5b10-106">属性</span><span class="sxs-lookup"><span data-stu-id="a5b10-106">Attributes</span></span>

| <span data-ttu-id="a5b10-107">属性</span><span class="sxs-lookup"><span data-stu-id="a5b10-107">Attribute</span></span>            | <span data-ttu-id="a5b10-108">必須</span><span class="sxs-lookup"><span data-stu-id="a5b10-108">Required</span></span> | <span data-ttu-id="a5b10-109">説明</span><span class="sxs-lookup"><span data-stu-id="a5b10-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="a5b10-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a5b10-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="a5b10-111">はい</span><span class="sxs-lookup"><span data-stu-id="a5b10-111">Yes</span></span>      | <span data-ttu-id="a5b10-112">定義する OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="a5b10-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="a5b10-113">子要素</span><span class="sxs-lookup"><span data-stu-id="a5b10-113">Child elements</span></span>

|  <span data-ttu-id="a5b10-114">要素</span><span class="sxs-lookup"><span data-stu-id="a5b10-114">Element</span></span> |  <span data-ttu-id="a5b10-115">必須</span><span class="sxs-lookup"><span data-stu-id="a5b10-115">Required</span></span>  |  <span data-ttu-id="a5b10-116">説明</span><span class="sxs-lookup"><span data-stu-id="a5b10-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a5b10-117">Control</span><span class="sxs-lookup"><span data-stu-id="a5b10-117">Control</span></span>](#control)    | <span data-ttu-id="a5b10-118">はい</span><span class="sxs-lookup"><span data-stu-id="a5b10-118">Yes</span></span> |  <span data-ttu-id="a5b10-119">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="a5b10-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="a5b10-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a5b10-120">xsi:type</span></span>

<span data-ttu-id="a5b10-121">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="a5b10-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="a5b10-p102">`ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="a5b10-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="a5b10-p103">`ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="a5b10-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="a5b10-126">コントロール</span><span class="sxs-lookup"><span data-stu-id="a5b10-126">Control</span></span>

<span data-ttu-id="a5b10-127">各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="a5b10-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="a5b10-128">例</span><span class="sxs-lookup"><span data-stu-id="a5b10-128">Example</span></span>

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
