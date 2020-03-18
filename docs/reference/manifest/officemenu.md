---
title: マニフェスト ファイルの OfficeMenu 要素
description: Office のコンテキストメニューに追加するコントロールのコレクションを定義するのは、OfficeMenu 要素です。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 89503533f7310898a420eb805d5fd66f096ad5f2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718049"
---
# <a name="officemenu-element"></a><span data-ttu-id="cf5ff-103">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="cf5ff-103">OfficeMenu element</span></span>

<span data-ttu-id="cf5ff-p101">Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="cf5ff-106">属性</span><span class="sxs-lookup"><span data-stu-id="cf5ff-106">Attributes</span></span>

| <span data-ttu-id="cf5ff-107">属性</span><span class="sxs-lookup"><span data-stu-id="cf5ff-107">Attribute</span></span>            | <span data-ttu-id="cf5ff-108">必須</span><span class="sxs-lookup"><span data-stu-id="cf5ff-108">Required</span></span> | <span data-ttu-id="cf5ff-109">説明</span><span class="sxs-lookup"><span data-stu-id="cf5ff-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="cf5ff-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf5ff-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="cf5ff-111">はい</span><span class="sxs-lookup"><span data-stu-id="cf5ff-111">Yes</span></span>      | <span data-ttu-id="cf5ff-112">定義する OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="cf5ff-113">子要素</span><span class="sxs-lookup"><span data-stu-id="cf5ff-113">Child elements</span></span>

|  <span data-ttu-id="cf5ff-114">要素</span><span class="sxs-lookup"><span data-stu-id="cf5ff-114">Element</span></span> |  <span data-ttu-id="cf5ff-115">必須</span><span class="sxs-lookup"><span data-stu-id="cf5ff-115">Required</span></span>  |  <span data-ttu-id="cf5ff-116">説明</span><span class="sxs-lookup"><span data-stu-id="cf5ff-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cf5ff-117">Control</span><span class="sxs-lookup"><span data-stu-id="cf5ff-117">Control</span></span>](#control)    | <span data-ttu-id="cf5ff-118">はい</span><span class="sxs-lookup"><span data-stu-id="cf5ff-118">Yes</span></span> |  <span data-ttu-id="cf5ff-119">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="cf5ff-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf5ff-120">xsi:type</span></span>

<span data-ttu-id="cf5ff-121">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="cf5ff-p102">`ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="cf5ff-p103">`ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="cf5ff-126">コントロール</span><span class="sxs-lookup"><span data-stu-id="cf5ff-126">Control</span></span>

<span data-ttu-id="cf5ff-127">各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="cf5ff-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="cf5ff-128">例</span><span class="sxs-lookup"><span data-stu-id="cf5ff-128">Example</span></span>

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
