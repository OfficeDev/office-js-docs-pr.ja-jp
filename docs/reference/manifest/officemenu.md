---
title: マニフェスト ファイルの OfficeMenu 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 20d020b8ab826049ef0271cbdb8d51201ee88ab4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452019"
---
# <a name="officemenu-element"></a><span data-ttu-id="4722d-102">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="4722d-102">OfficeMenu element</span></span>

<span data-ttu-id="4722d-p101">Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="4722d-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="4722d-105">属性</span><span class="sxs-lookup"><span data-stu-id="4722d-105">Attributes</span></span>

| <span data-ttu-id="4722d-106">属性</span><span class="sxs-lookup"><span data-stu-id="4722d-106">Attribute</span></span>            | <span data-ttu-id="4722d-107">必須</span><span class="sxs-lookup"><span data-stu-id="4722d-107">Required</span></span> | <span data-ttu-id="4722d-108">説明</span><span class="sxs-lookup"><span data-stu-id="4722d-108">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="4722d-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4722d-109">xsi:type</span></span>](#xsitype) | <span data-ttu-id="4722d-110">はい</span><span class="sxs-lookup"><span data-stu-id="4722d-110">Yes</span></span>      | <span data-ttu-id="4722d-111">定義する OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="4722d-111">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="4722d-112">子要素</span><span class="sxs-lookup"><span data-stu-id="4722d-112">Child elements</span></span>

|  <span data-ttu-id="4722d-113">要素</span><span class="sxs-lookup"><span data-stu-id="4722d-113">Element</span></span> |  <span data-ttu-id="4722d-114">必須</span><span class="sxs-lookup"><span data-stu-id="4722d-114">Required</span></span>  |  <span data-ttu-id="4722d-115">説明</span><span class="sxs-lookup"><span data-stu-id="4722d-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4722d-116">Control</span><span class="sxs-lookup"><span data-stu-id="4722d-116">Control</span></span>](#control)    | <span data-ttu-id="4722d-117">はい</span><span class="sxs-lookup"><span data-stu-id="4722d-117">Yes</span></span> |  <span data-ttu-id="4722d-118">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="4722d-118">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="4722d-119">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4722d-119">xsi:type</span></span>

<span data-ttu-id="4722d-120">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="4722d-120">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="4722d-p102">`ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="4722d-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="4722d-p103">`ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="4722d-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="4722d-125">コントロール</span><span class="sxs-lookup"><span data-stu-id="4722d-125">Control</span></span>

<span data-ttu-id="4722d-126">各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="4722d-126">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="4722d-127">例</span><span class="sxs-lookup"><span data-stu-id="4722d-127">Example</span></span>

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
