---
title: マニフェスト ファイルの OfficeMenu 要素
description: Office のコンテキストメニューに追加するコントロールのコレクションを定義するのは、OfficeMenu 要素です。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d181e0c6f489997a149b9713bdc257f4a2baeb16
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641445"
---
# <a name="officemenu-element"></a><span data-ttu-id="fa2c3-103">OfficeMenu 要素</span><span class="sxs-lookup"><span data-stu-id="fa2c3-103">OfficeMenu element</span></span>

<span data-ttu-id="fa2c3-p101">Office のコンテキスト メニューに追加するコントロールのコレクションを定義します。 Word、Excel、PowerPoint、OneNote アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="fa2c3-106">属性</span><span class="sxs-lookup"><span data-stu-id="fa2c3-106">Attributes</span></span>

| <span data-ttu-id="fa2c3-107">属性</span><span class="sxs-lookup"><span data-stu-id="fa2c3-107">Attribute</span></span>            | <span data-ttu-id="fa2c3-108">必須</span><span class="sxs-lookup"><span data-stu-id="fa2c3-108">Required</span></span> | <span data-ttu-id="fa2c3-109">説明</span><span class="sxs-lookup"><span data-stu-id="fa2c3-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="fa2c3-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fa2c3-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="fa2c3-111">はい</span><span class="sxs-lookup"><span data-stu-id="fa2c3-111">Yes</span></span>      | <span data-ttu-id="fa2c3-112">定義する OfficeMenu の種類。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="fa2c3-113">子要素</span><span class="sxs-lookup"><span data-stu-id="fa2c3-113">Child elements</span></span>

|  <span data-ttu-id="fa2c3-114">要素</span><span class="sxs-lookup"><span data-stu-id="fa2c3-114">Element</span></span> |  <span data-ttu-id="fa2c3-115">必須</span><span class="sxs-lookup"><span data-stu-id="fa2c3-115">Required</span></span>  |  <span data-ttu-id="fa2c3-116">説明</span><span class="sxs-lookup"><span data-stu-id="fa2c3-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fa2c3-117">Control</span><span class="sxs-lookup"><span data-stu-id="fa2c3-117">Control</span></span>](#control)    | <span data-ttu-id="fa2c3-118">はい</span><span class="sxs-lookup"><span data-stu-id="fa2c3-118">Yes</span></span> |  <span data-ttu-id="fa2c3-119">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="fa2c3-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fa2c3-120">xsi:type</span></span>

<span data-ttu-id="fa2c3-121">この Office アドインを追加する Office クライアント アプリケーションの組み込みメニューを指定します。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="fa2c3-p102">`ContextMenuText` -  テキストが選ばれ、選ばれたテキストのコンテキスト メニューをユーザーが開いたときに (右クリック)、コンテキスト メニューに項目が表示されます。 Word、Excel、PowerPoint、OneNote に適用されます。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="fa2c3-p103">`ContextMenuCell` -  ユーザーがスプレッドシートのセルのコンテキスト メニューを開くと (右クリック)、コンテキスト メニューに項目が表示されます。 Excel に適用されます。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span>

## <a name="control"></a><span data-ttu-id="fa2c3-126">コントロール</span><span class="sxs-lookup"><span data-stu-id="fa2c3-126">Control</span></span>

<span data-ttu-id="fa2c3-127">各 **OfficeMenu** 要素には、1 つ以上の [メニュー](control.md#menu-dropdown-button-controls) コントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="fa2c3-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="fa2c3-128">例</span><span class="sxs-lookup"><span data-stu-id="fa2c3-128">Example</span></span>

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
