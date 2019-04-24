---
title: マニフェスト ファイルの Control 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d77b464fde9898ef216ef9e47c651fb5750e4453
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450640"
---
# <a name="control-element"></a><span data-ttu-id="c536f-102">Control 要素</span><span class="sxs-lookup"><span data-stu-id="c536f-102">Control element</span></span>

<span data-ttu-id="c536f-p101">アクションを実行したり、作業ウィンドウを起動する JavaScript 関数を定義します。**Control** 要素は、[ボタン] または [メニュー] オプションのどちらかになります。少なくとも 1 つの **Control** に 1 つの [Group](group.md) 要素を含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="c536f-106">属性</span><span class="sxs-lookup"><span data-stu-id="c536f-106">Attributes</span></span>

|  <span data-ttu-id="c536f-107">属性</span><span class="sxs-lookup"><span data-stu-id="c536f-107">Attribute</span></span>  |  <span data-ttu-id="c536f-108">必須</span><span class="sxs-lookup"><span data-stu-id="c536f-108">Required</span></span>  |  <span data-ttu-id="c536f-109">説明</span><span class="sxs-lookup"><span data-stu-id="c536f-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="c536f-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c536f-110">**xsi:type**</span></span>|<span data-ttu-id="c536f-111">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-111">Yes</span></span>|<span data-ttu-id="c536f-p102">定義されているコントロールの型。`Button`、`Menu`、または `MobileButton` です。</span><span class="sxs-lookup"><span data-stu-id="c536f-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="c536f-114">**id**</span><span class="sxs-lookup"><span data-stu-id="c536f-114">**id**</span></span>|<span data-ttu-id="c536f-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="c536f-115">No</span></span>|<span data-ttu-id="c536f-p103">コントロール要素の ID です。最大で 125 文字です。</span><span class="sxs-lookup"><span data-stu-id="c536f-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="c536f-118">**xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。</span><span class="sxs-lookup"><span data-stu-id="c536f-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="c536f-119">これは、[MobileFormFactor](mobileformfactor.md) 要素内に含まれる **Control** 要素にのみ当てはまります。</span><span class="sxs-lookup"><span data-stu-id="c536f-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="c536f-120">ボタン コントロール</span><span class="sxs-lookup"><span data-stu-id="c536f-120">Button control</span></span>

<span data-ttu-id="c536f-p105">ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="c536f-124">子要素</span><span class="sxs-lookup"><span data-stu-id="c536f-124">Child elements</span></span>
|  <span data-ttu-id="c536f-125">要素</span><span class="sxs-lookup"><span data-stu-id="c536f-125">Element</span></span> |  <span data-ttu-id="c536f-126">必須</span><span class="sxs-lookup"><span data-stu-id="c536f-126">Required</span></span>  |  <span data-ttu-id="c536f-127">説明</span><span class="sxs-lookup"><span data-stu-id="c536f-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c536f-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="c536f-128">**Label**</span></span>     | <span data-ttu-id="c536f-129">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-129">Yes</span></span> |  <span data-ttu-id="c536f-p106">ボタンのテキストです。 **resid** 属性には、**Resources** 属性の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="c536f-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c536f-132">**ToolTip**</span></span>  |<span data-ttu-id="c536f-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="c536f-133">No</span></span>|<span data-ttu-id="c536f-p107">ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="c536f-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="c536f-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="c536f-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c536f-138">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-138">Yes</span></span> |  <span data-ttu-id="c536f-139">このボタンのヒントです。</span><span class="sxs-lookup"><span data-stu-id="c536f-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="c536f-140">Icon</span><span class="sxs-lookup"><span data-stu-id="c536f-140">Icon</span></span>](icon.md)      | <span data-ttu-id="c536f-141">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-141">Yes</span></span> |  <span data-ttu-id="c536f-142">ボタンの画像。</span><span class="sxs-lookup"><span data-stu-id="c536f-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="c536f-143">Action</span><span class="sxs-lookup"><span data-stu-id="c536f-143">Action</span></span>](action.md)    | <span data-ttu-id="c536f-144">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-144">Yes</span></span> |  <span data-ttu-id="c536f-145">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="c536f-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="c536f-146">ExecuteFunction ボタンの例</span><span class="sxs-lookup"><span data-stu-id="c536f-146">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="c536f-147">ShowTaskpane ボタンの例</span><span class="sxs-lookup"><span data-stu-id="c536f-147">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="c536f-148">メニュー (ドロップダウン ボタン) コントロール</span><span class="sxs-lookup"><span data-stu-id="c536f-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="c536f-p108">メニューは、静的なオプションの一覧を定義します。各メニュー項目は、関数を実行したり、作業ウィンドウを表示したりします。サブメニューはサポートされません。</span><span class="sxs-lookup"><span data-stu-id="c536f-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="c536f-152">[**PrimaryCommandSurface**] または [**ContextMenu**] [の拡張点](extensionpoint.md)が使用されている場合、メニュー コントロールによって以下が定義されます。</span><span class="sxs-lookup"><span data-stu-id="c536f-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="c536f-153">ルートレベルのメニュー項目。</span><span class="sxs-lookup"><span data-stu-id="c536f-153">A root-level menu item.</span></span>

- <span data-ttu-id="c536f-154">サブメニュー項目のリスト。</span><span class="sxs-lookup"><span data-stu-id="c536f-154">A list of submenu items.</span></span>

<span data-ttu-id="c536f-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span><span class="sxs-lookup"><span data-stu-id="c536f-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="c536f-p110">次の例では、2 つのサブメニュー項目を持つメニュー項目を定義する方法を示します。最初のサブメニュー項目は作業ウィンドウを示し、2 番目のサブメニュー項目は JavaScript 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="c536f-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="c536f-162">子要素</span><span class="sxs-lookup"><span data-stu-id="c536f-162">Child elements</span></span>

|  <span data-ttu-id="c536f-163">要素</span><span class="sxs-lookup"><span data-stu-id="c536f-163">Element</span></span> |  <span data-ttu-id="c536f-164">必須</span><span class="sxs-lookup"><span data-stu-id="c536f-164">Required</span></span>  |  <span data-ttu-id="c536f-165">説明</span><span class="sxs-lookup"><span data-stu-id="c536f-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c536f-166">**Label**</span><span class="sxs-lookup"><span data-stu-id="c536f-166">**Label**</span></span>     | <span data-ttu-id="c536f-167">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-167">Yes</span></span> |  <span data-ttu-id="c536f-p111">ボタンのテキストです。**resid** 属性は、**Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="c536f-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c536f-170">**ToolTip**</span></span>  |<span data-ttu-id="c536f-171">いいえ</span><span class="sxs-lookup"><span data-stu-id="c536f-171">No</span></span>|<span data-ttu-id="c536f-p112">ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="c536f-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="c536f-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="c536f-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c536f-176">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-176">Yes</span></span> |  <span data-ttu-id="c536f-177">このボタンのヒント。</span><span class="sxs-lookup"><span data-stu-id="c536f-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="c536f-178">Icon</span><span class="sxs-lookup"><span data-stu-id="c536f-178">Icon</span></span>](icon.md)      | <span data-ttu-id="c536f-179">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-179">Yes</span></span> |  <span data-ttu-id="c536f-180">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="c536f-180">An image for the button.</span></span>         |
|  <span data-ttu-id="c536f-181">**Items**</span><span class="sxs-lookup"><span data-stu-id="c536f-181">**Items**</span></span>     | <span data-ttu-id="c536f-182">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-182">Yes</span></span> |  <span data-ttu-id="c536f-p113">メニュー内で表示するボタンのコレクションです。各サブメニュー項目の **Item** 要素を含みます。各 **Item** 要素は、[ボタン コントロール](#button-control)の子要素を含みます。</span><span class="sxs-lookup"><span data-stu-id="c536f-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="c536f-186">メニュー コントロールの例</span><span class="sxs-lookup"><span data-stu-id="c536f-186">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="c536f-187">MobileButton コントロール</span><span class="sxs-lookup"><span data-stu-id="c536f-187">MobileButton control</span></span>

<span data-ttu-id="c536f-p114">モバイル ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各モバイル ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="c536f-p115">**xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="c536f-193">子要素</span><span class="sxs-lookup"><span data-stu-id="c536f-193">Child elements</span></span>
|  <span data-ttu-id="c536f-194">要素</span><span class="sxs-lookup"><span data-stu-id="c536f-194">Element</span></span> |  <span data-ttu-id="c536f-195">必須</span><span class="sxs-lookup"><span data-stu-id="c536f-195">Required</span></span>  |  <span data-ttu-id="c536f-196">説明</span><span class="sxs-lookup"><span data-stu-id="c536f-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c536f-197">**Label**</span><span class="sxs-lookup"><span data-stu-id="c536f-197">**Label**</span></span>     | <span data-ttu-id="c536f-198">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-198">Yes</span></span> |  <span data-ttu-id="c536f-p116">ボタンのテキストです。**resid** 属性には、**Resources** 属性の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c536f-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="c536f-201">Icon</span><span class="sxs-lookup"><span data-stu-id="c536f-201">Icon</span></span>](icon.md)      | <span data-ttu-id="c536f-202">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-202">Yes</span></span> |  <span data-ttu-id="c536f-203">ボタンの画像。</span><span class="sxs-lookup"><span data-stu-id="c536f-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="c536f-204">Action</span><span class="sxs-lookup"><span data-stu-id="c536f-204">Action</span></span>](action.md)    | <span data-ttu-id="c536f-205">はい</span><span class="sxs-lookup"><span data-stu-id="c536f-205">Yes</span></span> |  <span data-ttu-id="c536f-206">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="c536f-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="c536f-207">ExecuteFunction モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="c536f-207">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="c536f-208">ShowTaskpane モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="c536f-208">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
