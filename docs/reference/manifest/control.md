---
title: マニフェスト ファイルの Control 要素
description: ''
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: ccf7c3065db13a311825498292713b619f1cd745
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/07/2020
ms.locfileid: "42562086"
---
# <a name="control-element"></a><span data-ttu-id="3dfa2-102">Control 要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-102">Control element</span></span>

<span data-ttu-id="3dfa2-p101">アクションを実行したり、作業ウィンドウを起動する JavaScript 関数を定義します。**Control** 要素は、[ボタン] または [メニュー] オプションのどちらかになります。少なくとも 1 つの **Control** に 1 つの [Group](group.md) 要素を含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="3dfa2-106">属性</span><span class="sxs-lookup"><span data-stu-id="3dfa2-106">Attributes</span></span>

|  <span data-ttu-id="3dfa2-107">属性</span><span class="sxs-lookup"><span data-stu-id="3dfa2-107">Attribute</span></span>  |  <span data-ttu-id="3dfa2-108">必須</span><span class="sxs-lookup"><span data-stu-id="3dfa2-108">Required</span></span>  |  <span data-ttu-id="3dfa2-109">説明</span><span class="sxs-lookup"><span data-stu-id="3dfa2-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="3dfa2-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-110">**xsi:type**</span></span>|<span data-ttu-id="3dfa2-111">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-111">Yes</span></span>|<span data-ttu-id="3dfa2-p102">定義されているコントロールの型。`Button`、`Menu`、または `MobileButton` です。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="3dfa2-114">**id**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-114">**id**</span></span>|<span data-ttu-id="3dfa2-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="3dfa2-115">No</span></span>|<span data-ttu-id="3dfa2-p103">コントロール要素の ID です。最大で 125 文字です。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="3dfa2-118">**xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="3dfa2-119">これは、[MobileFormFactor](mobileformfactor.md) 要素内に含まれる **Control** 要素にのみ当てはまります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="3dfa2-120">ボタン コントロール</span><span class="sxs-lookup"><span data-stu-id="3dfa2-120">Button control</span></span>

<span data-ttu-id="3dfa2-p105">ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="3dfa2-124">子要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-124">Child elements</span></span>
|  <span data-ttu-id="3dfa2-125">要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-125">Element</span></span> |  <span data-ttu-id="3dfa2-126">必須</span><span class="sxs-lookup"><span data-stu-id="3dfa2-126">Required</span></span>  |  <span data-ttu-id="3dfa2-127">説明</span><span class="sxs-lookup"><span data-stu-id="3dfa2-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3dfa2-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-128">**Label**</span></span>     | <span data-ttu-id="3dfa2-129">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-129">Yes</span></span> |  <span data-ttu-id="3dfa2-130">ボタンのテキストです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-130">The text for the button.</span></span> <span data-ttu-id="3dfa2-131">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="3dfa2-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-132">**ToolTip**</span></span>  |<span data-ttu-id="3dfa2-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="3dfa2-133">No</span></span>|<span data-ttu-id="3dfa2-134">ボタンのヒントです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-134">The tooltip for the button.</span></span> <span data-ttu-id="3dfa2-135">**Resid**属性は、 **String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3dfa2-136">**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="3dfa2-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="3dfa2-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="3dfa2-138">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-138">Yes</span></span> |  <span data-ttu-id="3dfa2-139">このボタンのヒントです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="3dfa2-140">Icon</span><span class="sxs-lookup"><span data-stu-id="3dfa2-140">Icon</span></span>](icon.md)      | <span data-ttu-id="3dfa2-141">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-141">Yes</span></span> |  <span data-ttu-id="3dfa2-142">ボタンの画像。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="3dfa2-143">Action</span><span class="sxs-lookup"><span data-stu-id="3dfa2-143">Action</span></span>](action.md)    | <span data-ttu-id="3dfa2-144">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-144">Yes</span></span> |  <span data-ttu-id="3dfa2-145">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-145">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="3dfa2-146">Enabled</span><span class="sxs-lookup"><span data-stu-id="3dfa2-146">Enabled</span></span>](enabled.md)    | <span data-ttu-id="3dfa2-147">いいえ</span><span class="sxs-lookup"><span data-stu-id="3dfa2-147">No</span></span> |  <span data-ttu-id="3dfa2-148">アドインを起動するときにコントロールを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-148">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="3dfa2-149">ExecuteFunction ボタンの例</span><span class="sxs-lookup"><span data-stu-id="3dfa2-149">ExecuteFunction button example</span></span>

<span data-ttu-id="3dfa2-150">次の例では、アドインが起動するとボタンが無効になります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-150">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="3dfa2-151">プログラムを使用して有効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-151">It can be programmatically enabled.</span></span> <span data-ttu-id="3dfa2-152">詳細については、「[アドインコマンドを有効または無効](/office/dev/add-ins/design/disable-add-in-commands)にする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-152">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

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
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="3dfa2-153">ShowTaskpane ボタンの例</span><span class="sxs-lookup"><span data-stu-id="3dfa2-153">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="3dfa2-154">メニュー (ドロップダウン ボタン) コントロール</span><span class="sxs-lookup"><span data-stu-id="3dfa2-154">Menu (dropdown button) controls</span></span>

<span data-ttu-id="3dfa2-p109">メニューは、静的なオプションの一覧を定義します。各メニュー項目は、関数を実行したり、作業ウィンドウを表示したりします。サブメニューはサポートされません。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="3dfa2-158">[**PrimaryCommandSurface**] または [**ContextMenu**] [の拡張点](extensionpoint.md)が使用されている場合、メニュー コントロールによって以下が定義されます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-158">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="3dfa2-159">ルートレベルのメニュー項目。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-159">A root-level menu item.</span></span>

- <span data-ttu-id="3dfa2-160">サブメニュー項目のリスト。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-160">A list of submenu items.</span></span>

<span data-ttu-id="3dfa2-p110">**PrimaryCommandSurface** と共に使用すると、ルートのメニュー項目がリボンのボタンとして表示されます。ボタンを選択すると、サブメニューがドロップダウン リストとして表示されます。**ContextMenu** と共に使用すると、サブメニューのあるメニュー項目がコンテキスト メニューに挿入されます。どちらの場合も、各サブメニュー項目は JavaScript 関数を実行するか、作業ウィンドウを表示することができます。現時点では、サブメニューの 1 つのレベルのみがサポートされます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="3dfa2-p111">次の例では、2 つのサブメニュー項目を持つメニュー項目を定義する方法を示します。最初のサブメニュー項目は作業ウィンドウを示し、2 番目のサブメニュー項目は JavaScript 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="3dfa2-168">子要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-168">Child elements</span></span>

|  <span data-ttu-id="3dfa2-169">要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-169">Element</span></span> |  <span data-ttu-id="3dfa2-170">必須</span><span class="sxs-lookup"><span data-stu-id="3dfa2-170">Required</span></span>  |  <span data-ttu-id="3dfa2-171">説明</span><span class="sxs-lookup"><span data-stu-id="3dfa2-171">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3dfa2-172">**Label**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-172">**Label**</span></span>     | <span data-ttu-id="3dfa2-173">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-173">Yes</span></span> |  <span data-ttu-id="3dfa2-174">ボタンのテキストです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-174">The text for the button.</span></span> <span data-ttu-id="3dfa2-175">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-175">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="3dfa2-176">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-176">**ToolTip**</span></span>  |<span data-ttu-id="3dfa2-177">いいえ</span><span class="sxs-lookup"><span data-stu-id="3dfa2-177">No</span></span>|<span data-ttu-id="3dfa2-178">ボタンのヒントです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-178">The tooltip for the button.</span></span> <span data-ttu-id="3dfa2-179">**Resid**属性は、 **String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-179">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3dfa2-180">**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-180">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="3dfa2-181">Supertip</span><span class="sxs-lookup"><span data-stu-id="3dfa2-181">Supertip</span></span>](supertip.md)  | <span data-ttu-id="3dfa2-182">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-182">Yes</span></span> |  <span data-ttu-id="3dfa2-183">このボタンのヒント。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-183">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="3dfa2-184">Icon</span><span class="sxs-lookup"><span data-stu-id="3dfa2-184">Icon</span></span>](icon.md)      | <span data-ttu-id="3dfa2-185">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-185">Yes</span></span> |  <span data-ttu-id="3dfa2-186">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-186">An image for the button.</span></span>         |
|  <span data-ttu-id="3dfa2-187">**Items**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-187">**Items**</span></span>     | <span data-ttu-id="3dfa2-188">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-188">Yes</span></span> |  <span data-ttu-id="3dfa2-189">メニュー内で表示するボタンのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-189">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="3dfa2-190">各サブメニュー項目の **Item** 要素を含みます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-190">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="3dfa2-191">各 **Item** 要素は、[ボタン コントロール](#button-control)の子要素を含みます。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-191">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="3dfa2-192">メニュー コントロールの例</span><span class="sxs-lookup"><span data-stu-id="3dfa2-192">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="3dfa2-193">MobileButton コントロール</span><span class="sxs-lookup"><span data-stu-id="3dfa2-193">MobileButton control</span></span>

<span data-ttu-id="3dfa2-p115">モバイル ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各モバイル ボタン コントロールには、マニフェストで一意の `id` を持っている必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="3dfa2-p116">**xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="3dfa2-199">子要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-199">Child elements</span></span>
|  <span data-ttu-id="3dfa2-200">要素</span><span class="sxs-lookup"><span data-stu-id="3dfa2-200">Element</span></span> |  <span data-ttu-id="3dfa2-201">必須</span><span class="sxs-lookup"><span data-stu-id="3dfa2-201">Required</span></span>  |  <span data-ttu-id="3dfa2-202">説明</span><span class="sxs-lookup"><span data-stu-id="3dfa2-202">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3dfa2-203">**Label**</span><span class="sxs-lookup"><span data-stu-id="3dfa2-203">**Label**</span></span>     | <span data-ttu-id="3dfa2-204">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-204">Yes</span></span> |  <span data-ttu-id="3dfa2-205">ボタンのテキストです。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-205">The text for the button.</span></span> <span data-ttu-id="3dfa2-206">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-206">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="3dfa2-207">Icon</span><span class="sxs-lookup"><span data-stu-id="3dfa2-207">Icon</span></span>](icon.md)      | <span data-ttu-id="3dfa2-208">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-208">Yes</span></span> |  <span data-ttu-id="3dfa2-209">ボタンの画像。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-209">An image for the button.</span></span>         |
|  [<span data-ttu-id="3dfa2-210">Action</span><span class="sxs-lookup"><span data-stu-id="3dfa2-210">Action</span></span>](action.md)    | <span data-ttu-id="3dfa2-211">はい</span><span class="sxs-lookup"><span data-stu-id="3dfa2-211">Yes</span></span> |  <span data-ttu-id="3dfa2-212">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="3dfa2-212">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="3dfa2-213">ExecuteFunction モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="3dfa2-213">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="3dfa2-214">ShowTaskpane モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="3dfa2-214">ShowTaskpane mobile button example</span></span>

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
