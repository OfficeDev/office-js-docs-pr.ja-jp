# <a name="control-element"></a><span data-ttu-id="f2e38-101">Control 要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-101">Control element</span></span>

<span data-ttu-id="f2e38-p101">アクションを実行したり、作業ウィンドウを起動する JavaScript 関数を定義します。**Control** 要素は、[ボタン] または [メニュー] オプションのいずれかになります。\*\* Group\*\* 要素には少なくとも 1 つの [   Control](group.md) を含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="f2e38-105">属性</span><span class="sxs-lookup"><span data-stu-id="f2e38-105">Attributes</span></span>

|  <span data-ttu-id="f2e38-106">属性</span><span class="sxs-lookup"><span data-stu-id="f2e38-106">Attribute</span></span>  |  <span data-ttu-id="f2e38-107">必須</span><span class="sxs-lookup"><span data-stu-id="f2e38-107">Required</span></span>  |  <span data-ttu-id="f2e38-108">説明</span><span class="sxs-lookup"><span data-stu-id="f2e38-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="f2e38-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="f2e38-109">**xsi:type**</span></span>|<span data-ttu-id="f2e38-110">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-110">Yes</span></span>|<span data-ttu-id="f2e38-p102">定義されているコントロールの型です。`Button`、`Menu`、または `MobileButton` です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="f2e38-113">**ID**</span><span class="sxs-lookup"><span data-stu-id="f2e38-113">**id**</span></span>|<span data-ttu-id="f2e38-114">いいえ</span><span class="sxs-lookup"><span data-stu-id="f2e38-114">No</span></span>|<span data-ttu-id="f2e38-p103">コントロール要素の ID です。最大で 125 文字です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="f2e38-117">|||UNTRANSLATED_CONTENT_START|||The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="f2e38-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="f2e38-118">これは \*\* MobileFormFactor\*\* 要素内に含まれる [  Control](mobileformfactor.md) 要素にのみ当てはまります</span><span class="sxs-lookup"><span data-stu-id="f2e38-118">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="f2e38-119">ボタン コントロール</span><span class="sxs-lookup"><span data-stu-id="f2e38-119">Button control</span></span>

<span data-ttu-id="f2e38-p105">ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示します。各ボタン コントロールには、マニフェストに固有の `id` がある必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="f2e38-123">子要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-123">Child elements</span></span>
|  <span data-ttu-id="f2e38-124">要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-124">Element</span></span> |  <span data-ttu-id="f2e38-125">必須</span><span class="sxs-lookup"><span data-stu-id="f2e38-125">Required</span></span>  |  <span data-ttu-id="f2e38-126">説明</span><span class="sxs-lookup"><span data-stu-id="f2e38-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f2e38-127">**ラベル**</span><span class="sxs-lookup"><span data-stu-id="f2e38-127">**Label**</span></span>     | <span data-ttu-id="f2e38-128">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-128">Yes</span></span> |  <span data-ttu-id="f2e38-p106">|||UNTRANSLATED_CONTENT_START|||The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="f2e38-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="f2e38-131">**ヒント**</span><span class="sxs-lookup"><span data-stu-id="f2e38-131">**ToolTip**</span></span>  |<span data-ttu-id="f2e38-132">いいえ</span><span class="sxs-lookup"><span data-stu-id="f2e38-132">No</span></span>|<span data-ttu-id="f2e38-p107">ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="f2e38-136">ヒント</span><span class="sxs-lookup"><span data-stu-id="f2e38-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="f2e38-137">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-137">Yes</span></span> |  <span data-ttu-id="f2e38-138">このボタンのヒントです。</span><span class="sxs-lookup"><span data-stu-id="f2e38-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="f2e38-139">アイコン</span><span class="sxs-lookup"><span data-stu-id="f2e38-139">Icon</span></span>](icon.md)      | <span data-ttu-id="f2e38-140">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-140">Yes</span></span> |  <span data-ttu-id="f2e38-141">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="f2e38-142">アクション</span><span class="sxs-lookup"><span data-stu-id="f2e38-142">Action</span></span>](action.md)    | <span data-ttu-id="f2e38-143">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-143">Yes</span></span> |  <span data-ttu-id="f2e38-144">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="f2e38-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="f2e38-145">ExecuteFunction ボタンの例</span><span class="sxs-lookup"><span data-stu-id="f2e38-145">ExecuteFunction button example</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="f2e38-146">ShowTaskpane ボタンの例</span><span class="sxs-lookup"><span data-stu-id="f2e38-146">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="f2e38-147">メニュー (ドロップダウン ボタン) コントロール</span><span class="sxs-lookup"><span data-stu-id="f2e38-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="f2e38-p108">メニューは、静的なオプションのリストを定義します。各メニュー項目は、関数を実行したり、作業ウィンドウを表示します。サブメニューはサポートされません。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="f2e38-151">[**PrimaryCommandSurface**] または [**ContextMenu**] [の拡張点](extensionpoint.md)が使用されている場合、メニュー コントロールによって以下が定義されます。</span><span class="sxs-lookup"><span data-stu-id="f2e38-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="f2e38-152">ルートレベルのメニュー項目。</span><span class="sxs-lookup"><span data-stu-id="f2e38-152">A root-level menu item.</span></span>

- <span data-ttu-id="f2e38-153">サブメニュー項目のリスト。</span><span class="sxs-lookup"><span data-stu-id="f2e38-153">A list of submenu items.</span></span>

<span data-ttu-id="f2e38-p109">**PrimaryCommandSurface** と共に使用すると、ルートのメニュー項目がリボンのボタンとして表示されます。ボタンを選択すると、サブメニューがドロップダウン リストとして表示されます。**ContextMenu** と共に使用すると、サブメニューのあるメニュー項目がコンテキスト メニューに挿入されます。どちらの場合も、各サブメニュー項目は JavaScript 関数を実行するか、作業ウィンドウを表示することができます。現時点では、サブメニューの 1 つのレベルのみがサポートされます。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="f2e38-p110">次の例では、2 つのサブメニュー項目を持つメニュー項目を定義する方法を示します。最初のサブメニュー項目は作業ウィンドウを示し、2 番目のサブメニュー項目は JavaScript 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="f2e38-161">子要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-161">Child elements</span></span>

|  <span data-ttu-id="f2e38-162">要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-162">Element</span></span> |  <span data-ttu-id="f2e38-163">必須</span><span class="sxs-lookup"><span data-stu-id="f2e38-163">Required</span></span>  |  <span data-ttu-id="f2e38-164">説明</span><span class="sxs-lookup"><span data-stu-id="f2e38-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f2e38-165">**ラベル**</span><span class="sxs-lookup"><span data-stu-id="f2e38-165">**Label**</span></span>     | <span data-ttu-id="f2e38-166">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-166">Yes</span></span> |  <span data-ttu-id="f2e38-p111">ボタンのテキストです。**resid** 属性は、**Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="f2e38-169">**ヒント**</span><span class="sxs-lookup"><span data-stu-id="f2e38-169">**ToolTip**</span></span>  |<span data-ttu-id="f2e38-170">いいえ</span><span class="sxs-lookup"><span data-stu-id="f2e38-170">No</span></span>|<span data-ttu-id="f2e38-p112">ボタンのヒントです。**resid** 属性は、**String** 要素の **id** 属性の値に設定する必要があります。**String** 要素は、**LongStrings** 要素 ([Resources](resources.md) 要素の子要素) の子要素です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="f2e38-174">ヒント</span><span class="sxs-lookup"><span data-stu-id="f2e38-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="f2e38-175">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-175">Yes</span></span> |  <span data-ttu-id="f2e38-176">このボタンのヒント。</span><span class="sxs-lookup"><span data-stu-id="f2e38-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="f2e38-177">アイコン</span><span class="sxs-lookup"><span data-stu-id="f2e38-177">Icon</span></span>](icon.md)      | <span data-ttu-id="f2e38-178">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-178">Yes</span></span> |  <span data-ttu-id="f2e38-179">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-179">An image for the button.</span></span>         |
|  <span data-ttu-id="f2e38-180">**アイテム**</span><span class="sxs-lookup"><span data-stu-id="f2e38-180">**Items**</span></span>     | <span data-ttu-id="f2e38-181">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-181">Yes</span></span> |  <span data-ttu-id="f2e38-p113">メニュー内で表示するボタンのコレクションです。各サブメニュー項目の **Item** 要素を含みます。各 **Item** 要素は、[ボタン コントロール](#button-control)の子要素を含みます。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="f2e38-185">メニュー コントロールの例</span><span class="sxs-lookup"><span data-stu-id="f2e38-185">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="f2e38-186">MobileButton コントロール</span><span class="sxs-lookup"><span data-stu-id="f2e38-186">MobileButton control</span></span>

<span data-ttu-id="f2e38-p114">モバイル ボタンは、ユーザーが選択したときに 1 つのアクションを実行します。関数を実行するか、作業ウィンドウを表示することができます。各モバイル ボタン コントロールには、マニフェストに固有の `id` がある必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2e38-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="f2e38-p115">|||UNTRANSLATED_CONTENT_START|||The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="f2e38-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="f2e38-192">子要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-192">Child elements</span></span>
|  <span data-ttu-id="f2e38-193">要素</span><span class="sxs-lookup"><span data-stu-id="f2e38-193">Element</span></span> |  <span data-ttu-id="f2e38-194">必須</span><span class="sxs-lookup"><span data-stu-id="f2e38-194">Required</span></span>  |  <span data-ttu-id="f2e38-195">説明</span><span class="sxs-lookup"><span data-stu-id="f2e38-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f2e38-196">**ラベル**</span><span class="sxs-lookup"><span data-stu-id="f2e38-196">**Label**</span></span>     | <span data-ttu-id="f2e38-197">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-197">Yes</span></span> |  <span data-ttu-id="f2e38-p116">|||UNTRANSLATED_CONTENT_START|||The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="f2e38-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="f2e38-200">アイコン</span><span class="sxs-lookup"><span data-stu-id="f2e38-200">Icon</span></span>](icon.md)      | <span data-ttu-id="f2e38-201">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-201">Yes</span></span> |  <span data-ttu-id="f2e38-202">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="f2e38-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="f2e38-203">アクション</span><span class="sxs-lookup"><span data-stu-id="f2e38-203">Action</span></span>](action.md)    | <span data-ttu-id="f2e38-204">はい</span><span class="sxs-lookup"><span data-stu-id="f2e38-204">Yes</span></span> |  <span data-ttu-id="f2e38-205">実行するアクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="f2e38-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="f2e38-206">ExecuteFunction モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="f2e38-206">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="f2e38-207">ShowTaskpane モバイル ボタンの例</span><span class="sxs-lookup"><span data-stu-id="f2e38-207">ShowTaskpane mobile button example</span></span>

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