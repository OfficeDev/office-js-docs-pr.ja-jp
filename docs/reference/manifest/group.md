---
title: マニフェスト ファイルの Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173963"
---
# <a name="group-element"></a><span data-ttu-id="04bb1-103">Group 要素</span><span class="sxs-lookup"><span data-stu-id="04bb1-103">Group element</span></span>

<span data-ttu-id="04bb1-104">タブ内の UI コントロールのグループを定義します。カスタム タブでは、アドインは複数のグループを作成できます。</span><span class="sxs-lookup"><span data-stu-id="04bb1-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="04bb1-105">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="04bb1-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="04bb1-106">属性</span><span class="sxs-lookup"><span data-stu-id="04bb1-106">Attributes</span></span>

|  <span data-ttu-id="04bb1-107">属性</span><span class="sxs-lookup"><span data-stu-id="04bb1-107">Attribute</span></span>  |  <span data-ttu-id="04bb1-108">必須</span><span class="sxs-lookup"><span data-stu-id="04bb1-108">Required</span></span>  |  <span data-ttu-id="04bb1-109">説明</span><span class="sxs-lookup"><span data-stu-id="04bb1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="04bb1-110">id</span><span class="sxs-lookup"><span data-stu-id="04bb1-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="04bb1-111">はい</span><span class="sxs-lookup"><span data-stu-id="04bb1-111">Yes</span></span>  | <span data-ttu-id="04bb1-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="04bb1-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="04bb1-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="04bb1-113">id attribute</span></span>

<span data-ttu-id="04bb1-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="04bb1-118">子要素</span><span class="sxs-lookup"><span data-stu-id="04bb1-118">Child elements</span></span>

|  <span data-ttu-id="04bb1-119">要素</span><span class="sxs-lookup"><span data-stu-id="04bb1-119">Element</span></span> |  <span data-ttu-id="04bb1-120">必須</span><span class="sxs-lookup"><span data-stu-id="04bb1-120">Required</span></span>  |  <span data-ttu-id="04bb1-121">説明</span><span class="sxs-lookup"><span data-stu-id="04bb1-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="04bb1-122">Label</span><span class="sxs-lookup"><span data-stu-id="04bb1-122">Label</span></span>](#label)      | <span data-ttu-id="04bb1-123">はい</span><span class="sxs-lookup"><span data-stu-id="04bb1-123">Yes</span></span> |  <span data-ttu-id="04bb1-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="04bb1-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="04bb1-125">Icon</span><span class="sxs-lookup"><span data-stu-id="04bb1-125">Icon</span></span>](icon.md)      | <span data-ttu-id="04bb1-126">はい</span><span class="sxs-lookup"><span data-stu-id="04bb1-126">Yes</span></span> |  <span data-ttu-id="04bb1-127">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="04bb1-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="04bb1-128">Control</span><span class="sxs-lookup"><span data-stu-id="04bb1-128">Control</span></span>](#control)    | <span data-ttu-id="04bb1-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="04bb1-129">No</span></span> |  <span data-ttu-id="04bb1-130">Control オブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-130">Represents a Control object.</span></span> <span data-ttu-id="04bb1-131">0 以上を指定できます。</span><span class="sxs-lookup"><span data-stu-id="04bb1-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="04bb1-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="04bb1-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="04bb1-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="04bb1-133">No</span></span> | <span data-ttu-id="04bb1-134">組み込みのコントロールコントロールの 1 つOfficeします。</span><span class="sxs-lookup"><span data-stu-id="04bb1-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="04bb1-135">0 以上を指定できます。</span><span class="sxs-lookup"><span data-stu-id="04bb1-135">Can be zero or more.</span></span> |
|  [<span data-ttu-id="04bb1-136">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="04bb1-136">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="04bb1-137">いいえ</span><span class="sxs-lookup"><span data-stu-id="04bb1-137">No</span></span> |  <span data-ttu-id="04bb1-138">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにグループを表示するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-138">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span>  |

### <a name="label"></a><span data-ttu-id="04bb1-139">Label</span><span class="sxs-lookup"><span data-stu-id="04bb1-139">Label</span></span>

<span data-ttu-id="04bb1-140">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-140">Required.</span></span> <span data-ttu-id="04bb1-141">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="04bb1-141">The label of the group.</span></span> <span data-ttu-id="04bb1-142">**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="04bb1-142">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="04bb1-143">Icon</span><span class="sxs-lookup"><span data-stu-id="04bb1-143">Icon</span></span>

<span data-ttu-id="04bb1-144">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-144">Required.</span></span> <span data-ttu-id="04bb1-145">タブに多くのグループが含まれている場合、プログラム ウィンドウのサイズが変更された場合、指定した画像が代わりに表示される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="04bb1-145">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="04bb1-146">コントロール</span><span class="sxs-lookup"><span data-stu-id="04bb1-146">Control</span></span>

<span data-ttu-id="04bb1-147">省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeControl が必要です**。</span><span class="sxs-lookup"><span data-stu-id="04bb1-147">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="04bb1-148">サポートされるコントロールの種類の詳細については [、Control](control.md) 要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="04bb1-148">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="04bb1-149">マニフェスト内 **のコントロール** と **OfficeControl** の順序は同じであり、複数の要素がある場合は、これらの順序が異なる可能性がありますが、すべてが **Icon** 要素の下にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="04bb1-149">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a><span data-ttu-id="04bb1-150">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="04bb1-150">OfficeControl</span></span>

<span data-ttu-id="04bb1-151">省略可能ですが、存在しない場合は、少なくとも 1 つのコントロールが必要 **です**。</span><span class="sxs-lookup"><span data-stu-id="04bb1-151">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="04bb1-152">1 つ以上の組み込みのOffice要素を含むグループ内のコントロールを含 `<OfficeControl>` める。</span><span class="sxs-lookup"><span data-stu-id="04bb1-152">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="04bb1-153">この `id` 属性は、組み込みのコントロールコントロールの ID Officeします。</span><span class="sxs-lookup"><span data-stu-id="04bb1-153">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="04bb1-154">コントロールの ID を検索するには、「コントロールとコントロール グループの [ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="04bb1-154">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="04bb1-155">マニフェスト内 **のコントロール** と **OfficeControl** の順序は同じであり、複数の要素がある場合は、これらの順序が異なる可能性がありますが、すべてが **Icon** 要素の下にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="04bb1-155">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="04bb1-156">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="04bb1-156">OverriddenByRibbonApi</span></span>

<span data-ttu-id="04bb1-157">省略可能 (ブール値)。</span><span class="sxs-lookup"><span data-stu-id="04bb1-157">Optional (boolean).</span></span> <span data-ttu-id="04bb1-158">実行時にリボンにカスタムコンテキスト タブをインストールする API をサポートするアプリケーションとプラットフォームの組み合わせでグループを非表示にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="04bb1-158">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="04bb1-159">既定値 (存在しない場合) は次の値です `false` 。</span><span class="sxs-lookup"><span data-stu-id="04bb1-159">The default value, if not present, is `false`.</span></span> <span data-ttu-id="04bb1-160">使用する場合 **、OverriddenByRibbonApi は** Group の最初 *の* 子である必要 **があります**。</span><span class="sxs-lookup"><span data-stu-id="04bb1-160">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="04bb1-161">詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="04bb1-161">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
