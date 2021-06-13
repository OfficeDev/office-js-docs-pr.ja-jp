---
title: マニフェスト ファイル内の Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 89ed16f7996ab06bd21e1ebaa71c959b11af2029
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893513"
---
# <a name="group-element"></a><span data-ttu-id="f9356-103">Group 要素</span><span class="sxs-lookup"><span data-stu-id="f9356-103">Group element</span></span>

<span data-ttu-id="f9356-104">タブ内の UI コントロールのグループを定義します。カスタム タブでは、アドインは複数のグループを作成できます。</span><span class="sxs-lookup"><span data-stu-id="f9356-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="f9356-105">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f9356-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="f9356-106">属性</span><span class="sxs-lookup"><span data-stu-id="f9356-106">Attributes</span></span>

|  <span data-ttu-id="f9356-107">属性</span><span class="sxs-lookup"><span data-stu-id="f9356-107">Attribute</span></span>  |  <span data-ttu-id="f9356-108">必須</span><span class="sxs-lookup"><span data-stu-id="f9356-108">Required</span></span>  |  <span data-ttu-id="f9356-109">説明</span><span class="sxs-lookup"><span data-stu-id="f9356-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f9356-110">id</span><span class="sxs-lookup"><span data-stu-id="f9356-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="f9356-111">はい</span><span class="sxs-lookup"><span data-stu-id="f9356-111">Yes</span></span>  | <span data-ttu-id="f9356-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="f9356-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="f9356-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="f9356-113">id attribute</span></span>

<span data-ttu-id="f9356-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="f9356-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f9356-118">子要素</span><span class="sxs-lookup"><span data-stu-id="f9356-118">Child elements</span></span>

|  <span data-ttu-id="f9356-119">要素</span><span class="sxs-lookup"><span data-stu-id="f9356-119">Element</span></span> |  <span data-ttu-id="f9356-120">必須</span><span class="sxs-lookup"><span data-stu-id="f9356-120">Required</span></span>  |  <span data-ttu-id="f9356-121">説明</span><span class="sxs-lookup"><span data-stu-id="f9356-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f9356-122">Label</span><span class="sxs-lookup"><span data-stu-id="f9356-122">Label</span></span>](#label)      | <span data-ttu-id="f9356-123">はい</span><span class="sxs-lookup"><span data-stu-id="f9356-123">Yes</span></span> |  <span data-ttu-id="f9356-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="f9356-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="f9356-125">Icon</span><span class="sxs-lookup"><span data-stu-id="f9356-125">Icon</span></span>](icon.md)      | <span data-ttu-id="f9356-126">はい</span><span class="sxs-lookup"><span data-stu-id="f9356-126">Yes</span></span> |  <span data-ttu-id="f9356-127">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="f9356-127">The image for a group.</span></span> <span data-ttu-id="f9356-128">このアドインではOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-128">Not supported in Outlook add-ins.</span></span> |
|  [<span data-ttu-id="f9356-129">Control</span><span class="sxs-lookup"><span data-stu-id="f9356-129">Control</span></span>](#control)    | <span data-ttu-id="f9356-130">いいえ</span><span class="sxs-lookup"><span data-stu-id="f9356-130">No</span></span> |  <span data-ttu-id="f9356-131">Control オブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="f9356-131">Represents a Control object.</span></span> <span data-ttu-id="f9356-132">0 以上の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f9356-132">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="f9356-133">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="f9356-133">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="f9356-134">いいえ</span><span class="sxs-lookup"><span data-stu-id="f9356-134">No</span></span> | <span data-ttu-id="f9356-135">組み込みのコントロールの 1 Officeします。</span><span class="sxs-lookup"><span data-stu-id="f9356-135">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="f9356-136">0 以上の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f9356-136">Can be zero or more.</span></span> <span data-ttu-id="f9356-137">このアドインではOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-137">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="f9356-138">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f9356-138">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="f9356-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="f9356-139">No</span></span> |  <span data-ttu-id="f9356-140">カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせにグループを表示するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="f9356-140">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="f9356-141">このアドインではOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-141">Not supported in Outlook add-ins.</span></span> |

### <a name="label"></a><span data-ttu-id="f9356-142">Label</span><span class="sxs-lookup"><span data-stu-id="f9356-142">Label</span></span>

<span data-ttu-id="f9356-143">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="f9356-143">Required.</span></span> <span data-ttu-id="f9356-144">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="f9356-144">The label of the group.</span></span> <span data-ttu-id="f9356-145">**resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。</span><span class="sxs-lookup"><span data-stu-id="f9356-145">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="f9356-146">Icon</span><span class="sxs-lookup"><span data-stu-id="f9356-146">Icon</span></span>

<span data-ttu-id="f9356-147">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="f9356-147">Required.</span></span> <span data-ttu-id="f9356-148">タブに多くのグループが含まれている場合、プログラム ウィンドウのサイズが変更された場合、指定したイメージが代わりに表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="f9356-148">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

> [!NOTE]
> <span data-ttu-id="f9356-149">この子要素は、アドインOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-149">This child element is not supported in Outlook add-ins.</span></span>

### <a name="control"></a><span data-ttu-id="f9356-150">コントロール</span><span class="sxs-lookup"><span data-stu-id="f9356-150">Control</span></span>

<span data-ttu-id="f9356-151">省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeControl が必要です**。</span><span class="sxs-lookup"><span data-stu-id="f9356-151">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="f9356-152">サポートされるコントロールの種類の詳細については [、Control](control.md) 要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f9356-152">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="f9356-153">マニフェスト内の **Control** と **OfficeControl** の順序は交換可能で、複数の要素がある場合は相互に混同できますが、すべてが Icon 要素の下にある **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f9356-153">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="f9356-154">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="f9356-154">OfficeControl</span></span>

<span data-ttu-id="f9356-155">オプションですが、存在しない場合は少なくとも 1 つの Control が必要 **です**。</span><span class="sxs-lookup"><span data-stu-id="f9356-155">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="f9356-156">1 つ以上の組み込みOffice要素を含むコントロールをグループに含 `<OfficeControl>` める。</span><span class="sxs-lookup"><span data-stu-id="f9356-156">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="f9356-157">属性 `id` は、組み込みのコントロールの ID をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="f9356-157">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="f9356-158">コントロールの ID を見つけるには、「コントロールとコントロール グループの [ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="f9356-158">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="f9356-159">マニフェスト内の **Control** と **OfficeControl** の順序は交換可能で、複数の要素がある場合は相互に混同できますが、すべてが Icon 要素の下にある **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f9356-159">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

> [!NOTE]
> <span data-ttu-id="f9356-160">この子要素は、アドインOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-160">This child element is not supported in Outlook add-ins.</span></span>

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

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="f9356-161">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f9356-161">OverriddenByRibbonApi</span></span>

<span data-ttu-id="f9356-162">省略可能 (ブール型)。</span><span class="sxs-lookup"><span data-stu-id="f9356-162">Optional (boolean).</span></span> <span data-ttu-id="f9356-163">実行時にリボンにカスタムコンテキスト タブをインストールする API をサポートするアプリケーションとプラットフォームの組み合わせでグループを非表示にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="f9356-163">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="f9356-164">既定値 (存在しない場合) は、 です `false` 。</span><span class="sxs-lookup"><span data-stu-id="f9356-164">The default value, if not present, is `false`.</span></span> <span data-ttu-id="f9356-165">使用する場合 **、OverriddenByRibbonApi は Group** の *最初の* 子である **必要があります**。</span><span class="sxs-lookup"><span data-stu-id="f9356-165">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="f9356-166">詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="f9356-166">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!NOTE]
> <span data-ttu-id="f9356-167">この子要素は、アドインOutlookサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f9356-167">This child element is not supported in Outlook add-ins.</span></span>

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
