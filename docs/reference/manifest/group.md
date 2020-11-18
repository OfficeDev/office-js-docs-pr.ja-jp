---
title: マニフェストファイルの Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087946"
---
# <a name="group-element"></a><span data-ttu-id="c78a4-103">Group 要素</span><span class="sxs-lookup"><span data-stu-id="c78a4-103">Group element</span></span>

<span data-ttu-id="c78a4-104">タブ内の UI コントロールのグループを定義します。カスタムタブでは、アドインは複数のグループを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c78a4-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="c78a4-105">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c78a4-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="c78a4-106">属性</span><span class="sxs-lookup"><span data-stu-id="c78a4-106">Attributes</span></span>

|  <span data-ttu-id="c78a4-107">属性</span><span class="sxs-lookup"><span data-stu-id="c78a4-107">Attribute</span></span>  |  <span data-ttu-id="c78a4-108">必須</span><span class="sxs-lookup"><span data-stu-id="c78a4-108">Required</span></span>  |  <span data-ttu-id="c78a4-109">説明</span><span class="sxs-lookup"><span data-stu-id="c78a4-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c78a4-110">id</span><span class="sxs-lookup"><span data-stu-id="c78a4-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="c78a4-111">はい</span><span class="sxs-lookup"><span data-stu-id="c78a4-111">Yes</span></span>  | <span data-ttu-id="c78a4-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="c78a4-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="c78a4-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="c78a4-113">id attribute</span></span>

<span data-ttu-id="c78a4-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c78a4-118">子要素</span><span class="sxs-lookup"><span data-stu-id="c78a4-118">Child elements</span></span>

|  <span data-ttu-id="c78a4-119">要素</span><span class="sxs-lookup"><span data-stu-id="c78a4-119">Element</span></span> |  <span data-ttu-id="c78a4-120">必須</span><span class="sxs-lookup"><span data-stu-id="c78a4-120">Required</span></span>  |  <span data-ttu-id="c78a4-121">説明</span><span class="sxs-lookup"><span data-stu-id="c78a4-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c78a4-122">Label</span><span class="sxs-lookup"><span data-stu-id="c78a4-122">Label</span></span>](#label)      | <span data-ttu-id="c78a4-123">はい</span><span class="sxs-lookup"><span data-stu-id="c78a4-123">Yes</span></span> |  <span data-ttu-id="c78a4-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="c78a4-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="c78a4-125">Icon</span><span class="sxs-lookup"><span data-stu-id="c78a4-125">Icon</span></span>](icon.md)      | <span data-ttu-id="c78a4-126">はい</span><span class="sxs-lookup"><span data-stu-id="c78a4-126">Yes</span></span> |  <span data-ttu-id="c78a4-127">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="c78a4-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="c78a4-128">Control</span><span class="sxs-lookup"><span data-stu-id="c78a4-128">Control</span></span>](#control)    | <span data-ttu-id="c78a4-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="c78a4-129">No</span></span> |  <span data-ttu-id="c78a4-130">Control オブジェクトを表します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-130">Represents a Control object.</span></span> <span data-ttu-id="c78a4-131">0個以上の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c78a4-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="c78a4-132">Officeecontrol</span><span class="sxs-lookup"><span data-stu-id="c78a4-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="c78a4-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="c78a4-133">No</span></span> | <span data-ttu-id="c78a4-134">組み込みの Office コントロールの1つを表します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="c78a4-135">0個以上の値を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c78a4-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="c78a4-136">Label</span><span class="sxs-lookup"><span data-stu-id="c78a4-136">Label</span></span>

<span data-ttu-id="c78a4-137">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-137">Required.</span></span> <span data-ttu-id="c78a4-138">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="c78a4-138">The label of the group.</span></span> <span data-ttu-id="c78a4-139">**Resid** 属性は、 [Resources](resources.md)要素の Short **strings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c78a4-139">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="c78a4-140">Icon</span><span class="sxs-lookup"><span data-stu-id="c78a4-140">Icon</span></span>

<span data-ttu-id="c78a4-141">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-141">Required.</span></span> <span data-ttu-id="c78a4-142">タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="c78a4-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="c78a4-143">コントロール</span><span class="sxs-lookup"><span data-stu-id="c78a4-143">Control</span></span>

<span data-ttu-id="c78a4-144">省略可能。ただし、存在しない場合は、少なくとも1つの **Officeecontrol** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c78a4-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="c78a4-145">サポートされているコントロールの種類の詳細については、 [Control](control.md) 要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c78a4-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="c78a4-146">マニフェストでは、 **Control** と are **econtrol** の順序は相互に置き換え可能で、複数の要素がある場合は混在させることができますが、すべてが **Icon** 要素の下になければなりません。</span><span class="sxs-lookup"><span data-stu-id="c78a4-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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

### <a name="officecontrol"></a><span data-ttu-id="c78a4-147">Officeecontrol</span><span class="sxs-lookup"><span data-stu-id="c78a4-147">OfficeControl</span></span>

<span data-ttu-id="c78a4-148">省略可能。ただし、存在しない場合は、少なくとも1つの **コントロール** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c78a4-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="c78a4-149">1つ以上の組み込みの Office コントロールを要素を含むグループに含め `<OfficeControl>` ます。</span><span class="sxs-lookup"><span data-stu-id="c78a4-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="c78a4-150">属性は、 `id` 組み込みの Office コントロールの ID を指定します。</span><span class="sxs-lookup"><span data-stu-id="c78a4-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="c78a4-151">コントロールの ID を検索するには、「 [コントロールおよびコントロールグループの id を検索](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c78a4-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="c78a4-152">マニフェストでは、 **Control** と are **econtrol** の順序は相互に置き換え可能で、複数の要素がある場合は混在させることができますが、すべてが **Icon** 要素の下になければなりません。</span><span class="sxs-lookup"><span data-stu-id="c78a4-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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
