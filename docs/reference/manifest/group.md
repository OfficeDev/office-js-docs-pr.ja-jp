---
title: マニフェストファイルの Group 要素
description: タブ内の UI コントロールのグループを定義します。
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: a598232f230a120dccd58024e760c2172a769727
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611828"
---
# <a name="group-element"></a><span data-ttu-id="4a2d0-103">Group 要素</span><span class="sxs-lookup"><span data-stu-id="4a2d0-103">Group element</span></span>

<span data-ttu-id="4a2d0-p101">タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="4a2d0-107">属性</span><span class="sxs-lookup"><span data-stu-id="4a2d0-107">Attributes</span></span>

|  <span data-ttu-id="4a2d0-108">属性</span><span class="sxs-lookup"><span data-stu-id="4a2d0-108">Attribute</span></span>  |  <span data-ttu-id="4a2d0-109">必須</span><span class="sxs-lookup"><span data-stu-id="4a2d0-109">Required</span></span>  |  <span data-ttu-id="4a2d0-110">説明</span><span class="sxs-lookup"><span data-stu-id="4a2d0-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4a2d0-111">id</span><span class="sxs-lookup"><span data-stu-id="4a2d0-111">id</span></span>](#id-attribute)  |  <span data-ttu-id="4a2d0-112">はい</span><span class="sxs-lookup"><span data-stu-id="4a2d0-112">Yes</span></span>  | <span data-ttu-id="4a2d0-113">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-113">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="4a2d0-114">id 属性</span><span class="sxs-lookup"><span data-stu-id="4a2d0-114">id attribute</span></span>

<span data-ttu-id="4a2d0-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4a2d0-119">子要素</span><span class="sxs-lookup"><span data-stu-id="4a2d0-119">Child elements</span></span>
|  <span data-ttu-id="4a2d0-120">要素</span><span class="sxs-lookup"><span data-stu-id="4a2d0-120">Element</span></span> |  <span data-ttu-id="4a2d0-121">必須</span><span class="sxs-lookup"><span data-stu-id="4a2d0-121">Required</span></span>  |  <span data-ttu-id="4a2d0-122">説明</span><span class="sxs-lookup"><span data-stu-id="4a2d0-122">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4a2d0-123">Label</span><span class="sxs-lookup"><span data-stu-id="4a2d0-123">Label</span></span>](#label)      | <span data-ttu-id="4a2d0-124">○</span><span class="sxs-lookup"><span data-stu-id="4a2d0-124">Yes</span></span> |  <span data-ttu-id="4a2d0-125">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-125">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="4a2d0-126">Icon</span><span class="sxs-lookup"><span data-stu-id="4a2d0-126">Icon</span></span>](icon.md)      | <span data-ttu-id="4a2d0-127">はい</span><span class="sxs-lookup"><span data-stu-id="4a2d0-127">Yes</span></span> |  <span data-ttu-id="4a2d0-128">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-128">The image for a group.</span></span>  |
|  [<span data-ttu-id="4a2d0-129">Control</span><span class="sxs-lookup"><span data-stu-id="4a2d0-129">Control</span></span>](#control)    | <span data-ttu-id="4a2d0-130">はい</span><span class="sxs-lookup"><span data-stu-id="4a2d0-130">Yes</span></span> |  <span data-ttu-id="4a2d0-131">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-131">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="4a2d0-132">Label</span><span class="sxs-lookup"><span data-stu-id="4a2d0-132">Label</span></span> 

<span data-ttu-id="4a2d0-133">必須。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-133">Required.</span></span> <span data-ttu-id="4a2d0-134">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-134">The label of the group.</span></span> <span data-ttu-id="4a2d0-135">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="4a2d0-136">Icon</span><span class="sxs-lookup"><span data-stu-id="4a2d0-136">Icon</span></span>

<span data-ttu-id="4a2d0-137">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-137">Required.</span></span> <span data-ttu-id="4a2d0-138">タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-138">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="4a2d0-139">Control</span><span class="sxs-lookup"><span data-stu-id="4a2d0-139">Control</span></span>
<span data-ttu-id="4a2d0-140">1 つのグループに少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-140">A group requires at least one control.</span></span> <span data-ttu-id="4a2d0-141">サポートされているコントロールの種類の詳細については、 [Control](control.md)要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4a2d0-141">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
