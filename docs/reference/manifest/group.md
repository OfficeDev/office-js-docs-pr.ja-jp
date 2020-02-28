---
title: マニフェストファイルの Group 要素
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 27a168ea17352482e955e7a0d1f8267c7d6b17d8
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324863"
---
# <a name="group-element"></a><span data-ttu-id="4df36-102">Group 要素</span><span class="sxs-lookup"><span data-stu-id="4df36-102">Group element</span></span>

<span data-ttu-id="4df36-p101">タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="4df36-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="4df36-106">属性</span><span class="sxs-lookup"><span data-stu-id="4df36-106">Attributes</span></span>

|  <span data-ttu-id="4df36-107">属性</span><span class="sxs-lookup"><span data-stu-id="4df36-107">Attribute</span></span>  |  <span data-ttu-id="4df36-108">必須</span><span class="sxs-lookup"><span data-stu-id="4df36-108">Required</span></span>  |  <span data-ttu-id="4df36-109">説明</span><span class="sxs-lookup"><span data-stu-id="4df36-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4df36-110">id</span><span class="sxs-lookup"><span data-stu-id="4df36-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="4df36-111">はい</span><span class="sxs-lookup"><span data-stu-id="4df36-111">Yes</span></span>  | <span data-ttu-id="4df36-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="4df36-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="4df36-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="4df36-113">id attribute</span></span>

<span data-ttu-id="4df36-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="4df36-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4df36-118">子要素</span><span class="sxs-lookup"><span data-stu-id="4df36-118">Child elements</span></span>
|  <span data-ttu-id="4df36-119">要素</span><span class="sxs-lookup"><span data-stu-id="4df36-119">Element</span></span> |  <span data-ttu-id="4df36-120">必須</span><span class="sxs-lookup"><span data-stu-id="4df36-120">Required</span></span>  |  <span data-ttu-id="4df36-121">説明</span><span class="sxs-lookup"><span data-stu-id="4df36-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4df36-122">Label</span><span class="sxs-lookup"><span data-stu-id="4df36-122">Label</span></span>](#label)      | <span data-ttu-id="4df36-123">○</span><span class="sxs-lookup"><span data-stu-id="4df36-123">Yes</span></span> |  <span data-ttu-id="4df36-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="4df36-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="4df36-125">Icon</span><span class="sxs-lookup"><span data-stu-id="4df36-125">Icon</span></span>](icon.md)      | <span data-ttu-id="4df36-126">はい</span><span class="sxs-lookup"><span data-stu-id="4df36-126">Yes</span></span> |  <span data-ttu-id="4df36-127">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="4df36-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="4df36-128">Control</span><span class="sxs-lookup"><span data-stu-id="4df36-128">Control</span></span>](#control)    | <span data-ttu-id="4df36-129">はい</span><span class="sxs-lookup"><span data-stu-id="4df36-129">Yes</span></span> |  <span data-ttu-id="4df36-130">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="4df36-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="4df36-131">Label</span><span class="sxs-lookup"><span data-stu-id="4df36-131">Label</span></span> 

<span data-ttu-id="4df36-132">必須。</span><span class="sxs-lookup"><span data-stu-id="4df36-132">Required.</span></span> <span data-ttu-id="4df36-133">グループのラベルです。</span><span class="sxs-lookup"><span data-stu-id="4df36-133">The label of the group.</span></span> <span data-ttu-id="4df36-134">**Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4df36-134">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="4df36-135">Icon</span><span class="sxs-lookup"><span data-stu-id="4df36-135">Icon</span></span>

<span data-ttu-id="4df36-136">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="4df36-136">Required.</span></span> <span data-ttu-id="4df36-137">タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="4df36-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="4df36-138">Control</span><span class="sxs-lookup"><span data-stu-id="4df36-138">Control</span></span>
<span data-ttu-id="4df36-139">1 つのグループに少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="4df36-139">A group requires at least one control.</span></span> <span data-ttu-id="4df36-140">サポートされているコントロールの種類の詳細については、 [Control](control.md)要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4df36-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
