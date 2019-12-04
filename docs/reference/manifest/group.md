---
title: マニフェストファイルの Group 要素
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 35db4829b40078e97fbfc007e2fb552e00875f9c
ms.sourcegitcommit: 164b11b1e9d2ae20b3d816092025b32a9070450f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2019
ms.locfileid: "39818728"
---
# <a name="group-element"></a><span data-ttu-id="97719-102">Group 要素</span><span class="sxs-lookup"><span data-stu-id="97719-102">Group element</span></span>

<span data-ttu-id="97719-p101">タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="97719-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="97719-106">属性</span><span class="sxs-lookup"><span data-stu-id="97719-106">Attributes</span></span>

|  <span data-ttu-id="97719-107">属性</span><span class="sxs-lookup"><span data-stu-id="97719-107">Attribute</span></span>  |  <span data-ttu-id="97719-108">必須</span><span class="sxs-lookup"><span data-stu-id="97719-108">Required</span></span>  |  <span data-ttu-id="97719-109">説明</span><span class="sxs-lookup"><span data-stu-id="97719-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="97719-110">id</span><span class="sxs-lookup"><span data-stu-id="97719-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="97719-111">はい</span><span class="sxs-lookup"><span data-stu-id="97719-111">Yes</span></span>  | <span data-ttu-id="97719-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="97719-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="97719-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="97719-113">id attribute</span></span>

<span data-ttu-id="97719-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="97719-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="97719-118">子要素</span><span class="sxs-lookup"><span data-stu-id="97719-118">Child elements</span></span>
|  <span data-ttu-id="97719-119">要素</span><span class="sxs-lookup"><span data-stu-id="97719-119">Element</span></span> |  <span data-ttu-id="97719-120">必須</span><span class="sxs-lookup"><span data-stu-id="97719-120">Required</span></span>  |  <span data-ttu-id="97719-121">説明</span><span class="sxs-lookup"><span data-stu-id="97719-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="97719-122">Label</span><span class="sxs-lookup"><span data-stu-id="97719-122">Label</span></span>](#label)      | <span data-ttu-id="97719-123">○</span><span class="sxs-lookup"><span data-stu-id="97719-123">Yes</span></span> |  <span data-ttu-id="97719-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="97719-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="97719-125">Icon</span><span class="sxs-lookup"><span data-stu-id="97719-125">Icon</span></span>](icon.md)      | <span data-ttu-id="97719-126">はい</span><span class="sxs-lookup"><span data-stu-id="97719-126">Yes</span></span> |  <span data-ttu-id="97719-127">グループのイメージ。</span><span class="sxs-lookup"><span data-stu-id="97719-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="97719-128">Control</span><span class="sxs-lookup"><span data-stu-id="97719-128">Control</span></span>](#control)    | <span data-ttu-id="97719-129">はい</span><span class="sxs-lookup"><span data-stu-id="97719-129">Yes</span></span> |  <span data-ttu-id="97719-130">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="97719-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="97719-131">Label</span><span class="sxs-lookup"><span data-stu-id="97719-131">Label</span></span> 

<span data-ttu-id="97719-p103">必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="97719-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="97719-135">Icon</span><span class="sxs-lookup"><span data-stu-id="97719-135">Icon</span></span>

<span data-ttu-id="97719-136">必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="97719-136">Required.</span></span> <span data-ttu-id="97719-137">タブに多数のグループが含まれ、プログラムウィンドウのサイズが変更されると、代わりに、指定したイメージが表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="97719-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="97719-138">Control</span><span class="sxs-lookup"><span data-stu-id="97719-138">Control</span></span>
<span data-ttu-id="97719-139">1 つのグループに少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="97719-139">A group requires at least one control.</span></span> <span data-ttu-id="97719-140">サポートされているコントロールの種類の詳細については、 [Control](control.md)要素を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97719-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
