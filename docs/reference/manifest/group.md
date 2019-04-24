---
title: マニフェストファイルの Group 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7cc1f4c398eeb013eb6033b207b395466f7d72ca
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450710"
---
# <a name="group-element"></a><span data-ttu-id="b31a6-102">Group 要素</span><span class="sxs-lookup"><span data-stu-id="b31a6-102">Group element</span></span>

<span data-ttu-id="b31a6-p101">タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="b31a6-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="b31a6-106">属性</span><span class="sxs-lookup"><span data-stu-id="b31a6-106">Attributes</span></span>

|  <span data-ttu-id="b31a6-107">属性</span><span class="sxs-lookup"><span data-stu-id="b31a6-107">Attribute</span></span>  |  <span data-ttu-id="b31a6-108">必須</span><span class="sxs-lookup"><span data-stu-id="b31a6-108">Required</span></span>  |  <span data-ttu-id="b31a6-109">説明</span><span class="sxs-lookup"><span data-stu-id="b31a6-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b31a6-110">id</span><span class="sxs-lookup"><span data-stu-id="b31a6-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="b31a6-111">はい</span><span class="sxs-lookup"><span data-stu-id="b31a6-111">Yes</span></span>  | <span data-ttu-id="b31a6-112">グループの一意の ID。</span><span class="sxs-lookup"><span data-stu-id="b31a6-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="b31a6-113">id 属性</span><span class="sxs-lookup"><span data-stu-id="b31a6-113">id attribute</span></span>

<span data-ttu-id="b31a6-p102">必須。 グループの一意識別子。 最大 125 文字の文字列です。 マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。</span><span class="sxs-lookup"><span data-stu-id="b31a6-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b31a6-118">子要素</span><span class="sxs-lookup"><span data-stu-id="b31a6-118">Child elements</span></span>
|  <span data-ttu-id="b31a6-119">要素</span><span class="sxs-lookup"><span data-stu-id="b31a6-119">Element</span></span> |  <span data-ttu-id="b31a6-120">必須</span><span class="sxs-lookup"><span data-stu-id="b31a6-120">Required</span></span>  |  <span data-ttu-id="b31a6-121">説明</span><span class="sxs-lookup"><span data-stu-id="b31a6-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b31a6-122">Label</span><span class="sxs-lookup"><span data-stu-id="b31a6-122">Label</span></span>](#label)      | <span data-ttu-id="b31a6-123">○</span><span class="sxs-lookup"><span data-stu-id="b31a6-123">Yes</span></span> |  <span data-ttu-id="b31a6-124">CustomTab またはグループのラベル。</span><span class="sxs-lookup"><span data-stu-id="b31a6-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="b31a6-125">Control</span><span class="sxs-lookup"><span data-stu-id="b31a6-125">Control</span></span>](#control)    | <span data-ttu-id="b31a6-126">はい</span><span class="sxs-lookup"><span data-stu-id="b31a6-126">Yes</span></span> |  <span data-ttu-id="b31a6-127">1 つ以上のコントロール オブジェクトのコレクション。</span><span class="sxs-lookup"><span data-stu-id="b31a6-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="b31a6-128">Label</span><span class="sxs-lookup"><span data-stu-id="b31a6-128">Label</span></span> 

<span data-ttu-id="b31a6-p103">必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b31a6-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="b31a6-132">Control</span><span class="sxs-lookup"><span data-stu-id="b31a6-132">Control</span></span>
<span data-ttu-id="b31a6-133">1 つのグループに少なくとも 1 つのコントロールが必要です。</span><span class="sxs-lookup"><span data-stu-id="b31a6-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
