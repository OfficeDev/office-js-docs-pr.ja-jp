---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771327"
---
# <a name="customtab-element"></a><span data-ttu-id="20932-103">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="20932-103">CustomTab element</span></span>

<span data-ttu-id="20932-104">リボンで、アドイン コマンドのタブとグループを指定します。</span><span class="sxs-lookup"><span data-stu-id="20932-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="20932-105">これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="20932-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="20932-106">カスタム タブでは、アドインにカスタム グループまたは組み込みグループを含めできます。</span><span class="sxs-lookup"><span data-stu-id="20932-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="20932-107">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="20932-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="20932-108">**id 属性** はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="20932-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="20932-109">Outlook on Mac では、この要素 `CustomTab` は使用できないので、代わりに [OfficeTab を使用する](officetab.md) 必要があります。</span><span class="sxs-lookup"><span data-stu-id="20932-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="20932-110">子要素</span><span class="sxs-lookup"><span data-stu-id="20932-110">Child elements</span></span>

|  <span data-ttu-id="20932-111">要素</span><span class="sxs-lookup"><span data-stu-id="20932-111">Element</span></span> |  <span data-ttu-id="20932-112">必須</span><span class="sxs-lookup"><span data-stu-id="20932-112">Required</span></span>  |  <span data-ttu-id="20932-113">説明</span><span class="sxs-lookup"><span data-stu-id="20932-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="20932-114">Group</span><span class="sxs-lookup"><span data-stu-id="20932-114">Group</span></span>](group.md)      | <span data-ttu-id="20932-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="20932-115">No</span></span> |  <span data-ttu-id="20932-116">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="20932-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="20932-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="20932-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="20932-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="20932-118">No</span></span> |  <span data-ttu-id="20932-119">組み込みのコントロール グループOffice表します。</span><span class="sxs-lookup"><span data-stu-id="20932-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="20932-120">Label</span><span class="sxs-lookup"><span data-stu-id="20932-120">Label</span></span>](#label-tab)      | <span data-ttu-id="20932-121">はい</span><span class="sxs-lookup"><span data-stu-id="20932-121">Yes</span></span> |  <span data-ttu-id="20932-122">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="20932-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="20932-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="20932-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="20932-124">いいえ</span><span class="sxs-lookup"><span data-stu-id="20932-124">No</span></span> |  <span data-ttu-id="20932-125">指定した組み込みのタブの直後にカスタム タブをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="20932-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="20932-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="20932-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="20932-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="20932-127">No</span></span> |  <span data-ttu-id="20932-128">ユーザー設定タブを指定した組み込みタブの直前にOfficeします。</span><span class="sxs-lookup"><span data-stu-id="20932-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="20932-129">グループ</span><span class="sxs-lookup"><span data-stu-id="20932-129">Group</span></span>

<span data-ttu-id="20932-130">省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。</span><span class="sxs-lookup"><span data-stu-id="20932-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="20932-131">[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="20932-131">See [Group element](group.md).</span></span> <span data-ttu-id="20932-132">マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべて Label 要素の上に配置する **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="20932-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="20932-133">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="20932-133">OfficeGroup</span></span>

<span data-ttu-id="20932-134">省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。</span><span class="sxs-lookup"><span data-stu-id="20932-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="20932-135">組み込みのコントロール グループOffice表します。</span><span class="sxs-lookup"><span data-stu-id="20932-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="20932-136">**id 属性** は、グループに組み込Officeします。</span><span class="sxs-lookup"><span data-stu-id="20932-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="20932-137">組み込みのグループの ID を検索するには、「コントロールとコントロール グループの ID を検索する」 [を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="20932-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="20932-138">マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべて Label 要素の上に配置する **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="20932-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="20932-139">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="20932-139">Label (Tab)</span></span>

<span data-ttu-id="20932-140">Required.</span><span class="sxs-lookup"><span data-stu-id="20932-140">Required.</span></span> <span data-ttu-id="20932-141">カスタム タブのラベルを指定します。**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="20932-141">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="20932-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="20932-142">InsertAfter</span></span>

<span data-ttu-id="20932-143">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="20932-143">Optional.</span></span> <span data-ttu-id="20932-144">指定した組み込みのタブの直後にカスタム タブをOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。</span><span class="sxs-lookup"><span data-stu-id="20932-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="20932-145">(「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。存在する場合は、Label 要素の後 **に配置する必要** があります。</span><span class="sxs-lookup"><span data-stu-id="20932-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="20932-146">InsertAfter と **InsertBefore の両方を指定することはできません**。</span><span class="sxs-lookup"><span data-stu-id="20932-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="20932-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="20932-147">InsertBefore</span></span>

<span data-ttu-id="20932-148">省略可能です。</span><span class="sxs-lookup"><span data-stu-id="20932-148">Optional.</span></span> <span data-ttu-id="20932-149">ユーザー設定タブを指定した組み込みタブの直前にOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。</span><span class="sxs-lookup"><span data-stu-id="20932-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="20932-150">(「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。 存在する場合は、Label 要素の後 **に配置する必要** があります。</span><span class="sxs-lookup"><span data-stu-id="20932-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="20932-151">InsertAfter と **InsertBefore の両方を指定することはできません**。</span><span class="sxs-lookup"><span data-stu-id="20932-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="20932-152">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="20932-152">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
