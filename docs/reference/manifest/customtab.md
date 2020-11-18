---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087939"
---
# <a name="customtab-element"></a><span data-ttu-id="ec7e4-103">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="ec7e4-103">CustomTab element</span></span>

<span data-ttu-id="ec7e4-104">リボンで、アドインコマンドのタブとグループを指定します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="ec7e4-105">これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="ec7e4-106">カスタムタブでは、アドインにカスタムグループまたは組み込みグループを含めることができます。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="ec7e4-107">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="ec7e4-108">**Id** 属性はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ec7e4-109">Mac 上の Outlook では、要素を使用でき `CustomTab` ないため、代わりに [[officetab タブ](officetab.md) を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ec7e4-110">子要素</span><span class="sxs-lookup"><span data-stu-id="ec7e4-110">Child elements</span></span>

|  <span data-ttu-id="ec7e4-111">要素</span><span class="sxs-lookup"><span data-stu-id="ec7e4-111">Element</span></span> |  <span data-ttu-id="ec7e4-112">必須</span><span class="sxs-lookup"><span data-stu-id="ec7e4-112">Required</span></span>  |  <span data-ttu-id="ec7e4-113">説明</span><span class="sxs-lookup"><span data-stu-id="ec7e4-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ec7e4-114">Group</span><span class="sxs-lookup"><span data-stu-id="ec7e4-114">Group</span></span>](group.md)      | <span data-ttu-id="ec7e4-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="ec7e4-115">No</span></span> |  <span data-ttu-id="ec7e4-116">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="ec7e4-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="ec7e4-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="ec7e4-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="ec7e4-118">No</span></span> |  <span data-ttu-id="ec7e4-119">組み込みの Office コントロールグループを表します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="ec7e4-120">Label</span><span class="sxs-lookup"><span data-stu-id="ec7e4-120">Label</span></span>](#label-tab)      | <span data-ttu-id="ec7e4-121">はい</span><span class="sxs-lookup"><span data-stu-id="ec7e4-121">Yes</span></span> |  <span data-ttu-id="ec7e4-122">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="ec7e4-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="ec7e4-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="ec7e4-124">いいえ</span><span class="sxs-lookup"><span data-stu-id="ec7e4-124">No</span></span> |  <span data-ttu-id="ec7e4-125">指定した組み込みの Office タブの直後にカスタムタブを作成するように指定します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="ec7e4-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="ec7e4-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="ec7e4-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="ec7e4-127">No</span></span> |  <span data-ttu-id="ec7e4-128">指定した組み込みの Office タブの直前にカスタムタブを表示するように指定します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="ec7e4-129">Group</span><span class="sxs-lookup"><span data-stu-id="ec7e4-129">Group</span></span>

<span data-ttu-id="ec7e4-130">省略可能。ただし、指定されていない場合は、少なくとも1つの **Officegroup** 要素が存在している必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="ec7e4-131">[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-131">See [Group element](group.md).</span></span> <span data-ttu-id="ec7e4-132">マニフェスト内の **グループ** と **officegroup** の順序は、[ユーザー設定] タブに表示する順序にする必要があります。複数の要素がある場合には混在させることができますが、 **Label** 要素の上にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="ec7e4-133">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="ec7e4-133">OfficeGroup</span></span>

<span data-ttu-id="ec7e4-134">省略可能。ただし、指定されていない場合は、少なくとも1つの **Group** 要素が存在している必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="ec7e4-135">組み込みの Office コントロールグループを表します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="ec7e4-136">**Id** 属性は、組み込みの Office グループの id を指定します。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="ec7e4-137">組み込みグループの ID を検索するには、「 [コントロールおよびコントロールグループの id を検索](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="ec7e4-138">マニフェスト内の **グループ** と **officegroup** の順序は、[ユーザー設定] タブに表示する順序にする必要があります。複数の要素がある場合には混在させることができますが、 **Label** 要素の上にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="ec7e4-139">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="ec7e4-139">Label (Tab)</span></span>

<span data-ttu-id="ec7e4-140">必須です。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-140">Required.</span></span> <span data-ttu-id="ec7e4-141">ユーザー設定のタブのラベルを示します。**Resid** 属性は、 [Resources](resources.md)要素の Short **strings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-141">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="ec7e4-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="ec7e4-142">InsertAfter</span></span>

<span data-ttu-id="ec7e4-143">省略可能。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-143">Optional.</span></span> <span data-ttu-id="ec7e4-144">指定した組み込みの Office タブの直後にカスタムタブを作成するように指定します。要素の値は、組み込みタブの ID ("TabHome"、"Tabhome" など) です。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="ec7e4-145">(「 [コントロールおよびコントロールグループの id を検索する」を](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)参照してください)。指定する場合は、 **Label** 要素の後にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="ec7e4-146">**InsertAfter** と **insertbefore** の両方を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="ec7e4-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="ec7e4-147">InsertBefore</span></span>

<span data-ttu-id="ec7e4-148">省略可能。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-148">Optional.</span></span> <span data-ttu-id="ec7e4-149">指定した組み込みの Office タブの直前にカスタムタブを表示するように指定します。要素の値は、組み込みタブの ID ("TabHome"、"Tabhome" など) です。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="ec7e4-150">(「 [コントロールおよびコントロールグループの id を検索する」を](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)参照してください)。 指定する場合は、 **Label** 要素の後にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="ec7e4-151">**InsertAfter** と **insertbefore** の両方を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="ec7e4-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="ec7e4-152">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="ec7e4-152">CustomTab example</span></span>

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
