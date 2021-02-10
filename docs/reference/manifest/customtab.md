---
title: マニフェスト ファイルの CustomTab 要素
description: リボン上で、アドイン コマンドに使用するタブとグループを指定します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173928"
---
# <a name="customtab-element"></a><span data-ttu-id="f99db-103">CustomTab 要素</span><span class="sxs-lookup"><span data-stu-id="f99db-103">CustomTab element</span></span>

<span data-ttu-id="f99db-104">リボンで、アドイン コマンドのタブとグループを指定します。</span><span class="sxs-lookup"><span data-stu-id="f99db-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="f99db-105">これは既定のタブ ([**ホーム**]、[**メッセージ**]、[**会議**] のいずれか)、またはアドインで定義されたカスタム タブになります。</span><span class="sxs-lookup"><span data-stu-id="f99db-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="f99db-106">カスタム タブでは、アドインにカスタム グループまたは組み込みグループを含めできます。</span><span class="sxs-lookup"><span data-stu-id="f99db-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="f99db-107">アドインは、カスタム タブ 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f99db-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="f99db-108">**id 属性** はマニフェスト内で一意である必要があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f99db-109">Outlook on Mac では、この要素 `CustomTab` は使用できないので、代わりに [OfficeTab を使用する](officetab.md) 必要があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f99db-110">子要素</span><span class="sxs-lookup"><span data-stu-id="f99db-110">Child elements</span></span>

|  <span data-ttu-id="f99db-111">要素</span><span class="sxs-lookup"><span data-stu-id="f99db-111">Element</span></span> |  <span data-ttu-id="f99db-112">必須</span><span class="sxs-lookup"><span data-stu-id="f99db-112">Required</span></span>  |  <span data-ttu-id="f99db-113">説明</span><span class="sxs-lookup"><span data-stu-id="f99db-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f99db-114">Group</span><span class="sxs-lookup"><span data-stu-id="f99db-114">Group</span></span>](group.md)      | <span data-ttu-id="f99db-115">いいえ</span><span class="sxs-lookup"><span data-stu-id="f99db-115">No</span></span> |  <span data-ttu-id="f99db-116">コマンドのグループを定義します。</span><span class="sxs-lookup"><span data-stu-id="f99db-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="f99db-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f99db-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="f99db-118">いいえ</span><span class="sxs-lookup"><span data-stu-id="f99db-118">No</span></span> |  <span data-ttu-id="f99db-119">組み込みのコントロール グループOffice表します。</span><span class="sxs-lookup"><span data-stu-id="f99db-119">Represents a built-in Office control group.</span></span> <span data-ttu-id="f99db-120">**重要**: Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-120">**Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f99db-121">Label</span><span class="sxs-lookup"><span data-stu-id="f99db-121">Label</span></span>](#label-tab)      | <span data-ttu-id="f99db-122">はい</span><span class="sxs-lookup"><span data-stu-id="f99db-122">Yes</span></span> |  <span data-ttu-id="f99db-123">CustomTab または Group のラベル。</span><span class="sxs-lookup"><span data-stu-id="f99db-123">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="f99db-124">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f99db-124">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="f99db-125">いいえ</span><span class="sxs-lookup"><span data-stu-id="f99db-125">No</span></span> |  <span data-ttu-id="f99db-126">ユーザー設定のタブを、指定した組み込みのタブの直後に表示Office指定します。 **重要**: Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-126">Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f99db-127">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f99db-127">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="f99db-128">いいえ</span><span class="sxs-lookup"><span data-stu-id="f99db-128">No</span></span> |  <span data-ttu-id="f99db-129">ユーザー設定のタブを、指定した組み込みのタブの直前にOffice指定します。 **重要**: Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-129">Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f99db-130">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f99db-130">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="f99db-131">いいえ</span><span class="sxs-lookup"><span data-stu-id="f99db-131">No</span></span> |  <span data-ttu-id="f99db-132">カスタム タブを、カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームの組み合わせに表示するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="f99db-132">Specifies whether the custom tab should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="f99db-133">**重要**: Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-133">**Important**: Not available in Outlook.</span></span> |

### <a name="group"></a><span data-ttu-id="f99db-134">グループ</span><span class="sxs-lookup"><span data-stu-id="f99db-134">Group</span></span>

<span data-ttu-id="f99db-135">省略可能ですが、存在しない場合は、少なくとも 1 つの **OfficeGroup 要素が必要** です。</span><span class="sxs-lookup"><span data-stu-id="f99db-135">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="f99db-136">[Group 要素](group.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f99db-136">See [Group element](group.md).</span></span> <span data-ttu-id="f99db-137">マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべてが Label 要素の上にある **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-137">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="f99db-138">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f99db-138">OfficeGroup</span></span>

<span data-ttu-id="f99db-139">省略可能ですが、存在しない場合は、少なくとも 1 つの Group 要素が **必要** です。</span><span class="sxs-lookup"><span data-stu-id="f99db-139">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="f99db-140">組み込みのコントロール グループOffice表します。</span><span class="sxs-lookup"><span data-stu-id="f99db-140">Represents a built-in Office control group.</span></span> <span data-ttu-id="f99db-141">**id 属性** は、グループに組み込Officeします。</span><span class="sxs-lookup"><span data-stu-id="f99db-141">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="f99db-142">組み込みのグループの ID を検索するには、「コントロールとコントロール グループの ID を検索する」 [を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="f99db-142">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="f99db-143">マニフェスト内 **のグループ** と **OfficeGroup** の順序は、カスタム タブに表示する順序である必要があります。要素が複数ある場合は、これらの要素が不確定になる可能性がありますが、すべてが Label 要素の上にある **必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-143">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f99db-144">この `OfficeGroup` 要素は Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-144">The `OfficeGroup` element is not available in Outlook.</span></span>

### <a name="label-tab"></a><span data-ttu-id="f99db-145">Label (タブ)</span><span class="sxs-lookup"><span data-stu-id="f99db-145">Label (Tab)</span></span>

<span data-ttu-id="f99db-146">必須です。</span><span class="sxs-lookup"><span data-stu-id="f99db-146">Required.</span></span> <span data-ttu-id="f99db-147">カスタム タブのラベルを指定します。**resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-147">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="f99db-148">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f99db-148">InsertAfter</span></span>

<span data-ttu-id="f99db-149">省略可能。</span><span class="sxs-lookup"><span data-stu-id="f99db-149">Optional.</span></span> <span data-ttu-id="f99db-150">指定した組み込みのタブの直後にカスタム タブをOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。</span><span class="sxs-lookup"><span data-stu-id="f99db-150">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f99db-151">(「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。存在する場合は、Label 要素の後 **に配置する必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-151">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="f99db-152">InsertAfter と **InsertBefore の両方を指定することはできません**。</span><span class="sxs-lookup"><span data-stu-id="f99db-152">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f99db-153">この `InsertAfter` 要素は Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-153">The `InsertAfter` element is not available in Outlook.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="f99db-154">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f99db-154">InsertBefore</span></span>

<span data-ttu-id="f99db-155">省略可能。</span><span class="sxs-lookup"><span data-stu-id="f99db-155">Optional.</span></span> <span data-ttu-id="f99db-156">ユーザー設定タブを指定した組み込みタブの直前にOfficeします。要素の値は、"TabHome" や "TabReview" などの組み込みタブの ID です。</span><span class="sxs-lookup"><span data-stu-id="f99db-156">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f99db-157">(「 [コントロールとコントロール グループの ID を検索する」を参照してください](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups))。 存在する場合は、Label 要素の後 **に配置する必要** があります。</span><span class="sxs-lookup"><span data-stu-id="f99db-157">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="f99db-158">InsertAfter と **InsertBefore の両方を指定することはできません**。</span><span class="sxs-lookup"><span data-stu-id="f99db-158">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f99db-159">この `InsertBefore` 要素は Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-159">The `InsertBefore` element is not available in Outlook.</span></span>

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="f99db-160">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f99db-160">OverriddenByRibbonApi</span></span>

<span data-ttu-id="f99db-161">省略可能 (ブール値)。</span><span class="sxs-lookup"><span data-stu-id="f99db-161">Optional (boolean).</span></span> <span data-ttu-id="f99db-162">カスタム コンテキスト タブをリボンに実行時にインストールする API をサポートするアプリケーションとプラットフォームの組み合わせで **CustomTab** を非表示にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="f99db-162">Specifies whether the **CustomTab** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="f99db-163">既定値 (存在しない場合) は次の値です `false` 。</span><span class="sxs-lookup"><span data-stu-id="f99db-163">The default value, if not present, is `false`.</span></span> <span data-ttu-id="f99db-164">使用する場合 **、OverriddenByRibbonApi は** CustomTab の最初 \*の子\***である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="f99db-164">If used, **OverriddenByRibbonApi** must be the *first* child of **CustomTab**.</span></span> <span data-ttu-id="f99db-165">詳細については [、「OverriddenByRibbonApi」を参照してください](overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="f99db-165">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f99db-166">この `OverriddenByRibbonApi` 要素は Outlook では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f99db-166">The `OverriddenByRibbonApi` element is not available in Outlook.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="f99db-167">CustomTab の例</span><span class="sxs-lookup"><span data-stu-id="f99db-167">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
