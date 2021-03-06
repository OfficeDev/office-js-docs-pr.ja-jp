---
title: 組み込みのコントロール Officeカスタム コントロール グループとタブに統合する
description: カスタム コマンド グループとタブに組み込Officeボタンをリボンに含めるOfficeします。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505256"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a><span data-ttu-id="552bb-103">組み込みのコントロール Officeカスタム コントロール グループとタブに統合する</span><span class="sxs-lookup"><span data-stu-id="552bb-103">Integrate built-in Office buttons into custom control groups and tabs</span></span>

<span data-ttu-id="552bb-104">アドインのマニフェストでマークアップOffice使用して、Office リボンのカスタム コントロール グループに組み込みのコントロール ボタンを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="552bb-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="552bb-105">(カスタム アドイン コマンドは、組み込みのアドイン グループにOfficeできます)。また、組み込みのコントロール グループ全体Officeカスタム リボン タブに挿入することもできます。</span><span class="sxs-lookup"><span data-stu-id="552bb-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="552bb-106">この記事では、アドイン コマンドの基本的な概念に [精通している必要があります](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="552bb-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="552bb-107">最近行っていない場合は、確認してください。</span><span class="sxs-lookup"><span data-stu-id="552bb-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="552bb-108">この記事で説明するアドイン機能とマークアップは *、Web 上の PowerPoint でのみ使用できます*。</span><span class="sxs-lookup"><span data-stu-id="552bb-108">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="552bb-109">この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="552bb-109">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="552bb-110">後のセクション「 [サポートされていないプラットフォームでの動作」を参照してください](#behavior-on-unsupported-platforms)。</span><span class="sxs-lookup"><span data-stu-id="552bb-110">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="552bb-111">組み込みのコントロール グループをカスタム タブに挿入する</span><span class="sxs-lookup"><span data-stu-id="552bb-111">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="552bb-112">組み込みのコントロール グループOfficeタブに挿入するには、親要素に子要素として [OfficeGroup](../reference/manifest/customtab.md#officegroup) 要素を追加 `<CustomTab>` します。</span><span class="sxs-lookup"><span data-stu-id="552bb-112">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="552bb-113">要素 `id` の属性は `<OfficeGroup>` 、組み込みグループの ID に設定されます。</span><span class="sxs-lookup"><span data-stu-id="552bb-113">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="552bb-114">「 [コントロールとコントロール グループの ID を検索する」を参照してください](#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="552bb-114">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="552bb-115">次のマークアップ例は、Office段落コントロール グループをカスタム タブに追加し、カスタム グループの直後に表示する位置を設定します。</span><span class="sxs-lookup"><span data-stu-id="552bb-115">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="552bb-116">組み込みのコントロールをカスタム グループに挿入する</span><span class="sxs-lookup"><span data-stu-id="552bb-116">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="552bb-117">カスタム グループに組み込Officeコントロールを挿入するには、親要素に子要素として [OfficeControl](../reference/manifest/group.md#officecontrol) 要素を追加 `<Group>` します。</span><span class="sxs-lookup"><span data-stu-id="552bb-117">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="552bb-118">要素 `id` の属性 `<OfficeControl>` は、組み込みコントロールの ID に設定されます。</span><span class="sxs-lookup"><span data-stu-id="552bb-118">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="552bb-119">「 [コントロールとコントロール グループの ID を検索する」を参照してください](#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="552bb-119">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="552bb-120">次のマークアップの例では、Office Superscript コントロールをカスタム グループに追加し、カスタム ボタンの直後に表示する位置を設定します。</span><span class="sxs-lookup"><span data-stu-id="552bb-120">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> <span data-ttu-id="552bb-121">ユーザーは、アプリケーションでリボンをOfficeできます。</span><span class="sxs-lookup"><span data-stu-id="552bb-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="552bb-122">ユーザーのカスタマイズは、マニフェスト設定を上書きします。</span><span class="sxs-lookup"><span data-stu-id="552bb-122">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="552bb-123">たとえば、ユーザーは任意のグループからボタンを削除し、タブから任意のグループを削除できます。</span><span class="sxs-lookup"><span data-stu-id="552bb-123">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="552bb-124">コントロールとコントロール グループの ID を検索する</span><span class="sxs-lookup"><span data-stu-id="552bb-124">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="552bb-125">サポートされているコントロールとコントロール グループの ID は、repo コントロールのファイルOffice [に含まれています](https://github.com/OfficeDev/office-control-ids)。</span><span class="sxs-lookup"><span data-stu-id="552bb-125">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="552bb-126">そのレポの ReadMe ファイルの指示に従います。</span><span class="sxs-lookup"><span data-stu-id="552bb-126">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="552bb-127">サポートされていないプラットフォームでの動作</span><span class="sxs-lookup"><span data-stu-id="552bb-127">Behavior on unsupported platforms</span></span>

<span data-ttu-id="552bb-128">アドインが要件セット [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしていないプラットフォームにインストールされている場合、この記事で説明するマークアップは無視され、組み込みの Office コントロール/グループはカスタム グループ/タブに表示されません。</span><span class="sxs-lookup"><span data-stu-id="552bb-128">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="552bb-129">マークアップをサポートしないプラットフォームにアドインがインストールされるのを防ぐには、マニフェストのセクションで要件セットへの参照 `<Requirements>` を追加します。</span><span class="sxs-lookup"><span data-stu-id="552bb-129">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="552bb-130">手順については [、「Set the Requirements element in the manifest」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="552bb-130">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="552bb-131">または [、「JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)コードでランタイム チェックを使用する」の説明に従って、アドインを設計して **、AddinCommands 1.3** がサポートされていない場合に別のエクスペリエンスを提供するように設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="552bb-131">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="552bb-132">たとえば、組み込みボタンがカスタム グループにあると仮定する手順がアドインに含まれている場合は、組み込みボタンが通常の場所にのみ含まれていると仮定する別のバージョンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="552bb-132">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
