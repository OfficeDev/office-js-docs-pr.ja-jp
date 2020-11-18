---
title: 組み込みの Office ボタンをカスタムコントロールグループとタブに統合する
description: Office リボンのカスタムコマンドグループとタブに組み込みの Office ボタンを含める方法について説明します。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088176"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a><span data-ttu-id="b255d-103">組み込みの Office ボタンをカスタムコントロールグループとタブに統合する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b255d-103">Integrate built-in Office buttons into custom control groups and tabs (preview)</span></span>

<span data-ttu-id="b255d-104">アドインのマニフェストでマークアップを使用すると、office リボンのカスタムコントロールグループに組み込みの Office ボタンを挿入できます。</span><span class="sxs-lookup"><span data-stu-id="b255d-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="b255d-105">(組み込みの Office グループにカスタムアドインコマンドを挿入することはできません。)組み込みの Office コントロールグループのすべてをカスタムのリボンタブに挿入することもできます。</span><span class="sxs-lookup"><span data-stu-id="b255d-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="b255d-106">この記事では、 [アドインコマンドの基本的な概念](add-in-commands.md)について理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="b255d-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="b255d-107">まだ行っていない場合は、確認してください。</span><span class="sxs-lookup"><span data-stu-id="b255d-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="b255d-108">この記事に記載されているアドイン機能とマークアップはプレビュー段階であり、 *PowerPoint on the web でのみ利用でき* ます。</span><span class="sxs-lookup"><span data-stu-id="b255d-108">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="b255d-109">テストおよび開発環境でマークアップを試すことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b255d-109">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="b255d-110">運用環境または業務上重要なドキュメント内では、プレビューマークアップを使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="b255d-110">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="b255d-111">この記事に記載されているマークアップは、要件セット **Addincommands 1.3** をサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="b255d-111">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="b255d-112">サポートされてい [ないプラットフォームで](#behavior-on-unsupported-platforms)は、後のセクションの動作を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b255d-112">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="b255d-113">組み込みのコントロールグループをカスタムタブに挿入する</span><span class="sxs-lookup"><span data-stu-id="b255d-113">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="b255d-114">組み込みの Office コントロールグループをタブに挿入するには、 [Officegroup](../reference/manifest/customtab.md#officegroup) 要素を親要素の子要素として追加し `<CustomTab>` ます。</span><span class="sxs-lookup"><span data-stu-id="b255d-114">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="b255d-115">`id`要素の属性は、 `<OfficeGroup>` 組み込みのグループの ID に設定されます。</span><span class="sxs-lookup"><span data-stu-id="b255d-115">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="b255d-116">「 [コントロールおよびコントロールグループの id を検索する」を](#find-the-ids-of-controls-and-control-groups)参照してください。</span><span class="sxs-lookup"><span data-stu-id="b255d-116">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="b255d-117">次のマークアップの例では、ユーザー設定のタブに Office 段落コントロールグループを追加し、ユーザー設定のグループの直後に表示されるように配置します。</span><span class="sxs-lookup"><span data-stu-id="b255d-117">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="b255d-118">組み込みのコントロールをカスタムグループに挿入する</span><span class="sxs-lookup"><span data-stu-id="b255d-118">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="b255d-119">組み込みの Office コントロールをカスタムグループに挿入するには、親要素の子要素として、 [Officeecontrol](../reference/manifest/group.md#officecontrol) 要素を追加し `<Group>` ます。</span><span class="sxs-lookup"><span data-stu-id="b255d-119">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="b255d-120">`id`要素の属性 `<OfficeControl>` は、組み込みのコントロールの ID に設定されます。</span><span class="sxs-lookup"><span data-stu-id="b255d-120">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="b255d-121">「 [コントロールおよびコントロールグループの id を検索する」を](#find-the-ids-of-controls-and-control-groups)参照してください。</span><span class="sxs-lookup"><span data-stu-id="b255d-121">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="b255d-122">次のマークアップの例では、ユーザー設定のグループに Office の上付きコントロールを追加し、ユーザー設定のボタンの直後に表示されるように配置します。</span><span class="sxs-lookup"><span data-stu-id="b255d-122">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="b255d-123">ユーザーは、Office アプリケーションでリボンをカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="b255d-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="b255d-124">ユーザーのカスタマイズは、マニフェストの設定よりも優先されます。</span><span class="sxs-lookup"><span data-stu-id="b255d-124">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="b255d-125">たとえば、ユーザーは任意のグループからボタンを削除したり、タブから任意のグループを削除したりできます。</span><span class="sxs-lookup"><span data-stu-id="b255d-125">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="b255d-126">コントロールおよびコントロールグループの Id を検索する</span><span class="sxs-lookup"><span data-stu-id="b255d-126">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="b255d-127">サポートされているコントロールおよびコントロールグループの Id は、リポジトリの [Office コントロール id](https://github.com/OfficeDev/office-control-ids)のファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="b255d-127">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="b255d-128">そのリポジトリの ReadMe ファイルの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="b255d-128">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="b255d-129">サポートされていないプラットフォームでの動作</span><span class="sxs-lookup"><span data-stu-id="b255d-129">Behavior on unsupported platforms</span></span>

<span data-ttu-id="b255d-130">[要件セット AddinCommands コマンド 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしていないプラットフォームにアドインがインストールされている場合、この記事に記載されているマークアップは無視され、組み込みの Office コントロール/グループはカスタムグループ/タブに表示されません。</span><span class="sxs-lookup"><span data-stu-id="b255d-130">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="b255d-131">マークアップをサポートしていないプラットフォームにアドインがインストールされないようにするには、マニフェストのセクションにある要件セットへの参照を追加し `<Requirements>` ます。</span><span class="sxs-lookup"><span data-stu-id="b255d-131">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="b255d-132">手順については、「 [マニフェストの要件要素を設定する](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b255d-132">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="b255d-133">または、「 [JavaScript コードでランタイムチェックを使用](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)する」で説明されているように、 **addincommands 1.3** がサポートされていない場合に、アドインの代替操作を実行するようにアドインを設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="b255d-133">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="b255d-134">たとえば、アドインに組み込みのボタンがカスタムグループにあると想定される命令が含まれている場合は、その組み込みボタンが通常の場所にしかないことを前提とした代替バージョンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="b255d-134">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
