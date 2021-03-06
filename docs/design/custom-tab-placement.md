---
title: カスタム タブをリボンに配置する
description: カスタム タブがリボンに表示される場所と、既定Officeフォーカスが設定されているかどうかを制御する方法について説明します。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 6718a69191d1d84d96512c01b2544094ce276ab6
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505207"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a><span data-ttu-id="7967d-103">カスタム タブをリボンに配置する</span><span class="sxs-lookup"><span data-stu-id="7967d-103">Position a custom tab on the ribbon</span></span>

<span data-ttu-id="7967d-104">アドインのマニフェストでマークアップを使用して、Office アプリケーションのリボンにアドインのカスタム タブを表示する場所を指定できます。</span><span class="sxs-lookup"><span data-stu-id="7967d-104">You can specify where you want your add-in's custom tab to appear on the Office application's ribbon by using markup in the add-in's manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="7967d-105">この記事では、アドイン コマンドの基本的な概念に [精通している必要があります](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="7967d-105">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="7967d-106">最近行ったことがない場合は、確認してください。</span><span class="sxs-lookup"><span data-stu-id="7967d-106">Please review it if you have not done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="7967d-107">この記事で説明するアドイン機能とマークアップは *、Web 上の PowerPoint でのみ使用できます*。</span><span class="sxs-lookup"><span data-stu-id="7967d-107">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="7967d-108">この記事で説明するマークアップは、要件セット **AddinCommands 1.3** をサポートするプラットフォームでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="7967d-108">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="7967d-109">以下の [「サポートされていないプラットフォームでの動作」を参照](#behavior-on-unsupported-platforms) してください。</span><span class="sxs-lookup"><span data-stu-id="7967d-109">See [Behavior on unsupported platforms](#behavior-on-unsupported-platforms) below.</span></span>

<span data-ttu-id="7967d-110">カスタム タブを表示する場所を指定するには、カスタム タブの横に表示する組み込みの Office タブを特定し、組み込みタブの左側または右側に表示するかどうかを指定します。アドインのマニフェストの[CustomTab](../reference/manifest/customtab.md)要素に[InsertBefore](../reference/manifest/customtab.md#insertbefore) (左) または[InsertAfter](../reference/manifest/customtab.md#insertafter) (右) 要素を含めて、これらの仕様を指定します。</span><span class="sxs-lookup"><span data-stu-id="7967d-110">Specify where you want a custom tab to appear by identifying which built-in Office tab you want it to be next to and specifying whether it should be on the left or right side of the built-in tab. Make these specifications by including either an [InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) or an [InsertAfter](../reference/manifest/customtab.md#insertafter) (right) element in the [CustomTab](../reference/manifest/customtab.md) element of your add-in's manifest.</span></span> <span data-ttu-id="7967d-111">(両方の要素を持つ必要があります)。</span><span class="sxs-lookup"><span data-stu-id="7967d-111">(You cannot have both elements.)</span></span>

<span data-ttu-id="7967d-112">次の例では、カスタム タブが [レビュー] タブの直後 *に表示* するように **構成** されています。要素の値は、組み込みのタブの `<InsertAfter>` ID Office注意してください。</span><span class="sxs-lookup"><span data-stu-id="7967d-112">In the following example, the custom tab is configured to appear *just after* the **Review** tab. Note that the value of the `<InsertAfter>` element is the ID of the built-in Office tab.</span></span> 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

<span data-ttu-id="7967d-113">以下の点を念頭に置いておきます。</span><span class="sxs-lookup"><span data-stu-id="7967d-113">Keep the following points in mind.</span></span>

- <span data-ttu-id="7967d-114">要素  `<InsertBefore>` と  `<InsertAfter>` 要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="7967d-114">The  `<InsertBefore>` and  `<InsertAfter>` elements are optional.</span></span> <span data-ttu-id="7967d-115">どちらも使用しない場合は、カスタム タブがリボンの右端のタブとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-115">If you use neither, then your custom tab will appear as the rightmost tab on the ribbon.</span></span>
- <span data-ttu-id="7967d-116">要素  `<InsertBefore>` と  `<InsertAfter>` 要素は相互に排他的です。</span><span class="sxs-lookup"><span data-stu-id="7967d-116">The  `<InsertBefore>` and  `<InsertAfter>` elements are mutually exclusive.</span></span> <span data-ttu-id="7967d-117">両方を使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="7967d-117">You cannot use both.</span></span>
- <span data-ttu-id="7967d-118">ユーザーが複数のアドインをインストールし、そのユーザー設定タブが同じ場所に構成されている場合は、[確認]タブの後に、最近インストールしたアドインのタブがその場所に配置されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-118">If the user installs more than one add-in whose custom tab is configured for the same place, say after the **Review** tab, then the tab for the most recently installed add-in will be located in that place.</span></span> <span data-ttu-id="7967d-119">以前にインストールしたアドインのタブは、1 か所に移動されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-119">The tabs of the previously installed add-ins will be moved over one place.</span></span> <span data-ttu-id="7967d-120">たとえば、ユーザーはその順序でアドイン A、B、C をインストールし、すべて [レビュー] タブの後にタブを挿入するように構成され、タブは次の順序で表示されます。[レビュー] **、AddinCTab** **、AddinBTab、AddinATab** の順にタブが **表示** されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-120">For example, the user installs add-ins A, B, and C in that order and all are configured to insert a tab after the **Review** tab, then the tabs will appear in this order: **Review**, **AddinCTab**, **AddinBTab**, **AddinATab**.</span></span>
- <span data-ttu-id="7967d-121">ユーザーは、アプリケーションでリボンをOfficeできます。</span><span class="sxs-lookup"><span data-stu-id="7967d-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="7967d-122">たとえば、ユーザーはアドインのタブを移動または非表示にできます。これを防止したり、発生したことを検出したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="7967d-122">For example, a user can move or hide your add-in's tab. You cannot prevent this or detect that it has happened.</span></span>
- <span data-ttu-id="7967d-123">ユーザーが組み込みタブの 1 つを移動すると、Officeタブの既定の場所に関して要素が解釈 `<InsertBefore>` `<InsertAfter>` *されます*。たとえば、ユーザーが [レビュー]タブをリボンの右側に移動した場合、Office は上記の例のマークアップを「既定で [レビュー] タブが表示される場所の右側にカスタム タブを置く」という意味として解釈されます。 \*\*</span><span class="sxs-lookup"><span data-stu-id="7967d-123">If a user moves one of the built-in tabs, then Office interprets the `<InsertBefore>` and  `<InsertAfter>` elements in terms of *the default location of the built-in tab*. For example, if the user moves the **Review** tab to the right end of the ribbon, Office will interpret the markup in the example above as meaning "put the custom tab just to the right of *where the **Review** tab would be by default*."</span></span>

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a><span data-ttu-id="7967d-124">ドキュメントを開く際にフォーカスがあるタブを指定する</span><span class="sxs-lookup"><span data-stu-id="7967d-124">Specifying which tab has focus when the document opens</span></span>

<span data-ttu-id="7967d-125">Office、[ファイル] タブの右側にあるタブに既定のフォーカスが常に **表示** されます。既定では、[ホーム]**タブ** です。[ホーム] タブの前にカスタムタブを構成すると、ドキュメントが開くと、カスタム タブ `<InsertBefore>TabHome</InsertBefore>` にフォーカスが設定されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-125">Office always gives default focus to the tab that is immediately to the right of the **File** tab. By default this is the **Home** tab. If you configure your custom tab to be before the **Home** tab, with `<InsertBefore>TabHome</InsertBefore>`, then your custom tab will have focus when the document opens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7967d-126">アドインの不便さを過度に目立たせ、ユーザーや管理者を悩ませます。</span><span class="sxs-lookup"><span data-stu-id="7967d-126">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span> <span data-ttu-id="7967d-127">ユーザーがドキュメントを操作する主な方法がアドインではない限り、ユーザー設定タブを [ホーム] タブの前に配置しない。</span><span class="sxs-lookup"><span data-stu-id="7967d-127">Do not position a custom tab before the **Home** tab unless your add-in is the primary way users will interact with the document.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="7967d-128">サポートされていないプラットフォームでの動作</span><span class="sxs-lookup"><span data-stu-id="7967d-128">Behavior on unsupported platforms</span></span>

<span data-ttu-id="7967d-129">アドインが要件セット [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)をサポートしないプラットフォームにインストールされている場合、この記事で説明するマークアップは無視され、カスタム タブはリボンの右端のタブとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="7967d-129">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and your custom tab will appear as the rightmost tab on the ribbon.</span></span> <span data-ttu-id="7967d-130">マークアップをサポートしないプラットフォームにアドインがインストールされるのを防ぐには、マニフェストのセクションで要件セットへの参照 `<Requirements>` を追加します。</span><span class="sxs-lookup"><span data-stu-id="7967d-130">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="7967d-131">手順については [、「Set the Requirements element in the manifest」を参照してください](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="7967d-131">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="7967d-132">または [、「JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)コードでランタイム チェックを使用する」の説明に従って、アドインを設計して **、AddinCommands 1.3** がサポートされていない場合に別のエクスペリエンスを提供するように設計することもできます。</span><span class="sxs-lookup"><span data-stu-id="7967d-132">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="7967d-133">たとえば、カスタム タブが必要な場所を想定した手順がアドインに含まれている場合は、タブが右端にあると仮定する別のバージョンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="7967d-133">For example, if your add-in contains instructions that assume the custom tab is where you want it, you could have an alternate version that assumes the tab is the rightmost.</span></span>
