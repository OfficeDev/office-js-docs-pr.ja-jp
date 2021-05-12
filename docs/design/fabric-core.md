---
title: ファブリック コア (Office アドイン)
description: このアドインで Fabric Core および Fabric UI コンポーネントを使用する方法のOffice説明します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330203"
---
# <a name="fabric-core-in-office-add-ins"></a><span data-ttu-id="672b7-103">ファブリック コア (Office アドイン)</span><span class="sxs-lookup"><span data-stu-id="672b7-103">Fabric Core in Office Add-ins</span></span>

<span data-ttu-id="672b7-104">Fabric Core は、CSS クラスと SASS mixins のオープンソース コレクションであり、このコレクションは、アドイン以外のアドインで使用React Officeです。Fabric Core には、アイコン、色、書体、グリッドなどの Fluent UI デザイン言語の基本的な要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="672b7-104">Fabric Core is an open-source collection of CSS classes and SASS mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids.</span></span> <span data-ttu-id="672b7-105">Fabric Core はフレームワークに依存しないので、任意の単一ページ アプリケーションまたは任意のサーバー側 Web UI フレームワークで使用できます。</span><span class="sxs-lookup"><span data-stu-id="672b7-105">Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework.</span></span> <span data-ttu-id="672b7-106">(歴史的な理由から、"Fluent Core" の代わりに "Fabric Core" と呼ばれる)</span><span class="sxs-lookup"><span data-stu-id="672b7-106">(It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)</span></span>

<span data-ttu-id="672b7-107">アドインの UI が Reactベースでない場合は、一連の非カスタム コンポーネントReactできます。</span><span class="sxs-lookup"><span data-stu-id="672b7-107">If your add-in's UI is not React-based, you can also make use of a set of non-React components.</span></span> <span data-ttu-id="672b7-108">「USE [Office UI Fabric JS コンポーネント」を参照してください](#use-office-ui-fabric-js-components)。</span><span class="sxs-lookup"><span data-stu-id="672b7-108">See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).</span></span>

> [!NOTE]
> <span data-ttu-id="672b7-109">この記事では、アドインのコンテキストでの Fabric Core のOffice説明します。ただし、さまざまなアプリや拡張機能でもMicrosoft 365使用されます。</span><span class="sxs-lookup"><span data-stu-id="672b7-109">This article describes the use of Fabric Core in the context of Office Add-ins. But it's also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="672b7-110">詳細については[、「Fabric Core」](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core)および「Open source repo Office UI Fabric [Core」を参照してください](https://github.com/OfficeDev/office-ui-fabric-core)。</span><span class="sxs-lookup"><span data-stu-id="672b7-110">For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="672b7-111">Fabric Core を使用する: アイコン、フォント、色</span><span class="sxs-lookup"><span data-stu-id="672b7-111">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="672b7-112">Fabric Core の使用を開始するには:</span><span class="sxs-lookup"><span data-stu-id="672b7-112">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="672b7-113">ページの HTML に CDN 参照を追加します。</span><span class="sxs-lookup"><span data-stu-id="672b7-113">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="672b7-114">Fabric Core のアイコンとフォントを使用します。</span><span class="sxs-lookup"><span data-stu-id="672b7-114">Use Fabric Core icons and fonts.</span></span>

    <span data-ttu-id="672b7-115">Fabric Core アイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。</span><span class="sxs-lookup"><span data-stu-id="672b7-115">To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes.</span></span> <span data-ttu-id="672b7-116">アイコンのサイズは、フォント サイズを変更することで制御できます。</span><span class="sxs-lookup"><span data-stu-id="672b7-116">You can control the size of the icon by changing the font size.</span></span> <span data-ttu-id="672b7-117">たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="672b7-117">For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="672b7-118">詳細な手順については [、「Fluent UI アイコン」を参照してください](https://developer.microsoft.com/fluentui#/styles/web/icons)。</span><span class="sxs-lookup"><span data-stu-id="672b7-118">For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span></span> <span data-ttu-id="672b7-119">Fabric Core で使用可能なアイコンを見つけるには、そのページの検索機能を使用します。</span><span class="sxs-lookup"><span data-stu-id="672b7-119">To find more icons that are available in Fabric Core, use the search feature on that page.</span></span> <span data-ttu-id="672b7-120">アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。</span><span class="sxs-lookup"><span data-stu-id="672b7-120">When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="672b7-121">Fabric Core で使用できるフォント サイズと色の詳細については、「色」の [「Typography」](https://developer.microsoft.com/fluentui#/styles/web/typography) および **「Colors」** の目次を参照 [してください](https://developer.microsoft.com/fluentui#/styles/web/colors)。</span><span class="sxs-lookup"><span data-stu-id="672b7-121">For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span></span>

<span data-ttu-id="672b7-122">例については、この記事の [後半の「サンプル](#samples) 」に含まれています。</span><span class="sxs-lookup"><span data-stu-id="672b7-122">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="use-office-ui-fabric-js-components"></a><span data-ttu-id="672b7-123">JS Office UI Fabricを使用する</span><span class="sxs-lookup"><span data-stu-id="672b7-123">Use Office UI Fabric JS components</span></span>

<span data-ttu-id="672b7-124">非カスタム REACT のアドインでは[、Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js)の多くのコンポーネント (ボタン、ダイアログ、ピッカーなど) を使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="672b7-124">Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more.</span></span> <span data-ttu-id="672b7-125">手順については、repo の readme を参照してください。</span><span class="sxs-lookup"><span data-stu-id="672b7-125">See the readme of the repo for instructions.</span></span>

<span data-ttu-id="672b7-126">例については、この記事の [後半の「サンプル](#samples) 」に含まれています。</span><span class="sxs-lookup"><span data-stu-id="672b7-126">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="samples"></a><span data-ttu-id="672b7-127">サンプル</span><span class="sxs-lookup"><span data-stu-id="672b7-127">Samples</span></span>

<span data-ttu-id="672b7-128">次のサンプル アドインでは、Fabric Core または JS コンポーネントOffice UI Fabric使用します。</span><span class="sxs-lookup"><span data-stu-id="672b7-128">The following sample add-ins use Fabric Core and/or Office UI Fabric JS components.</span></span> <span data-ttu-id="672b7-129">これらのリポジトリの一部はアーカイブ済みであり、バグやセキュリティ修正プログラムで更新されなくなりましたが、それらを使用して Fabric Core および Fabric UI コンポーネントの使い方を学習できます。</span><span class="sxs-lookup"><span data-stu-id="672b7-129">Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.</span></span>

- [<span data-ttu-id="672b7-130">Excelアドイン JavaScript SalesTracker</span><span class="sxs-lookup"><span data-stu-id="672b7-130">Excel Add-in JavaScript SalesTracker</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [<span data-ttu-id="672b7-131">Excelアドイン SalesLeads</span><span class="sxs-lookup"><span data-stu-id="672b7-131">Excel Add-in SalesLeads</span></span>](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [<span data-ttu-id="672b7-132">Excelアドイン WoodGrove 経費の傾向</span><span class="sxs-lookup"><span data-stu-id="672b7-132">Excel Add-in WoodGrove Expense Trends</span></span>](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [<span data-ttu-id="672b7-133">Excelコンテンツ アドイン Humongous Insurance</span><span class="sxs-lookup"><span data-stu-id="672b7-133">Excel Content Add-in Humongous Insurance</span></span>](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [<span data-ttu-id="672b7-134">Officeアドイン ファブリック UI のサンプル</span><span class="sxs-lookup"><span data-stu-id="672b7-134">Office Add-in Fabric UI Sample</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="672b7-135">Office-Add-in-UX-Design-Patterns-Code</span><span class="sxs-lookup"><span data-stu-id="672b7-135">Office-Add-in-UX-Design-Patterns-Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="672b7-136">Outlookアドイン GifMe</span><span class="sxs-lookup"><span data-stu-id="672b7-136">Outlook Add-in GifMe</span></span>](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [<span data-ttu-id="672b7-137">PowerPointアドイン Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="672b7-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="672b7-138">Word アドイン Angular2 StyleChecker</span><span class="sxs-lookup"><span data-stu-id="672b7-138">Word Add-in Angular2 StyleChecker</span></span>](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [<span data-ttu-id="672b7-139">Word アドイン JS Redact</span><span class="sxs-lookup"><span data-stu-id="672b7-139">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="672b7-140">Word アドイン MarkdownConversion</span><span class="sxs-lookup"><span data-stu-id="672b7-140">Word Add-in MarkdownConversion</span></span>](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
