---
title: Office アドインでの Office UI Fabric
description: アドインで UI Fabric コンポーネントをOfficeする方法の概要Office説明します。
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 9799d98d795486203e4bcc23bffc043c2ead6e28
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237680"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="d21e4-103">Office アドインでの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="d21e4-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="d21e4-104">Office UI Fabric は、ユーザー エクスペリエンスを構築するための JavaScript フロントエンド フレームワークOffice。</span><span class="sxs-lookup"><span data-stu-id="d21e4-104">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="d21e4-105">Fabric は、拡張や改訂が可能な視覚効果に焦点を合わせたコンポーネントであり、Office アドインで使用できます。</span><span class="sxs-lookup"><span data-stu-id="d21e4-105">Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in.</span></span> <span data-ttu-id="d21e4-106">Fabric は Office デザイン言語を使用するため、Fabric の UX コンポーネントは Office に元々組み込まれているかのように自然に使うことができます。</span><span class="sxs-lookup"><span data-stu-id="d21e4-106">Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="d21e4-p102">アドインを作成する場合は、Office UI Fabric を使用してユーザー エクスペリエンスを作成することをお勧めします。Office UI Fabric の使用は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="d21e4-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="d21e4-109">次のセクションでは、Fabric を使用して要件を満たす方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="d21e4-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="d21e4-110">Fabric Core を使用する: アイコン、フォント、色</span><span class="sxs-lookup"><span data-stu-id="d21e4-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="d21e4-111">Fabric Core には、デザイン言語の基本的な要素 (アイコン、色、タイプ、グリッドなど) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d21e4-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="d21e4-112">Fabric Core は独立したフレームワークです。</span><span class="sxs-lookup"><span data-stu-id="d21e4-112">Fabric core is framework independent.</span></span> <span data-ttu-id="d21e4-113">Fabric Core は、Fabric React によって使用され、Fabric React に含まれます。</span><span class="sxs-lookup"><span data-stu-id="d21e4-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="d21e4-114">Fabric Core の使用を開始するには:</span><span class="sxs-lookup"><span data-stu-id="d21e4-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="d21e4-115">ページの HTML に CDN 参照を追加します。</span><span class="sxs-lookup"><span data-stu-id="d21e4-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="d21e4-116">Fabric のアイコンとフォントを使用します。</span><span class="sxs-lookup"><span data-stu-id="d21e4-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="d21e4-p104">Fabric のアイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。アイコンのサイズは、フォント サイズを変更することで制御できます。たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d21e4-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="d21e4-p105">その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://developer.microsoft.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。</span><span class="sxs-lookup"><span data-stu-id="d21e4-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="d21e4-122">Office UI Fabric で使用可能なフォントのサイズと色については、「[文字体裁](https://developer.microsoft.com/fabric#/styles/typography)」および「[色](https://developer.microsoft.com/fabric#/styles/colors)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d21e4-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="d21e4-123">Fabric コンポーネントを使用する</span><span class="sxs-lookup"><span data-stu-id="d21e4-123">Use Fabric Components</span></span>

<span data-ttu-id="d21e4-124">Fabric は、アドインの構築に使用できるさまざまな UX コンポーネントを提供します。</span><span class="sxs-lookup"><span data-stu-id="d21e4-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="d21e4-125">すべてのファブリック コンポーネントが 1 つのアドインで使用されるとは予想されません。</span><span class="sxs-lookup"><span data-stu-id="d21e4-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="d21e4-126">シナリオとユーザー エクスペリエンスに最適なコンポーネントを決定します (たとえば、作業ウィンドウに [階層](https://developer.microsoft.com/fabric#/components/breadcrumb) リンクを適切に表示するのは難しい場合があります)。</span><span class="sxs-lookup"><span data-stu-id="d21e4-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="d21e4-127">アドインでの使用をお勧めする [一](https://developer.microsoft.com/fluentui#/controls/web) 般的な Fabric React UX コンポーネントの一覧を次に示します。</span><span class="sxs-lookup"><span data-stu-id="d21e4-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="d21e4-128">Button</span><span class="sxs-lookup"><span data-stu-id="d21e4-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="d21e4-129">Checkbox</span><span class="sxs-lookup"><span data-stu-id="d21e4-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="d21e4-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="d21e4-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="d21e4-131">Dropdown</span><span class="sxs-lookup"><span data-stu-id="d21e4-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="d21e4-132">Label</span><span class="sxs-lookup"><span data-stu-id="d21e4-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="d21e4-133">List</span><span class="sxs-lookup"><span data-stu-id="d21e4-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="d21e4-134">Pivot</span><span class="sxs-lookup"><span data-stu-id="d21e4-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="d21e4-135">TextField</span><span class="sxs-lookup"><span data-stu-id="d21e4-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="d21e4-136">Toggle</span><span class="sxs-lookup"><span data-stu-id="d21e4-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="d21e4-p107">アドインの作成には、Angular や React など別の JavaScript フレームワークも使用できます。フレームワークで Fabric コンポーネントを使用するには、次のリソースを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d21e4-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="d21e4-139">**フレームワーク**</span><span class="sxs-lookup"><span data-stu-id="d21e4-139">**Framework**</span></span>|<span data-ttu-id="d21e4-140">**例**</span><span class="sxs-lookup"><span data-stu-id="d21e4-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="d21e4-141">**React**</span><span class="sxs-lookup"><span data-stu-id="d21e4-141">**React**</span></span>|[<span data-ttu-id="d21e4-142">Office アドインで Office UI Fabric React を使用する</span><span class="sxs-lookup"><span data-stu-id="d21e4-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="d21e4-143">**Angular**</span><span class="sxs-lookup"><span data-stu-id="d21e4-143">**Angular**</span></span>| [<span data-ttu-id="d21e4-144">Fabric コンポーネントと Angular 2 コンポーネントの折り返しを検討する</span><span class="sxs-lookup"><span data-stu-id="d21e4-144">Consider wrapping Fabric components with Angular 2 components</span></span>](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
