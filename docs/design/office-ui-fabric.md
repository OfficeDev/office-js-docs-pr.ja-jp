---
title: Office アドインでの Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 04964d5864eea4a960f7b57e5df6f7bd7c844fde
ms.sourcegitcommit: 4e4f7c095e8f33b06bd8a02534ee901125eb1d17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/28/2018
ms.locfileid: "20084071"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="40874-102">Office アドインでの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="40874-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="40874-p101">Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスをビルドするための JavaScript フロントエンドのフレームワークです。Fabric は、拡張や改訂が可能な視覚効果に焦点を合わせたコンポーネントであり、Office アドインで使用できます。Fabric は Office デザイン言語を使用するため、Fabric の UX コンポーネントは Office に元々組み込まれているかのように自然に使うことができます。</span><span class="sxs-lookup"><span data-stu-id="40874-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="40874-p102">アドインを作成する場合は、Office UI Fabric を使用してユーザー エクスペリエンスを作成することをお勧めします。Office UI Fabric の使用は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="40874-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="40874-108">次のセクションでは、Fabric を使用して要件を満たす方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="40874-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="40874-109">Fabric Core を使用する: アイコン、フォント、色</span><span class="sxs-lookup"><span data-stu-id="40874-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="40874-p103">Fabric Core には、デザイン言語の基本的な要素 (アイコン、色、タイプ、グリッドなど) が含まれます。Fabric Core は独立したフレームワークです。Fabric React と Fabric JS は、どちらも Fabric Core を使用します。</span><span class="sxs-lookup"><span data-stu-id="40874-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="40874-113">Fabric Core の使用を開始するには</span><span class="sxs-lookup"><span data-stu-id="40874-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="40874-114">ページの HTML に CDN 参照を追加します。</span><span class="sxs-lookup"><span data-stu-id="40874-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="40874-115">Fabric のアイコンとフォントを使用します。</span><span class="sxs-lookup"><span data-stu-id="40874-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="40874-p104">Fabric のアイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。アイコンのサイズは、フォント サイズを変更することで制御できます。たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="40874-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="40874-p105">その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://dev.office.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。</span><span class="sxs-lookup"><span data-stu-id="40874-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="40874-121">Office UI Fabric で使用可能なフォントのサイズと色については、「[文字体裁](https://dev.office.com/fabric#/styles/typography)」および「[色](https://dev.office.com/fabric#/styles/colors)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="40874-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="40874-122">Fabric コンポーネントを使用する</span><span class="sxs-lookup"><span data-stu-id="40874-122">Use Fabric Components</span></span> 
<span data-ttu-id="40874-123">Fabric には、次のタイプのコンポーネントを含む、さまざまな UX コンポーネントが用意されています。これらを使用してアドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="40874-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="40874-124">入力コンポーネント - 例: ボタン、チェック ボックス、および切り替え</span><span class="sxs-lookup"><span data-stu-id="40874-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="40874-125">移動コンポーネント - 例: コアドキュメントおよびパンくずリスト</span><span class="sxs-lookup"><span data-stu-id="40874-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="40874-126">通知コンポーネント - 例: MessageBar および Callout</span><span class="sxs-lookup"><span data-stu-id="40874-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="40874-127">すべての Fabric コンポーネントがアドインでの使用を推奨しているわけではありません。アドインでの使用を推奨する Fabric React UX コンポーネントのリストを以下に示します。</span><span class="sxs-lookup"><span data-stu-id="40874-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="40874-128">パンくずリスト</span><span class="sxs-lookup"><span data-stu-id="40874-128">Breadcrumb</span></span>](https://developer.microsoft.com/en-us/fabric#/components/breadcrumb)
- [<span data-ttu-id="40874-129">ボタン</span><span class="sxs-lookup"><span data-stu-id="40874-129">Button</span></span>](https://developer.microsoft.com/en-us/fabric#/components/button)
- [<span data-ttu-id="40874-130">チェックボックス</span><span class="sxs-lookup"><span data-stu-id="40874-130">Checkbox</span></span>](https://developer.microsoft.com/en-us/fabric#/components/checkbox)
- [<span data-ttu-id="40874-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="40874-131">ChoiceGroup</span></span>](https://developer.microsoft.com/en-us/fabric#/components/choicegroup)
- [<span data-ttu-id="40874-132">ドロップダウン</span><span class="sxs-lookup"><span data-stu-id="40874-132">Dropdown</span></span>](https://developer.microsoft.com/en-us/fabric#/components/dropdown)
- [<span data-ttu-id="40874-133">ラベル</span><span class="sxs-lookup"><span data-stu-id="40874-133">Label</span></span>](https://developer.microsoft.com/en-us/fabric#/components/label)
- [<span data-ttu-id="40874-134">リスト</span><span class="sxs-lookup"><span data-stu-id="40874-134">List</span></span>](https://developer.microsoft.com/en-us/fabric#/components/list)
- [<span data-ttu-id="40874-135">コアドキュメント</span><span class="sxs-lookup"><span data-stu-id="40874-135">Pivot</span></span>](https://developer.microsoft.com/en-us/fabric#/components/pivot)
- [<span data-ttu-id="40874-136">テキスト フィールド</span><span class="sxs-lookup"><span data-stu-id="40874-136">TextField</span></span>](https://developer.microsoft.com/en-us/fabric#/components/textfield)
- [<span data-ttu-id="40874-137">切り替え</span><span class="sxs-lookup"><span data-stu-id="40874-137">Toggle</span></span>](https://developer.microsoft.com/en-us/fabric#/components/toggle)

<span data-ttu-id="40874-p106">アドインの作成には、Angular や React など別の JavaScript フレームワークも使用できます。フレームワークで Fabric コンポーネントを使用するには、次のリソースを参照してください。</span><span class="sxs-lookup"><span data-stu-id="40874-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="40874-140">**フレームワーク**</span><span class="sxs-lookup"><span data-stu-id="40874-140">**Framework**</span></span>|<span data-ttu-id="40874-141">**例**</span><span class="sxs-lookup"><span data-stu-id="40874-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="40874-142">**React**</span><span class="sxs-lookup"><span data-stu-id="40874-142">**React**</span></span>|[<span data-ttu-id="40874-143">Office アドインで Office UI Fabric React を使用する</span><span class="sxs-lookup"><span data-stu-id="40874-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="40874-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="40874-144">**Angular**</span></span>| <span data-ttu-id="40874-145">Angular 1.5 ディレクティブのコミュニティ プロジェクトである「[ngOfficeUIFabric](http://ngofficeuifabric.com/)」と、「[Fabric コンポーネントと Angular 2 コンポーネントとのラッピングについて検討する](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="40874-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
