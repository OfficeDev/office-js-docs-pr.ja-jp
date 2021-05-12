---
title: Office アドイン用 UX 設計パターン
description: ナビゲーション、認証、初回実行、ブランド化のパターンなど、Officeアドインの UI デザイン パターンの概要を確認します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8544b56b85a25d522c95546b42a78fe01a3c2586
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330109"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="2557c-103">Office アドイン用 UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="2557c-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="2557c-104">Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2557c-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="2557c-105">この UX パターンはコンポーネントで構成されています。</span><span class="sxs-lookup"><span data-stu-id="2557c-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="2557c-106">コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="2557c-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="2557c-107">ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。</span><span class="sxs-lookup"><span data-stu-id="2557c-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="2557c-108">[Fluent UI Reactコンポーネントは](using-office-ui-fabric-react.md)、JS のフレームワークに依存しないコンポーネントと同様に、Office の一部のように見え、Office UI Fabric[動作します](fabric-core.md)。</span><span class="sxs-lookup"><span data-stu-id="2557c-108">[Fluent UI React components](using-office-ui-fabric-react.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md).</span></span> <span data-ttu-id="2557c-109">いずれかのコンポーネント セットを利用して、複数のコンポーネントと統合Office。</span><span class="sxs-lookup"><span data-stu-id="2557c-109">Take advantage of either set of components to integrate with Office.</span></span> <span data-ttu-id="2557c-110">または、アドインに独自の既存のコンポーネント言語がある場合は、その言語を破棄する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2557c-110">Alternatively, if your add-in has its own preexisting component language, you don't need to discard it.</span></span> <span data-ttu-id="2557c-111">Office と統合する際に、それを保持する機会を探します。</span><span class="sxs-lookup"><span data-stu-id="2557c-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="2557c-112">スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="2557c-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="2557c-113">提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。</span><span class="sxs-lookup"><span data-stu-id="2557c-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="2557c-114">これらは、アドインの設計と開発に関する簡単なエントリ ポイントと、Microsoft ブランド要素と独自のブランド要素のバランスを取るガイダンスの両方を提供することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="2557c-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own.</span></span> <span data-ttu-id="2557c-115">Microsoft の Fluent UI デザイン言語とパートナーの固有のブランド ID のバランスを取る、クリーンでモダンなユーザー エクスペリエンスを提供することで、アドインのユーザー保持と導入が向上する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="2557c-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="2557c-116">UX パターン テンプレートを使用して、次のことを行います。</span><span class="sxs-lookup"><span data-stu-id="2557c-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="2557c-117">よくある顧客のシナリオにソリューションとして適用する。</span><span class="sxs-lookup"><span data-stu-id="2557c-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="2557c-118">設計のベスト プラクティスとして適用する。</span><span class="sxs-lookup"><span data-stu-id="2557c-118">Apply design best practices.</span></span>
* <span data-ttu-id="2557c-119">[Fluent UI コンポーネントとスタイル](https://developer.microsoft.com/fluentui#/get-started)を組み込む。</span><span class="sxs-lookup"><span data-stu-id="2557c-119">Incorporate [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) components and styles.</span></span>
* <span data-ttu-id="2557c-120">Office の既定の UI に視覚的に溶け込むアドインをビルドする。</span><span class="sxs-lookup"><span data-stu-id="2557c-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="2557c-121">UX を観念化および可視化する。</span><span class="sxs-lookup"><span data-stu-id="2557c-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="2557c-122">はじめに</span><span class="sxs-lookup"><span data-stu-id="2557c-122">Getting started</span></span>

<span data-ttu-id="2557c-123">パターンは、キーの動作またはアドインに共通のエクスペリエンスによって構成されます。</span><span class="sxs-lookup"><span data-stu-id="2557c-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="2557c-124">主なグループは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="2557c-124">The main groups are:</span></span>

* [<span data-ttu-id="2557c-125">最初の実行エクスペリエンス (FRE)</span><span class="sxs-lookup"><span data-stu-id="2557c-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="2557c-126">認証</span><span class="sxs-lookup"><span data-stu-id="2557c-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="2557c-127">ナビゲーション</span><span class="sxs-lookup"><span data-stu-id="2557c-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="2557c-128">ブランド デザイン</span><span class="sxs-lookup"><span data-stu-id="2557c-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="2557c-129">各グループを参照して、ベスト プラクティスを使ってアドインを設計する方法を理解します。</span><span class="sxs-lookup"><span data-stu-id="2557c-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="2557c-130">このドキュメント全体を通して表示されている画面例は、**1366x768** の解像度で設計および表示されています。</span><span class="sxs-lookup"><span data-stu-id="2557c-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="2557c-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="2557c-131">See also</span></span>

* [<span data-ttu-id="2557c-132">デザイン ツール キット</span><span class="sxs-lookup"><span data-stu-id="2557c-132">Design tool kits</span></span>](design-toolkits.md)
* [<span data-ttu-id="2557c-133">Fluent UI</span><span class="sxs-lookup"><span data-stu-id="2557c-133">Fluent UI</span></span>](https://developer.microsoft.com/fluentui#)
* [<span data-ttu-id="2557c-134">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="2557c-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="2557c-135">Fluent UI ReactアドインOfficeに含む</span><span class="sxs-lookup"><span data-stu-id="2557c-135">Fluent UI React in Office Add-ins</span></span>](using-office-ui-fabric-react.md)
