---
title: Office アドイン用 UX 設計パターン
description: ナビゲーション、認証、初回実行、ブランド化のパターンなど、Office アドインの UI 設計パターンの概要について説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 164784fcacb8e0869d0c0b8031a71cf0358b03fb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719078"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="f15b7-103">Office アドイン用 UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="f15b7-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="f15b7-104">Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f15b7-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="f15b7-105">この UX パターンはコンポーネントで構成されています。</span><span class="sxs-lookup"><span data-stu-id="f15b7-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="f15b7-106">コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="f15b7-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="f15b7-107">ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。</span><span class="sxs-lookup"><span data-stu-id="f15b7-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="f15b7-108">Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。</span><span class="sxs-lookup"><span data-stu-id="f15b7-108">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="f15b7-109">Fabric を活用して、Office と簡単に統合します。</span><span class="sxs-lookup"><span data-stu-id="f15b7-109">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="f15b7-110">アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="f15b7-110">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="f15b7-111">Office と統合する際に、それを保持する機会を探します。</span><span class="sxs-lookup"><span data-stu-id="f15b7-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="f15b7-112">スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="f15b7-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="f15b7-113">提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。</span><span class="sxs-lookup"><span data-stu-id="f15b7-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="f15b7-114">それは、アドインの設計と開発のためのクイック エントリ ポイントと、Microsoft とブランド要素のバランスを実現するためのガイダンスの両方を提供することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="f15b7-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="f15b7-115">Microsoft の Fabric 設計言語とパートナー特有のブランドの独自性から得たデザイン要素のバランスを取る、クリーンでモダンなユーザー エクスペリエンスを提供することにより、ユーザーの保持とアドインの採用を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="f15b7-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="f15b7-116">UX パターン テンプレートを使用して、次のことを行います。</span><span class="sxs-lookup"><span data-stu-id="f15b7-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="f15b7-117">よくある顧客のシナリオにソリューションとして適用する。</span><span class="sxs-lookup"><span data-stu-id="f15b7-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="f15b7-118">設計のベスト プラクティスとして適用する。</span><span class="sxs-lookup"><span data-stu-id="f15b7-118">Apply design best practices.</span></span>
* <span data-ttu-id="f15b7-119">[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。</span><span class="sxs-lookup"><span data-stu-id="f15b7-119">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="f15b7-120">Office の既定の UI に視覚的に溶け込むアドインをビルドする。</span><span class="sxs-lookup"><span data-stu-id="f15b7-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="f15b7-121">UX を観念化および可視化する。</span><span class="sxs-lookup"><span data-stu-id="f15b7-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="f15b7-122">はじめに</span><span class="sxs-lookup"><span data-stu-id="f15b7-122">Getting started</span></span>

<span data-ttu-id="f15b7-123">パターンは、キーの動作またはアドインに共通のエクスペリエンスによって構成されます。</span><span class="sxs-lookup"><span data-stu-id="f15b7-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="f15b7-124">主なグループは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f15b7-124">The main groups are:</span></span>

* [<span data-ttu-id="f15b7-125">最初の実行エクスペリエンス (FRE)</span><span class="sxs-lookup"><span data-stu-id="f15b7-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="f15b7-126">認証</span><span class="sxs-lookup"><span data-stu-id="f15b7-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="f15b7-127">ナビゲーション</span><span class="sxs-lookup"><span data-stu-id="f15b7-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="f15b7-128">ブランド デザイン</span><span class="sxs-lookup"><span data-stu-id="f15b7-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="f15b7-129">各グループを参照して、ベスト プラクティスを使ってアドインを設計する方法を理解します。</span><span class="sxs-lookup"><span data-stu-id="f15b7-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="f15b7-130">このドキュメント全体を通して表示されている画面例は、**1366x768**の解像度で設計および表示されています。</span><span class="sxs-lookup"><span data-stu-id="f15b7-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="f15b7-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="f15b7-131">See also</span></span>

* [<span data-ttu-id="f15b7-132">デザインのツールキット</span><span class="sxs-lookup"><span data-stu-id="f15b7-132">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="f15b7-133">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="f15b7-133">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="f15b7-134">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="f15b7-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="f15b7-135">Fabric React の使用の開始</span><span class="sxs-lookup"><span data-stu-id="f15b7-135">Get started using Fabric React</span></span>](../design/using-office-ui-fabric-react.md)
