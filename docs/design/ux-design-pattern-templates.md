---
title: Office アドイン用 UX 設計パターン
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d903f6cb2c6cad90c07b05303eac6b25a05a4af2
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950419"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="7bb39-102">Office アドイン用 UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="7bb39-102">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="7bb39-103">Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bb39-103">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="7bb39-104">この UX パターンはコンポーネントで構成されています。</span><span class="sxs-lookup"><span data-stu-id="7bb39-104">Our UX patterns are composed of components.</span></span> <span data-ttu-id="7bb39-105">コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="7bb39-105">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="7bb39-106">ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。</span><span class="sxs-lookup"><span data-stu-id="7bb39-106">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="7bb39-107">Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。</span><span class="sxs-lookup"><span data-stu-id="7bb39-107">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="7bb39-108">Fabric を活用して、Office と簡単に統合します。</span><span class="sxs-lookup"><span data-stu-id="7bb39-108">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="7bb39-109">アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7bb39-109">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="7bb39-110">Office と統合する際に、それを保持する機会を探します。</span><span class="sxs-lookup"><span data-stu-id="7bb39-110">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="7bb39-111">スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="7bb39-111">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="7bb39-112">提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。</span><span class="sxs-lookup"><span data-stu-id="7bb39-112">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="7bb39-113">それは、アドインの設計と開発のためのクイック エントリ ポイントと、Microsoft とブランド要素のバランスを実現するためのガイダンスの両方を提供することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="7bb39-113">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="7bb39-114">Microsoft の Fabric 設計言語とパートナー特有のブランドの独自性から得たデザイン要素のバランスを取る、クリーンでモダンなユーザー エクスペリエンスを提供することにより、ユーザーの保持とアドインの採用を向上させることができます。</span><span class="sxs-lookup"><span data-stu-id="7bb39-114">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="7bb39-115">UX パターン テンプレートを使用して、次のことを行います。</span><span class="sxs-lookup"><span data-stu-id="7bb39-115">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="7bb39-116">よくある顧客のシナリオにソリューションとして適用する。</span><span class="sxs-lookup"><span data-stu-id="7bb39-116">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="7bb39-117">設計のベスト プラクティスとして適用する。</span><span class="sxs-lookup"><span data-stu-id="7bb39-117">Apply design best practices.</span></span>
* <span data-ttu-id="7bb39-118">[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。</span><span class="sxs-lookup"><span data-stu-id="7bb39-118">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="7bb39-119">Office の既定の UI に視覚的に溶け込むアドインをビルドする。</span><span class="sxs-lookup"><span data-stu-id="7bb39-119">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="7bb39-120">UX を観念化および可視化する。</span><span class="sxs-lookup"><span data-stu-id="7bb39-120">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="7bb39-121">はじめに</span><span class="sxs-lookup"><span data-stu-id="7bb39-121">Getting started</span></span>

<span data-ttu-id="7bb39-122">パターンは、キーの動作またはアドインに共通のエクスペリエンスによって構成されます。</span><span class="sxs-lookup"><span data-stu-id="7bb39-122">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="7bb39-123">主なグループは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7bb39-123">The main groups are:</span></span>

* [<span data-ttu-id="7bb39-124">最初の実行エクスペリエンス (FRE)</span><span class="sxs-lookup"><span data-stu-id="7bb39-124">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="7bb39-125">認証</span><span class="sxs-lookup"><span data-stu-id="7bb39-125">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="7bb39-126">ナビゲーション</span><span class="sxs-lookup"><span data-stu-id="7bb39-126">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="7bb39-127">ブランド デザイン</span><span class="sxs-lookup"><span data-stu-id="7bb39-127">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="7bb39-128">各グループを参照して、ベスト プラクティスを使ってアドインを設計する方法を理解します。</span><span class="sxs-lookup"><span data-stu-id="7bb39-128">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="7bb39-129">このドキュメント全体を通して表示されている画面例は、**1366x768**の解像度で設計および表示されています。</span><span class="sxs-lookup"><span data-stu-id="7bb39-129">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="7bb39-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="7bb39-130">See also</span></span>

* [<span data-ttu-id="7bb39-131">デザインのツールキット</span><span class="sxs-lookup"><span data-stu-id="7bb39-131">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="7bb39-132">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="7bb39-132">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="7bb39-133">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="7bb39-133">Best practices for developing Office Add-ins</span></span>](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="7bb39-134">Fabric React の使用の開始</span><span class="sxs-lookup"><span data-stu-id="7bb39-134">Get started using Fabric React</span></span>](/office/dev/add-ins/design/using-office-ui-fabric-react)
