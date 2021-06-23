---
title: Office アドインのブランド パターン設計ガイドライン
description: ビジュアル デザインと互換性をOfficeしながら、アドインをブランド化する方法Office。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: b42d3a722e4f8805e8c03d2e1a5db528a66f1202
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076372"
---
# <a name="branding-patterns"></a><span data-ttu-id="ba11c-103">ブランド パターン</span><span class="sxs-lookup"><span data-stu-id="ba11c-103">Branding patterns</span></span>

<span data-ttu-id="ba11c-104">これらのパターンは、アドイン ユーザーにブランドの可視性とコンテキストを提供します。</span><span class="sxs-lookup"><span data-stu-id="ba11c-104">These patterns provide brand visibility and context to your add-in users.</span></span>

## <a name="best-practices"></a><span data-ttu-id="ba11c-105">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="ba11c-105">Best practices</span></span>

|<span data-ttu-id="ba11c-106">するべきこと</span><span class="sxs-lookup"><span data-stu-id="ba11c-106">Do</span></span> |<span data-ttu-id="ba11c-107">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="ba11c-107">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="ba11c-108">使い慣れた UI コンポーネントを使用して、タイポグラフィや色などのブランディングのアクセントを適用する。</span><span class="sxs-lookup"><span data-stu-id="ba11c-108">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="ba11c-109">確立された Office UI と矛盾する新しい UI コンポーネントを考案しない。</span><span class="sxs-lookup"><span data-stu-id="ba11c-109">Don't invent new UI components that contradict established Office UI.</span></span> |
| <span data-ttu-id="ba11c-110">UI の下部にあるブランド バーのフッターにアドインのブランドを配置する。</span><span class="sxs-lookup"><span data-stu-id="ba11c-110">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="ba11c-111">UI の上部にある、隣接するブランド バーで作業ウィンドウ名を繰り返さない。</span><span class="sxs-lookup"><span data-stu-id="ba11c-111">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="ba11c-112">ブランド要素は控えめに使用する。</span><span class="sxs-lookup"><span data-stu-id="ba11c-112">Use brand elements sparingly.</span></span> <span data-ttu-id="ba11c-113">ソリューションを補完的な方法で Office に適合させる。</span><span class="sxs-lookup"><span data-stu-id="ba11c-113">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="ba11c-114">ブランド化された構成要素を過度に Office UI に挿入して、顧客の気をそらしたり混乱させたりしない。</span><span class="sxs-lookup"><span data-stu-id="ba11c-114">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="ba11c-115">ソリューションを認識できるようにし、一貫したビジュアル要素によって画面を接続する。</span><span class="sxs-lookup"><span data-stu-id="ba11c-115">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="ba11c-116">認識不能で一貫性のないビジュアル要素を適用して、ソリューションを覆い隠したりしない。</span><span class="sxs-lookup"><span data-stu-id="ba11c-116">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="ba11c-117">親サービスまたはビジネスとの接続を構築して、顧客がソリューションを確実に認識し信頼できるようにする。</span><span class="sxs-lookup"><span data-stu-id="ba11c-117">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="ba11c-118">信頼と価値を築くために活用できる有用で理解しやすいリレーションシップがあるなら、顧客に新しいブランド コンセプトを感じさせないようにする。</span><span class="sxs-lookup"><span data-stu-id="ba11c-118">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |

<span data-ttu-id="ba11c-119">ユーザーがアドインの完全なユーティリティを使用できるようにするには、次のパターンとコンポーネントを適用します。</span><span class="sxs-lookup"><span data-stu-id="ba11c-119">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>

## <a name="brand-bar"></a><span data-ttu-id="ba11c-120">ブランド バー</span><span class="sxs-lookup"><span data-stu-id="ba11c-120">Brand Bar</span></span>

<span data-ttu-id="ba11c-121">ブランド バーは、ブランド名とロゴを含むフッター内のスペースです。</span><span class="sxs-lookup"><span data-stu-id="ba11c-121">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="ba11c-122">また、あなたのブランドの Web サイトや、オプションのアクセス ロケーションへのリンクとして機能します。</span><span class="sxs-lookup"><span data-stu-id="ba11c-122">It also serves as a link to your brand's website and an optional access location.</span></span>

![デスクトップ アプリケーションのアドイン作業ウィンドウに表示されるブランド バー Office表示されます。](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="ba11c-124">スプラッシュ スクリーン</span><span class="sxs-lookup"><span data-stu-id="ba11c-124">Splash Screen</span></span>

<span data-ttu-id="ba11c-125">アドインの読み込み中や UI の状態が切り替えられる間、この画面を使用してブランドを表示します。</span><span class="sxs-lookup"><span data-stu-id="ba11c-125">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![デスクトップ アプリケーションのアドイン作業ウィンドウに表示されるブランドスプラッシュOffice表示されます。](../images/add-in-splash-screen.png)
