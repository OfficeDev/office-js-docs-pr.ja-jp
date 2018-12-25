---
title: Office アドインのブランド パターン設計ガイドライン
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: a94e723b222dfe1b004d8b558da59804faf51e69
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433705"
---
# <a name="branding-patterns"></a><span data-ttu-id="14d12-102">ブランド パターン</span><span class="sxs-lookup"><span data-stu-id="14d12-102">Branding patterns</span></span>

<span data-ttu-id="14d12-103">これらのパターンでは、ブランドの視認性とコンテキストをアドイン ユーザーに提供します。</span><span class="sxs-lookup"><span data-stu-id="14d12-103">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="14d12-104">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="14d12-104">Best practices</span></span>

|<span data-ttu-id="14d12-105">するべきこと</span><span class="sxs-lookup"><span data-stu-id="14d12-105">Do</span></span> |<span data-ttu-id="14d12-106">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="14d12-106">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="14d12-107">使い慣れた UI コンポーネントを使用して、タイポグラフィや色などのブランディングのアクセントを適用する。</span><span class="sxs-lookup"><span data-stu-id="14d12-107">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="14d12-108">確立された Office UI と矛盾する新しい UI コンポーネントを考案しない。</span><span class="sxs-lookup"><span data-stu-id="14d12-108">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="14d12-109">UI の下部にあるブランド バーのフッターにアドインのブランドを配置する。</span><span class="sxs-lookup"><span data-stu-id="14d12-109">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="14d12-110">UI の上部にある、隣接するブランド バーで作業ウィンドウ名を繰り返さない。</span><span class="sxs-lookup"><span data-stu-id="14d12-110">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="14d12-111">ブランド要素は控えめに使用する。</span><span class="sxs-lookup"><span data-stu-id="14d12-111">Use brand elements sparingly.</span></span> <span data-ttu-id="14d12-112">ソリューションを補完的な方法で Office に適合させる。</span><span class="sxs-lookup"><span data-stu-id="14d12-112">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="14d12-113">ブランド化された構成要素を過度に Office UI に挿入して、顧客の気をそらしたり混乱させたりしない。</span><span class="sxs-lookup"><span data-stu-id="14d12-113">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="14d12-114">ソリューションを認識できるようにし、一貫したビジュアル要素によって画面を接続する。</span><span class="sxs-lookup"><span data-stu-id="14d12-114">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="14d12-115">認識不能で一貫性のないビジュアル要素を適用して、ソリューションを覆い隠したりしない。</span><span class="sxs-lookup"><span data-stu-id="14d12-115">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="14d12-116">親サービスまたはビジネスとの接続を構築して、顧客がソリューションを確実に認識し信頼できるようにする。</span><span class="sxs-lookup"><span data-stu-id="14d12-116">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="14d12-117">信頼と価値を築くために活用できる有用で理解しやすいリレーションシップがあるなら、顧客に新しいブランド コンセプトを感じさせないようにする。</span><span class="sxs-lookup"><span data-stu-id="14d12-117">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="14d12-118">ユーザーがアドインの完全なユーティリティを使用できるようにするには、次のパターンとコンポーネントを適用します。</span><span class="sxs-lookup"><span data-stu-id="14d12-118">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="14d12-119">ブランド バー</span><span class="sxs-lookup"><span data-stu-id="14d12-119">Brand Bar</span></span>

<span data-ttu-id="14d12-120">ブランド バーは、ブランド名とロゴを含むフッター内のスペースです。</span><span class="sxs-lookup"><span data-stu-id="14d12-120">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="14d12-121">また、あなたのブランドの Web サイトや、オプションのアクセス ロケーションへのリンクとして機能します。</span><span class="sxs-lookup"><span data-stu-id="14d12-121">It also serves as a link to your brand's website and an optional access location.</span></span>

![ブランド バー - デスクトップ作業ウィンドウの仕様](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="14d12-123">スプラッシュ スクリーン</span><span class="sxs-lookup"><span data-stu-id="14d12-123">Splash Screen</span></span>

<span data-ttu-id="14d12-124">アドインの読み込み中や UI の状態が切り替えられる間、この画面を使用してブランドを表示します。</span><span class="sxs-lookup"><span data-stu-id="14d12-124">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![ブランド スプラッシュ スクリーン - デスクトップ作業ウィンドウの仕様](../images/add-in-splash-screen.png)