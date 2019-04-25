---
title: Office アドインの最初の実行エクスペリエンス パターン
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 85f8e4f7e0082e00ad5064333470f589e449af45
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446849"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="fdc98-102">最初の実行エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="fdc98-102">First-run experience patterns</span></span>

<span data-ttu-id="fdc98-103">最初の実行エクスペリエンス (FRE) は、ユーザーに対するアドインの紹介です。</span><span class="sxs-lookup"><span data-stu-id="fdc98-103">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="fdc98-104">FRE は、ユーザーが初めてアドインを開いた時に表示され、新機能、特徴、および/またはアドインのメリットがわかります。</span><span class="sxs-lookup"><span data-stu-id="fdc98-104">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="fdc98-105">このエクスペリエンスは、ユーザーにアドインを印象付け、継続的にアドオンを使用したり、また使用を再開する、などの可能性を強く形成します。</span><span class="sxs-lookup"><span data-stu-id="fdc98-105">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="fdc98-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="fdc98-106">Best practices</span></span>


<span data-ttu-id="fdc98-107">最初の実行エクスペリエンスを作成する際、次のベストプラクティスに従ってください。</span><span class="sxs-lookup"><span data-stu-id="fdc98-107">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="fdc98-108">するべきこと</span><span class="sxs-lookup"><span data-stu-id="fdc98-108">Do</span></span>|<span data-ttu-id="fdc98-109">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="fdc98-109">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="fdc98-110">アドインの主な操作を簡単に、短く紹介する。</span><span class="sxs-lookup"><span data-stu-id="fdc98-110">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="fdc98-111">はじめるのに関係のない情報やコールアウトを含めない。</span><span class="sxs-lookup"><span data-stu-id="fdc98-111">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="fdc98-112">アドインの使用にプラスの影響を与える操作を完了する機会をユーザーに与える。</span><span class="sxs-lookup"><span data-stu-id="fdc98-112">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="fdc98-113">ユーザーが一度にすべてを覚えるとは思わない。</span><span class="sxs-lookup"><span data-stu-id="fdc98-113">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="fdc98-114">最も価値を提供する操作に焦点を当てる。</span><span class="sxs-lookup"><span data-stu-id="fdc98-114">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="fdc98-115">ユーザーが完了したいと思うような、魅力的なエクスペリエンスを作成する。</span><span class="sxs-lookup"><span data-stu-id="fdc98-115">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="fdc98-116">ユーザーを強制的に最初の実行エクスペリエンスへ導かない。</span><span class="sxs-lookup"><span data-stu-id="fdc98-116">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="fdc98-117">ユーザーには、最初の実行エクスペリエンスを迂回する選択肢を与える。</span><span class="sxs-lookup"><span data-stu-id="fdc98-117">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="fdc98-118">ユーザーに最初の実行エクスペリエンスを 1 回示すか、定期的に示すかを検討することがシナリオにとって重要かどうかを検討する。</span><span class="sxs-lookup"><span data-stu-id="fdc98-118">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="fdc98-119">たとえば、アドインが周期的にのみ活用される場合、ユーザーはアドインにあまり親しんでいない可能性があり、他の最初の実行エクスペリエンスを体験することにメリットを感じる場合があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="fdc98-119">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="fdc98-120">該当する場合、次のパターンを適用してアドインの最初の実行エクスペリエンスを作成または向上させましょう。</span><span class="sxs-lookup"><span data-stu-id="fdc98-120">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="fdc98-121">カルーセル</span><span class="sxs-lookup"><span data-stu-id="fdc98-121">Carousel</span></span>


<span data-ttu-id="fdc98-122">カルーセルは、ユーザーがアドインを使用する前に、ユーザーに一連の特徴や情報ページを表示します。</span><span class="sxs-lookup"><span data-stu-id="fdc98-122">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="fdc98-123">*図 1: カルーセル フローで先に進む、または最初のページをスキップできるようにします。* 
![最初の実行 - カルーセル - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="fdc98-123">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="fdc98-124">\*図 2: ユーザーに表示するカルーセル画面の数を効率的にメッセージを伝えるのに必要な最小限の数に抑えます \*
![最初の実行 - カルーセル - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="fdc98-124">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="fdc98-125">*図 3: 最初の実行エクスペリエンスを終了するための、行動を促す明確な言葉を提供します。* 
![最初の実行 - カルーセル - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="fdc98-125">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="fdc98-126">価値プレイスマット</span><span class="sxs-lookup"><span data-stu-id="fdc98-126">Value Placemat</span></span>

<span data-ttu-id="fdc98-127">価値プレイスマットは、ロゴの配置、明確に示される価値提案、機能ハイライト、概要、行動を促す言葉などにより、アドインの価値提案を行います。</span><span class="sxs-lookup"><span data-stu-id="fdc98-127">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="fdc98-128">![最初の実行 - 価値プレイスマット - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-value.png)
*ロゴ、明確な価値提案、機能概要、行動を促す言葉が含まれる価値プレイスマット。*</span><span class="sxs-lookup"><span data-stu-id="fdc98-128">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="fdc98-129">ビデオ プレイスマット</span><span class="sxs-lookup"><span data-stu-id="fdc98-129">Video Placemat</span></span>

<span data-ttu-id="fdc98-130">ビデオ プレイスマットはアドインの使用を開始する前に、ユーザーにビデオを表示します。</span><span class="sxs-lookup"><span data-stu-id="fdc98-130">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="fdc98-131">*図 1: 最初の実行プレイスマット - この画面には、再生ボタンと行動を促す明確な言葉のボタンが含まれる、ビデオの静止画が含まれます。*![ビデオ プレイスマット - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="fdc98-131">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="fdc98-132">*図 2: ビデオ プレーヤー - ユーザーには、ダイアログ ウィンドウ内にビデオが表示されます。* 
![ビデオ プレイスマット - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="fdc98-132">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
