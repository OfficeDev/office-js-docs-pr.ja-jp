---
title: Office アドインの最初の実行エクスペリエンス パターン
description: Office アドインで初回実行時エクスペリエンスを設計するためのベストプラクティスについて説明します。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: c0528e869dd8ee7fe779785fb1a9b6d347deab75
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292954"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="c5ca4-103">最初の実行エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="c5ca4-103">First-run experience patterns</span></span>

<span data-ttu-id="c5ca4-104">最初の実行エクスペリエンス (FRE) は、ユーザーに対するアドインの紹介です。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="c5ca4-105">FRE は、ユーザーが初めてアドインを開いた時に表示され、新機能、特徴、および/またはアドインのメリットがわかります。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="c5ca4-106">このエクスペリエンスは、ユーザーにアドインを印象付け、継続的にアドオンを使用したり、また使用を再開する、などの可能性を強く形成します。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="c5ca4-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c5ca4-107">Best practices</span></span>


<span data-ttu-id="c5ca4-108">最初の実行エクスペリエンスを作成する際、次のベストプラクティスに従ってください。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="c5ca4-109">するべきこと</span><span class="sxs-lookup"><span data-stu-id="c5ca4-109">Do</span></span>|<span data-ttu-id="c5ca4-110">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="c5ca4-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="c5ca4-111">アドインの主な操作を簡単に、短く紹介する。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="c5ca4-112">はじめるのに関係のない情報やコールアウトを含めない。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="c5ca4-113">アドインの使用にプラスの影響を与える操作を完了する機会をユーザーに与える。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="c5ca4-114">ユーザーが一度にすべてを覚えるとは思わない。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="c5ca4-115">最も価値を提供する操作に焦点を当てる。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="c5ca4-116">ユーザーが完了したいと思うような、魅力的なエクスペリエンスを作成する。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="c5ca4-117">ユーザーを強制的に最初の実行エクスペリエンスへ導かない。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="c5ca4-118">ユーザーには、最初の実行エクスペリエンスを迂回する選択肢を与える。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-118">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="c5ca4-119">ユーザーに最初の実行エクスペリエンスを 1 回示すか、定期的に示すかを検討することがシナリオにとって重要かどうかを検討する。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="c5ca4-120">たとえば、アドインが周期的にのみ活用される場合、ユーザーはアドインにあまり親しんでいない可能性があり、他の最初の実行エクスペリエンスを体験することにメリットを感じる場合があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="c5ca4-121">該当する場合、次のパターンを適用してアドインの最初の実行エクスペリエンスを作成または向上させましょう。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="c5ca4-122">カルーセル</span><span class="sxs-lookup"><span data-stu-id="c5ca4-122">Carousel</span></span>


<span data-ttu-id="c5ca4-123">カルーセルは、ユーザーがアドインを使用する前に、ユーザーに一連の特徴や情報ページを表示します。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="c5ca4-124">*図 1: ユーザーがカルーセルフローの最初のページを進めるかスキップできるようにする。* 
 ![最初の実行-カルーセルの手順 1-デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="c5ca4-124">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel Step 1 - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="c5ca4-125">*図 2: ユーザーに対して表示されるカルーセル画面の数を最小限に抑え、メッセージを効果的に伝えるために必要なものだけを指定します。* 
 ![最初の実行-カルーセルの手順 2-デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="c5ca4-125">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message.*
![First Run - Carousel Step 2 - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="c5ca4-126">*図 3: 最初の実行環境を終了するための、アクションを明確に呼び出す方法を提供します。* 
 ![最初の実行-カルーセルの手順 3-デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="c5ca4-126">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel Step 3 - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="c5ca4-127">価値プレイスマット</span><span class="sxs-lookup"><span data-stu-id="c5ca4-127">Value Placemat</span></span>

<span data-ttu-id="c5ca4-128">価値プレイスマットは、ロゴの配置、明確に示される価値提案、機能ハイライト、概要、行動を促す言葉などにより、アドインの価値提案を行います。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-128">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="c5ca4-129">![最初の実行値プレイスマット-デスクトップ作業ウィンドウの仕様 ](../images/add-in-FRE-value.png)
 *。ロゴ付きの値のプレイスマット、価値提案のクリア、機能の概要、アクションの呼び出し。*</span><span class="sxs-lookup"><span data-stu-id="c5ca4-129">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call-to-action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="c5ca4-130">ビデオ プレイスマット</span><span class="sxs-lookup"><span data-stu-id="c5ca4-130">Video Placemat</span></span>

<span data-ttu-id="c5ca4-131">ビデオ プレイスマットはアドインの使用を開始する前に、ユーザーにビデオを表示します。</span><span class="sxs-lookup"><span data-stu-id="c5ca4-131">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="c5ca4-132">*図 1: 最初の実行プレイスマット-画面には、[再生] ボタンと [アクションをクリア] ボタンが付いたビデオの静止画像が含まれています。* 
 ![ビデオプレイスマット-デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="c5ca4-132">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call-to-action button.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="c5ca4-133">*図 2: ビデオプレーヤー-ユーザーには、ダイアログウィンドウ内にビデオが表示されます。* 
 ![ビデオプレイスマット-ダイアログ-デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="c5ca4-133">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Dialog - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
