---
title: Office アドインの最初の実行エクスペリエンス パターン
description: アドインの初回実行エクスペリエンスを設計するためのベスト プラクティスOffice説明します。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d020a281aca10805ba8fd1176403f3788f6d716c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076344"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="df978-103">最初の実行エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="df978-103">First-run experience patterns</span></span>

<span data-ttu-id="df978-104">最初の実行エクスペリエンス (FRE) は、ユーザーに対するアドインの紹介です。</span><span class="sxs-lookup"><span data-stu-id="df978-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="df978-105">FRE は、ユーザーが初めてアドインを開いた時に表示され、新機能、特徴、および/またはアドインのメリットがわかります。</span><span class="sxs-lookup"><span data-stu-id="df978-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="df978-106">このエクスペリエンスは、ユーザーにアドインを印象付け、継続的にアドオンを使用したり、また使用を再開する、などの可能性を強く形成します。</span><span class="sxs-lookup"><span data-stu-id="df978-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="df978-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="df978-107">Best practices</span></span>

<span data-ttu-id="df978-108">最初の実行エクスペリエンスを作成する際、次のベストプラクティスに従ってください。</span><span class="sxs-lookup"><span data-stu-id="df978-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="df978-109">するべきこと</span><span class="sxs-lookup"><span data-stu-id="df978-109">Do</span></span>|<span data-ttu-id="df978-110">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="df978-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="df978-111">アドインの主な操作を簡単に、短く紹介する。</span><span class="sxs-lookup"><span data-stu-id="df978-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="df978-112">はじめるのに関係のない情報やコールアウトを含めない。</span><span class="sxs-lookup"><span data-stu-id="df978-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="df978-113">アドインの使用にプラスの影響を与える操作を完了する機会をユーザーに与える。</span><span class="sxs-lookup"><span data-stu-id="df978-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="df978-114">ユーザーが一度にすべてを覚えるとは思わない。</span><span class="sxs-lookup"><span data-stu-id="df978-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="df978-115">最も価値を提供する操作に焦点を当てる。</span><span class="sxs-lookup"><span data-stu-id="df978-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="df978-116">ユーザーが完了したいと思うような、魅力的なエクスペリエンスを作成する。</span><span class="sxs-lookup"><span data-stu-id="df978-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="df978-117">ユーザーを強制的に最初の実行エクスペリエンスへ導かない。</span><span class="sxs-lookup"><span data-stu-id="df978-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="df978-118">ユーザーには、最初の実行エクスペリエンスを迂回する選択肢を与える。</span><span class="sxs-lookup"><span data-stu-id="df978-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="df978-119">ユーザーに最初の実行エクスペリエンスを 1 回示すか、定期的に示すかを検討することがシナリオにとって重要かどうかを検討する。</span><span class="sxs-lookup"><span data-stu-id="df978-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="df978-120">たとえば、アドインが周期的にのみ活用される場合、ユーザーはアドインにあまり親しんでいない可能性があり、他の最初の実行エクスペリエンスを体験することにメリットを感じる場合があるかもしれません。</span><span class="sxs-lookup"><span data-stu-id="df978-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="df978-121">該当する場合、次のパターンを適用してアドインの最初の実行エクスペリエンスを作成または向上させましょう。</span><span class="sxs-lookup"><span data-stu-id="df978-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="df978-122">カルーセル</span><span class="sxs-lookup"><span data-stu-id="df978-122">Carousel</span></span>

<span data-ttu-id="df978-123">カルーセルは、ユーザーがアドインを使用する前に、ユーザーに一連の特徴や情報ページを表示します。</span><span class="sxs-lookup"><span data-stu-id="df978-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="df978-124">*図 1.ユーザーがカルーセル フローの先頭ページを進むかスキップできる*</span><span class="sxs-lookup"><span data-stu-id="df978-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![デスクトップ アプリケーション作業ウィンドウの最初の実行エクスペリエンスでのカルーセルの手順 1 をOffice図。](../images/add-in-FRE-step-1.png)

<span data-ttu-id="df978-127">*図 2.カルーセル画面の数を最小限に抑え、メッセージを効果的に伝えるために必要な画面のみ*</span><span class="sxs-lookup"><span data-stu-id="df978-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![デスクトップ アプリケーション作業ウィンドウの最初の実行エクスペリエンスでのカルーセルの手順 2 をOffice図。](../images/add-in-FRE-step-2.png)

<span data-ttu-id="df978-130">*図 3.最初の実行エクスペリエンスを終了する明確なアクション呼び出しを提供する*</span><span class="sxs-lookup"><span data-stu-id="df978-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![デスクトップ アプリケーション作業ウィンドウの最初の実行エクスペリエンスでのカルーセルの手順 3 をOffice図。](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="df978-133">価値プレイスマット</span><span class="sxs-lookup"><span data-stu-id="df978-133">Value Placemat</span></span>

<span data-ttu-id="df978-134">価値プレイスマットは、ロゴの配置、明確に示される価値提案、機能ハイライト、概要、行動を促す言葉などにより、アドインの価値提案を行います。</span><span class="sxs-lookup"><span data-stu-id="df978-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="df978-135">*図 4.ロゴ、明確な値の提案、機能の概要、および行動の呼び出しを含む値の配置*</span><span class="sxs-lookup"><span data-stu-id="df978-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![デスクトップ アプリケーション作業ウィンドウの最初の実行エクスペリエンスでのOfficeを示す図。](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="df978-138">ビデオ プレイスマット</span><span class="sxs-lookup"><span data-stu-id="df978-138">Video Placemat</span></span>

<span data-ttu-id="df978-139">ビデオ プレイスマットはアドインの使用を開始する前に、ユーザーにビデオを表示します。</span><span class="sxs-lookup"><span data-stu-id="df978-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="df978-140">*図 5.最初にビデオの配置を実行する - 画面には、再生ボタンとアクションの呼び出しボタンをクリアしたビデオの画像が含まれている*</span><span class="sxs-lookup"><span data-stu-id="df978-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![デスクトップ アプリケーション作業ウィンドウの最初の実行エクスペリエンスでのビデオ の配置Office図。](../images/add-in-FRE-video.png)

<span data-ttu-id="df978-142">*図 6.ビデオ プレーヤー - ダイアログ ウィンドウ内でビデオを表示するユーザー*</span><span class="sxs-lookup"><span data-stu-id="df978-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![デスクトップ アプリケーションとアドイン作業ウィンドウがバックグラウンドOfficeウィンドウ内のビデオを示す図。](../images/add-in-FRE-video-dialog.png)
