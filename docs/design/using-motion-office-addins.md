---
title: Office アドインでモーションを使用する
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d3be2454b36fe1003c0697f0bca3c29d743e5330
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449059"
---
# <a name="using-motion-in-office-add-ins"></a><span data-ttu-id="410e1-102">Office アドインでモーションを使用する</span><span class="sxs-lookup"><span data-stu-id="410e1-102">Using motion in Office Add-ins</span></span>

<span data-ttu-id="410e1-p101">Office アドインを設計する際、モーションを使用してユーザー エクスペリエンスを向上させられます。 UI 要素、コントロール、コンポーネントには多くの場合、切り替え、モーション、アニメーションを必要とする対話型の動作が関係します。 UI 要素全体においてモーションの共通の特性は、デザイン言語のアニメーション要素を定義することです。</span><span class="sxs-lookup"><span data-stu-id="410e1-p101">When you design an Office Add-in, you can use motion to enhance the user experience. UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language.</span></span> 

<span data-ttu-id="410e1-p102">Office は生産性に重点を置いているため、Office のアニメーション言語は、お客様による業務の遂行を支援するという目的をサポートします。 このアニメーション言語は、優れた応答性、信頼できるビジュアル、きめ細やかな魅力をバランスよく実現しています。 Office に埋め込まれるアドインは、この既存のアニメーション言語を利用します。 したがって、モーションを使用する場合、次のガイドラインを検討することが重要です。</span><span class="sxs-lookup"><span data-stu-id="410e1-p102">Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight. Add-ins embedded in Office sit within this existing animation language. Given this context, it is important to consider the following guidelines when applying motion.</span></span> 


## <a name="create-motion-with-a-purpose"></a><span data-ttu-id="410e1-110">用途に合わせてモーションを作成する</span><span class="sxs-lookup"><span data-stu-id="410e1-110">Create motion with a purpose</span></span>

<span data-ttu-id="410e1-p103">モーションは、ユーザーが価値を実感できるものである必要があります。 アニメーションを選択する際は、コンテンツのトーンと目的を検討します。 重要なメッセージは、探索的ナビゲーションとは異なる方法で処理します。</span><span class="sxs-lookup"><span data-stu-id="410e1-p103">Motion should have a purpose that communicates additional value to the user. Consider the tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigations.</span></span>

<span data-ttu-id="410e1-p104">アドインで使用される標準的な要素では、モーションを組み込むことにより、ユーザーの注意を引く、要素どうしの関連性を示す、ユーザーの操作を確認するなどの用途に役立てられます。 要素にモーションを付けることで、階層モデルやメンタル モデルを明確にできます。</span><span class="sxs-lookup"><span data-stu-id="410e1-p104">Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions. Choreograph elements to reinforce hierarchy and mental models.</span></span>

### <a name="best-practices"></a><span data-ttu-id="410e1-116">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="410e1-116">Best practices</span></span>

|<span data-ttu-id="410e1-117">するべきこと</span><span class="sxs-lookup"><span data-stu-id="410e1-117">Do</span></span>|<span data-ttu-id="410e1-118">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="410e1-118">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="410e1-p105">アドインの中でモーションを設定する必要のある主要な要素を特定します。 アドインの要素では、パネル、オーバーレイ、モーダル、ツール ヒント、メニュー、ティーチング コールアウトに、よくアニメーションが付けられます。</span><span class="sxs-lookup"><span data-stu-id="410e1-p105">Identify key elements in the add-in that should have motion. Commonly animated elements in an add-in are panels, overlays, modals, tool tips, menus, and teaching call outs.</span></span>| <span data-ttu-id="410e1-p106">すべての要素をアニメーション化して、ユーザーを圧倒することは避けてください。 一度に多くの要素に注意を引こうとして、複数のモーションを適用することがないようにします。</span><span class="sxs-lookup"><span data-stu-id="410e1-p106">Don't overwhelm the user by animating every element. Avoid applying multiple motions that attempt to lead or focus the user on many elements at once.</span></span> |
|<span data-ttu-id="410e1-p107">ユーザーが予想できる動作をする、わかりやすく自然なモーションを使用します。要素のトリガー元を検討します。モーションを使用して、操作と結果の UI がつながるようにします。</span><span class="sxs-lookup"><span data-stu-id="410e1-p107">Use simple, subtle motion that behaves in expected ways. Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.</span></span> | <span data-ttu-id="410e1-p108">モーションのための待機時間ができないようにします。 タスクの完了を妨げるモーションは、アドインで使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-p108">Don't create wait time for a motion. Motion in add-ins should not hinder task completion.</span></span>|

![最小限の要素が動いてパネルが開く gif と、たくさんの要素が動いてパネルが開く gif](../images/add-in-motion-purpose.gif)

## <a name="use-expected-motions"></a><span data-ttu-id="410e1-129">予想される動作を使用する</span><span class="sxs-lookup"><span data-stu-id="410e1-129">Use expected motions</span></span>

<span data-ttu-id="410e1-130">[Office UI Fabric](https://developer.microsoft.com/fabric) を使用して、Office プラットフォームと視覚的に関連付けることをお勧めします。また、Fabric モーション言語に合わせてモーションを作成するため、[Fabric アニメーション](https://developer.microsoft.com/fabric#/styles/animations)を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="410e1-130">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric) to create a visual connection with the Office platform, and we also encourage the use of [Fabric Animations](https://developer.microsoft.com/fabric#/styles/animations) to create motions that align with the Fabric motion language.</span></span> 

<span data-ttu-id="410e1-p109">これを、Office とシームレスに適合するように使用します。こうすることにより、直感的なエクスペリエンスを実現できます。アニメーション CSS クラスには、Office のメンタル モデルを明確にするのに役立つ、方向性、開始/終了、期間に関する詳細な設定が用意されており、アドインの操作方法も学べるようになっています。</span><span class="sxs-lookup"><span data-stu-id="410e1-p109">Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.</span></span>

### <a name="best-practices"></a><span data-ttu-id="410e1-134">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="410e1-134">Best practices</span></span>

|<span data-ttu-id="410e1-135">するべきこと</span><span class="sxs-lookup"><span data-stu-id="410e1-135">Do</span></span>|<span data-ttu-id="410e1-136">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="410e1-136">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="410e1-137">Fabric の動作と合うモーションを使用します。</span><span class="sxs-lookup"><span data-stu-id="410e1-137">Use motion that aligns with behaviors in Fabric.</span></span>| <span data-ttu-id="410e1-138">Office の一般的なモーション パターンに干渉または競合するモーションは作成しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-138">Don't create motions that interfere or conflict with common motion patterns in Office.</span></span>
|<span data-ttu-id="410e1-139">同じような要素に対して一貫したアプリケーションの動作があることを確認します。</span><span class="sxs-lookup"><span data-stu-id="410e1-139">Ensure that there is a consistent application of motion across like elements.</span></span>| <span data-ttu-id="410e1-140">同じコンポーネントやオブジェクトのアニメーションに、異なるモーションを使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-140">Don't use different motions to animate the same component or object.</span></span>|
|<span data-ttu-id="410e1-p110">アニメーションの方向にも一貫性があるようにします。 たとえば、右側から開くパネルは、右側に閉じるようにします。</span><span class="sxs-lookup"><span data-stu-id="410e1-p110">Create consistency with use of direction in animation. For example, a panel that opens from the right should close to the right.</span></span>|<span data-ttu-id="410e1-143">要素をアニメーション化する際に、複数の方向を使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-143">Don't animate an element using multiple directions.</span></span>

![予想される方法でモーダルが開く gif と、予想に反する方法でモーダル開く gif](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a><span data-ttu-id="410e1-145">要素に合わないモーションを避ける</span><span class="sxs-lookup"><span data-stu-id="410e1-145">Avoid out of character motion for an element</span></span>

<span data-ttu-id="410e1-p111">モーションを実装する際は、HTML キャンバス (作業ウィンドウ、ダイアログ ボックス、コンテンツ アドイン) のサイズを考慮に入れます。 制約のあるスペースにモーションを詰め込み過ぎないようにします。 要素の動き方は、Office に合わせる必要があります。 アドイン モーションは、高パフォーマンスで、信頼性があり、滑らかなものにする必要があります。 生産性を損なわずに、情報伝達や操作性が向上するようにします。</span><span class="sxs-lookup"><span data-stu-id="410e1-p111">Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion. Avoid overloading in constrained spaces. Moving element(s) should be in tune with Office. The character of add-in motion should be performant, reliable, and fluid. Instead of impeding productivity, aim to inform and direct.</span></span>

### <a name="best-practices"></a><span data-ttu-id="410e1-151">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="410e1-151">Best practices</span></span>

|<span data-ttu-id="410e1-152">するべきこと</span><span class="sxs-lookup"><span data-stu-id="410e1-152">Do</span></span>|<span data-ttu-id="410e1-153">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="410e1-153">Don't</span></span>|
|:-----|:-----|
| <span data-ttu-id="410e1-154">[推奨モーション期間](https://developer.microsoft.com/fabric#/styles/animations)を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="410e1-154">Use [recommended motion durations](https://developer.microsoft.com/fabric#/styles/animations).</span></span> | <span data-ttu-id="410e1-p112">大げさなアニメーションを使用しないでください。 ユーザーの注意をそらす装飾目的のエクスペリエンスは作成しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-p112">Don't use exaggerated animations. Avoid creating experiences that embellish and distract your customers.</span></span>
| <span data-ttu-id="410e1-157">[推奨イージング曲線](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion)に従ってください。</span><span class="sxs-lookup"><span data-stu-id="410e1-157">Follow [recommended easing curves](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span></span>  |<span data-ttu-id="410e1-p113">ぎくしゃくした動きやばらばらな動きは使用しないでください。 期待、バウンス、輪ゴムなどの自然界の物理特性を模倣するだけの効果は使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="410e1-p113">Don't move elements in a jerky or disjointed manner. Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.</span></span>|

![ゆっくりフェードインしてタイルが読み込まれる gif と、バウンスを使用してタイルが読み込まれる gif](../images/add-in-motion-character.gif)

## <a name="see-also"></a><span data-ttu-id="410e1-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="410e1-161">See also</span></span>

* [<span data-ttu-id="410e1-162">Fabric アニメーションのガイドライン</span><span class="sxs-lookup"><span data-stu-id="410e1-162">Fabric animation guidelines</span></span>](https://developer.microsoft.com/fabric#/styles/animations)
* [<span data-ttu-id="410e1-163">ユニバーサル Windows プラットフォーム アプリ用のモーション</span><span class="sxs-lookup"><span data-stu-id="410e1-163">Motion for Universal Windows Platform apps</span></span>](/windows/uwp/design/motion)
