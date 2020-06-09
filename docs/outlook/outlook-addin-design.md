---
title: Outlook アドインの設計
description: Windows、Web、iOS、Mac、Android 上の Outlook にアプリを最適な方法で取り込むための魅力的なアドインを設計、作成するのに役立つガイドラインです。
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: ed2ffe1b46ba4673dea531450a0452afa8de11c5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44606526"
---
# <a name="outlook-add-in-design-guidelines"></a><span data-ttu-id="d3b42-103">Outlook アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="d3b42-103">Outlook add-in design guidelines</span></span>

<span data-ttu-id="d3b42-p101">アドインは、パートナーが、コア機能セットを超えて Outlook の機能を拡張する優れた手段になります。アドインを使用すると、ユーザーは受信トレイから移動することなく、サード パーティのエクスペリエンス、タスク、コンテンツを利用できます。Outlook アドインを一度インストールすると、あらゆるプラットフォームとデバイスで使用できます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p101">Add-ins are a great way for partners to extend the functionality of Outlook beyond our core feature set. Add-ins enable users to access third-party experiences, tasks, and content without needing to leave their inbox. Once installed, Outlook add-ins are available on every platform and device.</span></span>  

<span data-ttu-id="d3b42-107">以下に示す基本ガイドラインは、Windows、Web、iOS、Mac、Android 上の Outlook にアプリを最適な方法で取り込むための魅力的なアドインを設計、作成するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-107">The following high-level guidelines will help you design and build a compelling add-in, which brings the best of your app right into Outlook&mdash;on Windows, Web, iOS, Mac, and Android.</span></span>

## <a name="principles"></a><span data-ttu-id="d3b42-108">原則</span><span class="sxs-lookup"><span data-stu-id="d3b42-108">Principles</span></span>

1. <span data-ttu-id="d3b42-109">**いくつかの主要なタスクに重点を置き、それらのタスクを適切に実行できるようにする**</span><span class="sxs-lookup"><span data-stu-id="d3b42-109">**Focus on a few key tasks; do them well**</span></span>

   <span data-ttu-id="d3b42-p102">最適な設計が施されたアドインは、ユーザーが簡単に使用でき、目的が明確で、実際の価値をユーザーに提供します。アドインは Outlook 内で実行されるため、この原則にはより重点が置かれています。Outlook は生産性アプリで、ユーザーが作業する場所となります。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p102">The best designed add-ins are simple to use, focused, and provide real value to users. Because your add-in will run inside of Outlook, there is additional emphasis placed on this principle. Outlook is a productivity app&mdash;it's where people go to get things done.</span></span>

   <span data-ttu-id="d3b42-p103">Microsoft が提供するエクスペリエンスを拡張することになるため、Outlook 内において自然な仕方で調和するシナリオにすることが重要となります。一般的なユース ケースについて、電子メールや予定表を使用する際に最もメリットが多いものを注意深く検討します。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p103">You will be an extension of our experience and it is important to make sure the scenarios you enable feel like a natural fit inside of Outlook. Think carefully about which of your common use cases will benefit the most from having hooks to them from within our email and calendaring experiences.</span></span>

   <span data-ttu-id="d3b42-p104">1 つのアドインで、アプリが実行するすべての処理を行おうとしないでください。Outlook コンテンツにおいて最も頻繁に使用し、適切なアクションであることに重点を置く必要があります。アクションのきっかけとなる事柄についてよく考えて、作業ウィンドウが開くときにユーザーが実行すべき事柄を明確にしてください。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p104">An add-in should not attempt to do everything your app does. The focus should be on the most frequently used, and appropriate, actions in the context of Outlook content. Think about your call to action and make it clear what the user should do when your task pane opens.</span></span>

2. <span data-ttu-id="d3b42-118">**可能な限りネイティブであると感じるようにする**</span><span class="sxs-lookup"><span data-stu-id="d3b42-118">**Make it feel as native as possible**</span></span>

   <span data-ttu-id="d3b42-p105">アドインは、Outlook を実行するプラットフォームにネイティブなパターンを使用して設計する必要があります。そのためには、各プラットフォームで定められている相互作用および視覚に関するガイドラインに従って実装してください。Outlook には独自のガイドラインがあり、それを考慮に入れることも重要です。独自のエクスペリエンス、プラットフォーム、Outlook の 3 つを適切に組み合わせたアドインが、適切に設計されたアドインと言えます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p105">Your add-in should be designed using patterns native to the platform that Outlook is running on. To achieve this, be sure to respect and implement the interaction and visual guidelines set forth by each platform. Outlook has its own guidelines and those are also important to consider. A well-designed add-in will be an appropriate blend of your experience, the platform, and Outlook.</span></span>

   <span data-ttu-id="d3b42-p106">これは、アドインを Outlook on iOS で実行するときと Outlook on Android で実行するときでは視覚的に異なっている必要があるということです。スタイルの 1 例として「[Framework7](https://framework7.io/)」をご覧になってみてください。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p106">This does mean that your add-in will have to visually be different when it runs in Outlook on iOS versus Android. We recommend taking a look at [Framework7](https://framework7.io/) as one option to help you with styling.</span></span>

3. <span data-ttu-id="d3b42-125">**楽しく使用できるようにし、詳細な点に気を配る**</span><span class="sxs-lookup"><span data-stu-id="d3b42-125">**Make it enjoyable to use and get the details right**</span></span>

   <span data-ttu-id="d3b42-p107">機能的および視覚的に魅力的な製品は使用していて楽しいものです。すべての相互作用と視覚上の詳細を注意深く考慮してエクスペリエンスを作り上げると、優れたアドインを実現できます。特定のタスクを実行するために必要な手順を明確にし、関連付ける必要があります。すべてのアクションを 1、2 回のクリックだけで実行できるようにすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p107">People enjoy using products that are both functionally and visually appealing. You can help ensure the success of your add-in by crafting an experience where you've carefully considered every interaction and visual detail. The necessary steps to complete a task must be clear and relevant. Ideally, no action should be further than a click or two away.</span></span> 
   
   <span data-ttu-id="d3b42-130">ユーザーが特定のアクションを行うために、実行中の操作を中断することがないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="d3b42-130">Try not to take a user out of context to complete an action.</span></span> <span data-ttu-id="d3b42-131">ユーザーは、アドインを簡単に出入りし、作業に戻ることができる必要があります。</span><span class="sxs-lookup"><span data-stu-id="d3b42-131">A user should easily be able to get in and out of your add-in and back to whatever she was doing before.</span></span> <span data-ttu-id="d3b42-132">アドインによって、多くの時間が奪われるのではなく、コア機能を拡張することが目的です。</span><span class="sxs-lookup"><span data-stu-id="d3b42-132">An add-in is not meant to be a destination to spend a lot of time in&mdash;it is an enhancement to our core functionality.</span></span> <span data-ttu-id="d3b42-133">適切に設計されたアドインを使用すると、ユーザーの生産性を向上するという目標を達成できます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-133">If done properly, your add-in will help us deliver on the goal of making people more productive.</span></span>

4. <span data-ttu-id="d3b42-134">**賢明な方法でブランド化する**</span><span class="sxs-lookup"><span data-stu-id="d3b42-134">**Brand wisely**</span></span>

   <span data-ttu-id="d3b42-135">ブランド化には大きな価値があり、ユーザーに固有のエクスペリエンスを提供することは重要だと感じています。</span><span class="sxs-lookup"><span data-stu-id="d3b42-135">We value great branding, and we know it is important to provide users with your unique experience.</span></span> <span data-ttu-id="d3b42-136">とはいえ、優れたアドインを設計するために最適なのは、さり気ない方法でブランド要素を取り入れて直感的なエクスぺリンスを作り上げる方法です。対照的に、執拗に押しつけがましい方法でブランド要素を表示すると、邪魔されずにシステム内を移動しようとするユーザーの気を散らすことになるだけです。</span><span class="sxs-lookup"><span data-stu-id="d3b42-136">But we feel the best way to ensure your add-in's success is to build an intuitive experience that subtly incorporates elements of your brand versus displaying persistent or obtrusive brand elements that only distract a user from moving through your system in an unencumbered manner.</span></span> 
    
   <span data-ttu-id="d3b42-137">ブランドを取り込むための有意義で優れた方法は、ブランドの色、アイコン、音声を使用するというものです (ただし、推奨されるプラットフォーム パターンやユーザー補助機能の要件と競合しないことが前提です)。</span><span class="sxs-lookup"><span data-stu-id="d3b42-137">A good way to incorporate your brand in a meaningful way is through the use of your brand colors, icons, and voice&mdash;assuming these don't conflict with the preferred platform patterns or accessibility requirements.</span></span> <span data-ttu-id="d3b42-138">ブランドに注意を向けるよりも、コンテンツやタスクを完了することに重点を置いてください。</span><span class="sxs-lookup"><span data-stu-id="d3b42-138">Strive to keep the focus on content and task completion, not brand attention.</span></span> 
    
   > [!NOTE]
   >  <span data-ttu-id="d3b42-139">iOS または Android 上のアドイン内に広告が表示されないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d3b42-139">Ads should not be shown within add-ins on iOS or Android.</span></span>

## <a name="design-patterns"></a><span data-ttu-id="d3b42-140">設計パターン</span><span class="sxs-lookup"><span data-stu-id="d3b42-140">Design patterns</span></span>

> [!NOTE]
> <span data-ttu-id="d3b42-141">前述の原則はすべてのエンドポイントおよびプラットフォームに適用されますが、以下のパターンと例は iOS プラットフォームのモバイル アドインに固有です。</span><span class="sxs-lookup"><span data-stu-id="d3b42-141">While the above principles apply to all endpoints/platforms, the following patterns and examples are specific to mobile add-ins on the iOS platform.</span></span>

<span data-ttu-id="d3b42-p111">適切に設計されたアドインを作成できるよう、Outlook Mobile 環境内で動作する iOS モバイル パターンを含む[テンプレート](../design/ux-design-pattern-templates.md)を準備しています。これらの固有のパターンを使用すると、iOS プラットフォームと Outlook Mobile の両方においてネイティブに感じるアドインを作成できます。以下に、これらのパターンの詳細について説明します。これですべてというわけではなく、ライブラリの開始に過ぎません。今後も、これらのアドインに含めるその他のパラダイム パターンを見つけて、作成する予定です。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p111">To help you create a well-designed add-in, we have [templates](../design/ux-design-pattern-templates.md) that contain iOS mobile patterns that work within the Outlook Mobile environment. Leveraging these specific patterns will help ensure your add-in feels native to both the iOS platform and Outlook Mobile. These patterns are also detailed below. While not exhaustive, this is the start of a library that we will continue to build upon as we uncover additional paradigms partners wish to include in their add-ins.</span></span>  

### <a name="overview"></a><span data-ttu-id="d3b42-146">概要</span><span class="sxs-lookup"><span data-stu-id="d3b42-146">Overview</span></span>

<span data-ttu-id="d3b42-147">標準的なアドインは、次のコンポーネントで構成されます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-147">A typical add-in is made up of the following components.</span></span>

![iOS での作業ウィンドウの基本 UX パターンのダイアグラム](../images/outlook-mobile-design-overview.png)

![Android での作業ウィンドウの基本 UX パターンのダイアグラム](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a><span data-ttu-id="d3b42-150">読み込み中</span><span class="sxs-lookup"><span data-stu-id="d3b42-150">Loading</span></span>

<span data-ttu-id="d3b42-p112">ユーザーがアドインをタップしたら、できるだけ早く UX が表示される必要があります。遅延がある場合、進行状況バーやアクティビティ インジケータを使用します。所要時間が分かる場合には進行状況バーを使用し、分からない場合にはアクティビティ インジケータを使用します。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p112">When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.</span></span>

<span data-ttu-id="d3b42-154">**iOS でのページの読み込みの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-154">**An example of loading pages on iOS**</span></span>

![iOS の進行状況バーとアクティビティ インジケータの例](../images/outlook-mobile-design-loading.png)

<span data-ttu-id="d3b42-156">**Android でのページの読み込みの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-156">**An example of loading pages on Android**</span></span>

![Android の進行状況バーとアクティビティ インジケータの例](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a><span data-ttu-id="d3b42-158">サインイン/サインアップ</span><span class="sxs-lookup"><span data-stu-id="d3b42-158">Sign in/Sign up</span></span>

<span data-ttu-id="d3b42-159">サインイン (およびサインアップ) フローを簡単で使用しやすいものにします。</span><span class="sxs-lookup"><span data-stu-id="d3b42-159">Make your sign in (and sign up) flow straightforward and simple to use.</span></span>

<span data-ttu-id="d3b42-160">**iOS のサインイン ページとサインアップ ページの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-160">**An example sign in and sign up page on iOS**</span></span>

![iOS のサインイン ページとサインアップ ページの例](../images/outlook-mobile-design-signin.png)

<span data-ttu-id="d3b42-162">**Android のサインイン ページの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-162">**An example sign in page on Android**</span></span>

![Android のサインイン ページの例](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a><span data-ttu-id="d3b42-164">ブランド バー</span><span class="sxs-lookup"><span data-stu-id="d3b42-164">Brand bar</span></span>

<span data-ttu-id="d3b42-p113">アドインの最初の画面には、ブランドの構成要素を含める必要があります。ブランド バーは、企業の存在を認識するように設計されているため、ユーザーに背景を説明する際にも役立ちます。ナビゲーション バーには会社/ブランドの名前が含まれているため、後続のページのブランド バーには繰り返して含める必要がありません。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p113">The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company/brand, it's unnecessary to repeat the brand bar on subsequent pages.</span></span>

<span data-ttu-id="d3b42-168">**iOS でのブランド化の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-168">**An example of branding on iOS**</span></span>

![iOS でのブランド バーの例](../images/outlook-mobile-design-branding.png)

<span data-ttu-id="d3b42-170">**Android でのブランド化の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-170">**An example of branding on Android**</span></span>

![Android でのブランド バーの例](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a><span data-ttu-id="d3b42-172">余白</span><span class="sxs-lookup"><span data-stu-id="d3b42-172">Margins</span></span>

<span data-ttu-id="d3b42-173">Outlook iOS に合わせるため、モバイルの余白を両側でそれぞれ 15 ピクセル (画面の 8%) に設定します。Outlook Android の場合は、モバイルの余白を両側でそれぞれ 16 ピクセルに設定します。</span><span class="sxs-lookup"><span data-stu-id="d3b42-173">Mobile margins should be set to 15px (8% of screen) for each side, to align with Outlook iOS and 16px for each side to align with Outlook Android.</span></span>

![iOS の余白の例](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a><span data-ttu-id="d3b42-175">文字体裁</span><span class="sxs-lookup"><span data-stu-id="d3b42-175">Typography</span></span>

<span data-ttu-id="d3b42-176">文字体裁の使用法は Outlook iOS に合わせ、スキャンできるようにシンプルにします。</span><span class="sxs-lookup"><span data-stu-id="d3b42-176">Typography usage is aligned to Outlook iOS and is kept simple for scannability.</span></span>

<span data-ttu-id="d3b42-177">**iOS の文字体裁**</span><span class="sxs-lookup"><span data-stu-id="d3b42-177">**Typography on iOS**</span></span>

![iOS の文字体裁のサンプル](../images/outlook-mobile-design-typography.png)

<span data-ttu-id="d3b42-179">**Android の文字体裁**</span><span class="sxs-lookup"><span data-stu-id="d3b42-179">**Typography on Android**</span></span>

![Android の文字体裁のサンプル](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a><span data-ttu-id="d3b42-181">カラー パレット</span><span class="sxs-lookup"><span data-stu-id="d3b42-181">Color palette</span></span>

<span data-ttu-id="d3b42-p114">Outlook iOS における色の使用法は明確ではありません。合わせるには、ブランド バーでのみ固有の色を使用して、その他の色の使用に関しては操作とエラーの状態に応じてローカライズするようお願いいたします。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p114">Color usage is subtle in Outlook iOS.  To align, we ask that usage of color is localized to actions and error states, with only the brand bar using a unique color.</span></span>

![iOS のカラー パレット](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a><span data-ttu-id="d3b42-185">セル</span><span class="sxs-lookup"><span data-stu-id="d3b42-185">Cells</span></span>

<span data-ttu-id="d3b42-186">ナビゲーション バーを使用してページにラベルを付けることはできないため、セクション タイトルを使用してページにラベルを付けます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-186">Since the navigation bar cannot be used to label a page, use section titles to label pages.</span></span>

<span data-ttu-id="d3b42-187">**iOS のセルの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-187">**Examples of cells on iOS**</span></span>

![iOS のセルの種類](../images/outlook-mobile-design-cell-types.png)
* * *
![iOS で「使用可能」なセル](../images/outlook-mobile-design-cell-dos.png)
* * *
![iOS で「使用不可」のセル](../images/outlook-mobile-design-cell-donts.png)
* * *
![iOS のセルと入力](../images/outlook-mobile-design-cell-input.png)

<span data-ttu-id="d3b42-192">**Android のセルの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-192">**Examples of cells on Android**</span></span>

![Android のセルの種類](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Android で「使用可能」なセル](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Android で「使用不可」なセル](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Android のセルと入力パート 1](../images/outlook-mobile-design-cell-input-1-android.png)

![Android のセルと入力パート 2](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a><span data-ttu-id="d3b42-198">アクション</span><span class="sxs-lookup"><span data-stu-id="d3b42-198">Actions</span></span>

<span data-ttu-id="d3b42-199">アプリでさまざまなアクションを処理する場合も、アドインで実行する最も重要なアクションについて考慮し、それらに注意を集中します。</span><span class="sxs-lookup"><span data-stu-id="d3b42-199">Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.</span></span>

<span data-ttu-id="d3b42-200">**iOS でのアクションの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-200">**Examples of actions on iOS**</span></span>

![iOS でのアクションとセル](../images/outlook-mobile-design-action-cells.png)
* * *
![iOS で「使用可能」なアクション](../images/outlook-mobile-design-action-dos.png)

<span data-ttu-id="d3b42-203">**Android でのアクションの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-203">**Examples of actions on Android**</span></span>

![Android でのアクションとセル](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Android で「使用可能」なアクション](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a><span data-ttu-id="d3b42-206">ボタン</span><span class="sxs-lookup"><span data-stu-id="d3b42-206">Buttons</span></span>

<span data-ttu-id="d3b42-207">ボタンは、その後に他の UX 要素がある場合に使用します (一方、アクションは画面上における最後の要素です)。</span><span class="sxs-lookup"><span data-stu-id="d3b42-207">Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).</span></span>

<span data-ttu-id="d3b42-208">**iOS のボタンの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-208">**Examples of buttons on iOS**</span></span>

![iOS のボタンの例](../images/outlook-mobile-design-buttons.png)

<span data-ttu-id="d3b42-210">**Android のボタンの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-210">**Examples of buttons on Android**</span></span>

![Android のボタンの例](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a><span data-ttu-id="d3b42-212">タブ</span><span class="sxs-lookup"><span data-stu-id="d3b42-212">Tabs</span></span>

<span data-ttu-id="d3b42-213">タブを使用すると、コンテンツを整理できます。</span><span class="sxs-lookup"><span data-stu-id="d3b42-213">Tabs can aid in content organization.</span></span>

<span data-ttu-id="d3b42-214">**iOS のタブの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-214">**Examples of tabs on iOS**</span></span>

![iOS のタブの例](../images/outlook-mobile-design-tabs.png)

<span data-ttu-id="d3b42-216">**Android のタブの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-216">**Examples of tabs on Android**</span></span>

![Android のタブの例](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a><span data-ttu-id="d3b42-218">アイコン</span><span class="sxs-lookup"><span data-stu-id="d3b42-218">Icons</span></span>

<span data-ttu-id="d3b42-p115">アイコンに関しては、可能な場合には Outlook iOS の現行デザインに従ってください。当社の標準的なサイズと色を使用します。</span><span class="sxs-lookup"><span data-stu-id="d3b42-p115">Icons should follow the current Outlook iOS design when possible. Use our standard size and color.</span></span>

<span data-ttu-id="d3b42-221">**iOS のアイコンの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-221">**Examples of icons on iOS**</span></span>

![iOS のアイコンの例](../images/outlook-mobile-design-icons.png)

<span data-ttu-id="d3b42-223">**Android のアイコンの例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-223">**Examples of icons on Android**</span></span>

![Android のアイコンの例](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a><span data-ttu-id="d3b42-225">エンド ツー エンドの例</span><span class="sxs-lookup"><span data-stu-id="d3b42-225">End-to-end examples</span></span>

<span data-ttu-id="d3b42-226">v1 Outlook Mobile アドインを発表して以降、アドインを作成したパートナーと緊密に連携してきました。Outlook Mobile におけるアドインの効果性を示すため、当社のデザイナーは、当社のガイドラインとパターンを使用して、各アドインのエンド ツー エンド フローをまとめました。</span><span class="sxs-lookup"><span data-stu-id="d3b42-226">For our v1 Outlook Mobile Add-ins launch, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook Mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d3b42-227">次に示す例は、アドインの相互作用と視覚上の設計を行うための理想的な方法を強調するためのもので、出荷バージョンのアドインの機能セットと正確に一致しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="d3b42-227">These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.</span></span> 

### <a name="giphy"></a><span data-ttu-id="d3b42-228">GIPHY</span><span class="sxs-lookup"><span data-stu-id="d3b42-228">GIPHY</span></span>

<span data-ttu-id="d3b42-229">**iOS での GIPHY の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-229">**An example of GIPHY on iOS**</span></span>

![iOS の GIPHY アドインのエンド ツー エンド設計](../images/outlook-mobile-design-giphy.png)

<span data-ttu-id="d3b42-231">**Android での GIPHY の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-231">**An example of GIPHY on Android**</span></span>

![Android の GIPHY アドインのエンド ツー エンド設計](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a><span data-ttu-id="d3b42-233">Nimble</span><span class="sxs-lookup"><span data-stu-id="d3b42-233">Nimble</span></span>

<span data-ttu-id="d3b42-234">**iOS での Nimble の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-234">**An example of Nimble on iOS**</span></span>

![iOS の Nimble アドインのエンド ツー エンド設計](../images/outlook-mobile-design-nimble.png)

<span data-ttu-id="d3b42-236">**Android での Nimble の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-236">**An example of Nimble on Android**</span></span>

![Android の Nimble アドインのエンド ツー エンド設計](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a><span data-ttu-id="d3b42-238">Trello</span><span class="sxs-lookup"><span data-stu-id="d3b42-238">Trello</span></span>

<span data-ttu-id="d3b42-239">**iOS での Trello の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-239">**An example of Trello on iOS**</span></span>

![iOS の Trello アドインのエンド ツー エンド設計パート 1](../images/outlook-mobile-design-trello-1.png)
* * *
![iOS の Trello アドインのエンド ツー エンド設計パート 2](../images/outlook-mobile-design-trello-2.png)
* * *
![iOS の Trello アドインのエンド ツー エンド設計パート 3](../images/outlook-mobile-design-trello-3.png)

<span data-ttu-id="d3b42-243">**Android での Trello の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-243">**An example of Trello on Android**</span></span>

![Android の Trello アドインのエンド ツー エンド設計パート 1](../images/outlook-mobile-design-trello-1-android.png)
* * *
![Android の Trello アドインのエンド ツー エンド設計パート 2](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a><span data-ttu-id="d3b42-246">Dynamics CRM</span><span class="sxs-lookup"><span data-stu-id="d3b42-246">Dynamics CRM</span></span>

<span data-ttu-id="d3b42-247">**iOS での Dynamics CRM の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-247">**An example of Dynamics CRM on iOS**</span></span>

![iOS の Dynamics CRM アドインのエンド ツー エンド設計](../images/outlook-mobile-design-crm.png)

<span data-ttu-id="d3b42-249">**Android での Dynamics CRM の例**</span><span class="sxs-lookup"><span data-stu-id="d3b42-249">**An example of Dynamics CRM on Android**</span></span>

![Android の Dynamics CRM アドインのエンド ツー エンド設計](../images/outlook-mobile-design-crm-android.png)
