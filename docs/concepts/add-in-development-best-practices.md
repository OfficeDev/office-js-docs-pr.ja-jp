---
title: Office アドイン開発のベスト プラクティス
description: Office アドインを作成するための開発時にベストプラクティスを適用します。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 17393d921129efcfb74eed3dd168633c2f58291b
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132180"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="4b02a-103">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="4b02a-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="4b02a-p101">効果的なアドインは、目で見て分かる方法で Office アプリケーションを拡張する、ユニークで頼もしい機能を提供します。優れたアドインを作成するには、魅力的な初回エクスペリエンスをユーザーに提供して、最高の UI エクスペリエンスを設計し、アドインのパフォーマンスを最適化します。この記事で説明するベスト プラクティスを適用して、ユーザーが迅速かつ効率的に仕事を遂行するための助けになるアドインを作成してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="4b02a-107">価値を明確にする</span><span class="sxs-lookup"><span data-stu-id="4b02a-107">Provide clear value</span></span>

- <span data-ttu-id="4b02a-p102">ユーザーがタスクをすばやく効率的に完了するのに役立つアドインを作成します。Office アプリケーションに当てはまるシナリオに絞ります。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
  - <span data-ttu-id="4b02a-111">コア オーサリング タスクをよりスピーディかつ簡単にし、中断を減らします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
  - <span data-ttu-id="4b02a-112">Office 内で新しいシナリオを有効にします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-112">Enable new scenarios within Office.</span></span>
  - <span data-ttu-id="4b02a-113">Office アプリケーション内に補完的サービスを埋め込む。</span><span class="sxs-lookup"><span data-stu-id="4b02a-113">Embed complementary services within Office applications.</span></span>
  - <span data-ttu-id="4b02a-114">Office エクスペリエンスを向上させて生産性を高めます。</span><span class="sxs-lookup"><span data-stu-id="4b02a-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="4b02a-115">[魅力的な初回実行時エクスペリエンス](#create-an-engaging-first-run-experience)を作成して、ユーザーがアドインの価値をすぐに感じられるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="4b02a-p103">[効果的な AppSource リスト](/office/dev/store/create-effective-office-store-listings)を作成します。タイトルと説明から、アドインのメリットが明確にわかるようにします。アドインの内容を伝えるのに、ブランドだけに頼ることはしないでください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>

## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="4b02a-119">魅力的な初回実行時エクスペリエンスを作成する</span><span class="sxs-lookup"><span data-stu-id="4b02a-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="4b02a-p104">非常に使いやすく、直観で理解しやすいファースト エクスペリエンスによって、新しいユーザーを引き込みます。ユーザーは、アドインをストアからダウンロードした後も、使用するか中止するかを判断しています。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="4b02a-p105">ユーザーがアドインを使用するのに必要な手順を明確にします。ビデオ、マット、ページング パネル、その他のリソースを使用して、ユーザーを誘導します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="4b02a-124">単純にユーザーにサインを求めるのではなく、起動時にアドインの価値を強調します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="4b02a-125">使い方や UI を個人用に設定する方法を説明する UI を提供します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-125">Provide teaching UI to guide users and make your UI personal.</span></span>

  !["Do" と "いいえ" の比較を示すスクリーンショット。](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="4b02a-129">コンテンツ アドインがユーザーのドキュメント内のデータにバインドされている場合は、サンプル データまたはテンプレートを含めて、使用するデータ形式をユーザーに表示します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-129">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

  !["Do" と "いいえ" の比較を示すスクリーンショット。](../images/add-in-title.png)

- <span data-ttu-id="4b02a-p108">[無料の試用版](/office/dev/store/decide-on-a-pricing-model)を提供します。アドインでサブスクリプションを要求する場合は、一部の機能をサブスクリプションなしでも利用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p108">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="4b02a-p109">サインアップをシンプルにします。情報 (電子メール、表示名) を事前に入力し、電子メールの確認はスキップします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p109">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="4b02a-p110">ポップアップは使用しないようにします。使用する必要がある場合は、ポップアップを有効にするようユーザーに指示します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p110">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="4b02a-139">最初の実行エクスペリエンスを開発する際に適用できるパターンについては、「[Office アドインの UX 設計パターン](../design/first-run-experience-patterns.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-139">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="4b02a-140">アドイン コマンドを使用する</span><span class="sxs-lookup"><span data-stu-id="4b02a-140">Use add-in commands</span></span>

- <span data-ttu-id="4b02a-p111">アドイン コマンドを使用することで、アドインに関連する UI エントリ ポイントを提供します。設計のベスト プラクティスを含む詳細については、「[アドイン コマンド](../design/add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p111">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="4b02a-143">UX 設計原則を適用する</span><span class="sxs-lookup"><span data-stu-id="4b02a-143">Apply UX design principles</span></span>

- <span data-ttu-id="4b02a-p112">アドインの外観と機能が、Office のエクスペリエンスと合っていることを確認します。[Office UI Fabric](https://developer.microsoft.com/fabric) を使用します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p112">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="4b02a-p113">クロムよりもコンテンツを優先します。ユーザー エクスペリエンスの価値を高めない余分な UI 要素を追加しないようにします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p113">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="4b02a-p114">ユーザーをよく管理します。ユーザーが重要な決定事項を理解し、アドインが実行するアクションを簡単に取り消すことができるようにします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p114">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="4b02a-p115">ユーザーの信頼を得て、ユーザーを引き込むために
ブランドを利用します。ユーザーを圧倒するためや、宣伝のためにブランドを使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p115">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="4b02a-p116">スクロールしないようにします。1366 x 768 の解像度用に最適化します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p116">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="4b02a-154">使用許可を得ていないイメージを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-154">Do not include unlicensed images.</span></span>

- <span data-ttu-id="4b02a-155">アドインでは[明確でシンプルな表現](../design/voice-guidelines.md)を使用してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-155">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="4b02a-156">アクセシビリティを考慮してください。すべてのユーザーにとって操作しやすいアドインにして、画面リーダーなどの支援テクノロジが利用できるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-156">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="4b02a-p117">すべてのプラットフォームと入力方法 (マウスやキーボード、および [タッチ](#optimize-for-touch)など) に対応するように設計してください。UI が様々なフォーム ファクターに対応するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p117">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="4b02a-159">タッチ用に最適化する</span><span class="sxs-lookup"><span data-stu-id="4b02a-159">Optimize for touch</span></span>

- <span data-ttu-id="4b02a-160">[Context.touchenabled](/javascript/api/office/office.context#touchenabled)プロパティを使用して、アドインが実行されている Office アプリケーションがタッチに対応しているかどうかを検出します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-160">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="4b02a-161">このプロパティは、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4b02a-161">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="4b02a-p118">すべてのコントロールがタッチ操作に適したサイズになっていることを確認します。たとえば、ボタンに適切なタッチ ターゲットを設定し、入力ボックスはユーザーが入力するのに十分な大きさにします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p118">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="4b02a-164">ホバーや右クリックなどの非タッチの入力方法に依存しないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-164">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="4b02a-p119">縦向きと横向きの両方のモードでアドインが機能することを確認します。タッチ デバイスで、アドインの一部がソフトキーボードの後ろに隠れることがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p119">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="4b02a-167">[サイドロード](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)を使用して、アドインを実際のデバイスでテストしてください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-167">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="4b02a-168">[Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) をデザイン要素に使用している場合は、これらの要素の多くが適切に設定されます。</span><span class="sxs-lookup"><span data-stu-id="4b02a-168">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="4b02a-169">アドインのパフォーマンスを最適化して監視する</span><span class="sxs-lookup"><span data-stu-id="4b02a-169">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="4b02a-p120">UI が素早く応答する感覚を与えるようにします。アドインが 500 ミリ秒以内で読み込まれるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p120">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="4b02a-172">すべてのユーザー操作が 1 秒以内で応答することを確認します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-172">Ensure that all user interactions respond in under one second.</span></span>

- <span data-ttu-id="4b02a-173">長時間実行する操作には、読み込みインジケーターを提供します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-173">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="4b02a-p121">画像、リソース、および一般的なライブラリを CDN を使用してホストします。可能な限り多くのものを 1 つの場所から読み込みます。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p121">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="4b02a-p122">Web ページを最適化するには、標準的な Web の慣習に従います。運用環境では、ライブラリの縮小バージョンのみを使用します。必要なリソースのみを読み込み、リソースが読み込まれる方法を最適化します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p122">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="4b02a-p123">操作の実行に時間がかかる場合は、ユーザーにフィードバックを提供します。次の表のしきい値に注意してください。追加情報については、「[Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p123">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="4b02a-182">操作クラス</span><span class="sxs-lookup"><span data-stu-id="4b02a-182">Interaction class</span></span>|<span data-ttu-id="4b02a-183">Target</span><span class="sxs-lookup"><span data-stu-id="4b02a-183">Target</span></span>|<span data-ttu-id="4b02a-184">上限</span><span class="sxs-lookup"><span data-stu-id="4b02a-184">Upper bound</span></span>|<span data-ttu-id="4b02a-185">人間の感覚</span><span class="sxs-lookup"><span data-stu-id="4b02a-185">Human perception</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="4b02a-186">即時</span><span class="sxs-lookup"><span data-stu-id="4b02a-186">Instant</span></span>|<span data-ttu-id="4b02a-187">50 ミリ秒以下</span><span class="sxs-lookup"><span data-stu-id="4b02a-187"><=50 ms</span></span>|<span data-ttu-id="4b02a-188">100 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-188">100 ms</span></span>|<span data-ttu-id="4b02a-189">顕著な遅延はない。</span><span class="sxs-lookup"><span data-stu-id="4b02a-189">No noticeable delay.</span></span>|
  |<span data-ttu-id="4b02a-190">速く</span><span class="sxs-lookup"><span data-stu-id="4b02a-190">Fast</span></span>|<span data-ttu-id="4b02a-191">50 から 100 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-191">50-100 ms</span></span>|<span data-ttu-id="4b02a-192">200 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-192">200 ms</span></span>|<span data-ttu-id="4b02a-p124">最低限知覚される遅延。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p124">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="4b02a-195">普通</span><span class="sxs-lookup"><span data-stu-id="4b02a-195">Typical</span></span>|<span data-ttu-id="4b02a-196">100-300 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-196">100-300 ms</span></span>|<span data-ttu-id="4b02a-197">500 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-197">500 ms</span></span>|<span data-ttu-id="4b02a-p125">速い。しかし、高速とまではいかない。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p125">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="4b02a-200">速い</span><span class="sxs-lookup"><span data-stu-id="4b02a-200">Responsive</span></span>|<span data-ttu-id="4b02a-201">300-500 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-201">300-500 ms</span></span>|<span data-ttu-id="4b02a-202">1 秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-202">1 second</span></span>|<span data-ttu-id="4b02a-p126">高速ではないが、速いという実感はある。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p126">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="4b02a-205">連続</span><span class="sxs-lookup"><span data-stu-id="4b02a-205">Continuous</span></span>|<span data-ttu-id="4b02a-206">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-206">>500 ms</span></span>|<span data-ttu-id="4b02a-207">5 秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-207">5 seconds</span></span>|<span data-ttu-id="4b02a-p127">中程度の待ち時間。速いという実感はない。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p127">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="4b02a-210">キャプティブ</span><span class="sxs-lookup"><span data-stu-id="4b02a-210">Captive</span></span>|<span data-ttu-id="4b02a-211">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-211">>500 ms</span></span>|<span data-ttu-id="4b02a-212">10 秒</span><span class="sxs-lookup"><span data-stu-id="4b02a-212">10 seconds</span></span>|<span data-ttu-id="4b02a-p128">長い。しかし、何か他のことを行えるほどの長さではない。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p128">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="4b02a-215">拡張</span><span class="sxs-lookup"><span data-stu-id="4b02a-215">Extended</span></span>|<span data-ttu-id="4b02a-216">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-216">>500 ms</span></span>|<span data-ttu-id="4b02a-217">10 秒より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-217">>10 seconds</span></span>|<span data-ttu-id="4b02a-p129">待機中に他のことを行うのに十分な長さ。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p129">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="4b02a-220">長時間実行</span><span class="sxs-lookup"><span data-stu-id="4b02a-220">Long running</span></span>|<span data-ttu-id="4b02a-221">5 秒より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-221">>5 seconds</span></span>|<span data-ttu-id="4b02a-222">1 分より長い</span><span class="sxs-lookup"><span data-stu-id="4b02a-222">>1 minute</span></span>|<span data-ttu-id="4b02a-223">ユーザーは確実に別のことを行えます。</span><span class="sxs-lookup"><span data-stu-id="4b02a-223">Users will certainly do something else.</span></span>|

- <span data-ttu-id="4b02a-224">サービスの正常性を監視し、テレメトリを使用して、ユーザーが正常に完了したか監視します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-224">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="4b02a-225">アドインと Office ドキュメント間のデータ交換を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-225">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="4b02a-226">詳細については、「 [ループでのコンテキストの同期方法の使用を避ける](correlated-objects-pattern.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-226">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="4b02a-227">アドインを売り込む</span><span class="sxs-lookup"><span data-stu-id="4b02a-227">Market your add-in</span></span>

- <span data-ttu-id="4b02a-p131">アドインを [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) に公開して、Web サイトで[それを宣伝](/office/dev/store/promote-your-office-store-solution)します。[効果的な AppSource リストを作成します](/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p131">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="4b02a-p132">アドイン タイトルを簡潔でわかりやすいものにします。128 文字以下にします。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p132">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="4b02a-p133">アドインの短くて魅力的な説明を作成します。「このアドインでどんな問題が解決しますか？」という質問への答えを提供します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p133">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="4b02a-p134">タイトルと説明でアドインの価値提案を行います。ブランドに依存しないでください。</span><span class="sxs-lookup"><span data-stu-id="4b02a-p134">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="4b02a-236">ユーザーがアドインを見つけて使うことができる Web サイトを作成します。</span><span class="sxs-lookup"><span data-stu-id="4b02a-236">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="4b02a-237">Internet Explorer をサポートする JavaScript の使用</span><span class="sxs-lookup"><span data-stu-id="4b02a-237">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="4b02a-238">関連項目</span><span class="sxs-lookup"><span data-stu-id="4b02a-238">See also</span></span>

- [<span data-ttu-id="4b02a-239">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4b02a-239">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="4b02a-240">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="4b02a-240">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
