---
title: Office アドイン開発のベスト プラクティス
description: Office アドインを作成するための開発時にベストプラクティスを適用します。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 72bcf8c7bd994c499c1e85154a383c3a91464a79
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293066"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="cf42f-103">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cf42f-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="cf42f-p101">効果的なアドインは、目で見て分かる方法で Office アプリケーションを拡張する、ユニークで頼もしい機能を提供します。優れたアドインを作成するには、魅力的な初回エクスペリエンスをユーザーに提供して、最高の UI エクスペリエンスを設計し、アドインのパフォーマンスを最適化します。この記事で説明するベスト プラクティスを適用して、ユーザーが迅速かつ効率的に仕事を遂行するための助けになるアドインを作成してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="cf42f-107">価値を明確にする</span><span class="sxs-lookup"><span data-stu-id="cf42f-107">Provide clear value</span></span>

- <span data-ttu-id="cf42f-p102">ユーザーがタスクをすばやく効率的に完了するのに役立つアドインを作成します。Office アプリケーションに当てはまるシナリオに絞ります。次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="cf42f-111">コア オーサリング タスクをよりスピーディかつ簡単にし、中断を減らします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="cf42f-112">Office 内で新しいシナリオを有効にします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-112">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="cf42f-113">Office アプリケーション内に補完的サービスを埋め込む。</span><span class="sxs-lookup"><span data-stu-id="cf42f-113">Embed complementary services within Office applications.</span></span>
 - <span data-ttu-id="cf42f-114">Office エクスペリエンスを向上させて生産性を高めます。</span><span class="sxs-lookup"><span data-stu-id="cf42f-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="cf42f-115">[魅力的な初回実行時エクスペリエンス](#create-an-engaging-first-run-experience)を作成して、ユーザーがアドインの価値をすぐに感じられるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="cf42f-p103">[効果的な AppSource リスト](/office/dev/store/create-effective-office-store-listings)を作成します。タイトルと説明から、アドインのメリットが明確にわかるようにします。アドインの内容を伝えるのに、ブランドだけに頼ることはしないでください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="cf42f-119">魅力的な初回実行時エクスペリエンスを作成する</span><span class="sxs-lookup"><span data-stu-id="cf42f-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="cf42f-p104">非常に使いやすく、直観で理解しやすいファースト エクスペリエンスによって、新しいユーザーを引き込みます。ユーザーは、アドインをストアからダウンロードした後も、使用するか中止するかを判断しています。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="cf42f-p105">ユーザーがアドインを使用するのに必要な手順を明確にします。ビデオ、マット、ページング パネル、その他のリソースを使用して、ユーザーを誘導します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="cf42f-124">単純にユーザーにサインを求めるのではなく、起動時にアドインの価値を強調します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="cf42f-125">使い方や UI を個人用に設定する方法を説明する UI を提供します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-125">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![作業の開始手順が含まれないアドインの隣に、作業の開始手順を含むアドインの作業ウィンドウを示すスクリーンショット](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="cf42f-127">コンテンツ アドインがユーザーのドキュメント内のデータにバインドされている場合は、サンプル データまたはテンプレートを含めて、使用するデータ形式をユーザーに表示します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-127">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![データを含まないコンテンツ アドインの隣に、データを含むコンテンツ アドインを示すスクリーンショット](../images/add-in-title.png)

- <span data-ttu-id="cf42f-p106">[無料の試用版](/office/dev/store/decide-on-a-pricing-model)を提供します。アドインでサブスクリプションを要求する場合は、一部の機能をサブスクリプションなしでも利用できるようにします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p106">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="cf42f-p107">サインアップをシンプルにします。情報 (電子メール、表示名) を事前に入力し、電子メールの確認はスキップします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p107">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="cf42f-p108">ポップアップは使用しないようにします。使用する必要がある場合は、ポップアップを有効にするようユーザーに指示します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p108">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="cf42f-135">最初の実行エクスペリエンスを開発する際に適用できるパターンについては、「[Office アドインの UX 設計パターン](../design/first-run-experience-patterns.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-135">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="cf42f-136">アドイン コマンドを使用する</span><span class="sxs-lookup"><span data-stu-id="cf42f-136">Use add-in commands</span></span>

- <span data-ttu-id="cf42f-p109">アドイン コマンドを使用することで、アドインに関連する UI エントリ ポイントを提供します。設計のベスト プラクティスを含む詳細については、「[アドイン コマンド](../design/add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p109">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="cf42f-139">UX 設計原則を適用する</span><span class="sxs-lookup"><span data-stu-id="cf42f-139">Apply UX design principles</span></span>

- <span data-ttu-id="cf42f-p110">アドインの外観と機能が、Office のエクスペリエンスと合っていることを確認します。[Office UI Fabric](https://developer.microsoft.com/fabric) を使用します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p110">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="cf42f-p111">クロムよりもコンテンツを優先します。ユーザー エクスペリエンスの価値を高めない余分な UI 要素を追加しないようにします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p111">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="cf42f-p112">ユーザーをよく管理します。ユーザーが重要な決定事項を理解し、アドインが実行するアクションを簡単に取り消すことができるようにします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p112">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="cf42f-p113">ユーザーの信頼を得て、ユーザーを引き込むために
ブランドを利用します。ユーザーを圧倒するためや、宣伝のためにブランドを使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p113">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="cf42f-p114">スクロールしないようにします。1366 x 768 の解像度用に最適化します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p114">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="cf42f-150">使用許可を得ていないイメージを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-150">Do not include unlicensed images.</span></span>

- <span data-ttu-id="cf42f-151">アドインでは[明確でシンプルな表現](../design/voice-guidelines.md)を使用してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-151">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="cf42f-152">アクセシビリティを考慮してください。すべてのユーザーにとって操作しやすいアドインにして、画面リーダーなどの支援テクノロジが利用できるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-152">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="cf42f-p115">すべてのプラットフォームと入力方法 (マウスやキーボード、および [タッチ](#optimize-for-touch)など) に対応するように設計してください。UI が様々なフォーム ファクターに対応するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p115">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="cf42f-155">タッチ用に最適化する</span><span class="sxs-lookup"><span data-stu-id="cf42f-155">Optimize for touch</span></span>

- <span data-ttu-id="cf42f-156">[Context.touchenabled](/javascript/api/office/office.context#touchenabled)プロパティを使用して、アドインが実行されている Office アプリケーションがタッチに対応しているかどうかを検出します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-156">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="cf42f-157">このプロパティは、Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cf42f-157">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="cf42f-p116">すべてのコントロールがタッチ操作に適したサイズになっていることを確認します。たとえば、ボタンに適切なタッチ ターゲットを設定し、入力ボックスはユーザーが入力するのに十分な大きさにします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p116">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="cf42f-160">ホバーや右クリックなどの非タッチの入力方法に依存しないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-160">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="cf42f-p117">縦向きと横向きの両方のモードでアドインが機能することを確認します。タッチ デバイスで、アドインの一部がソフトキーボードの後ろに隠れることがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p117">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="cf42f-163">[サイドロード](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)を使用して、アドインを実際のデバイスでテストしてください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-163">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="cf42f-164">[Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) をデザイン要素に使用している場合は、これらの要素の多くが適切に設定されます。</span><span class="sxs-lookup"><span data-stu-id="cf42f-164">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="cf42f-165">アドインのパフォーマンスを最適化して監視する</span><span class="sxs-lookup"><span data-stu-id="cf42f-165">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="cf42f-p118">UI が素早く応答する感覚を与えるようにします。アドインが 500 ミリ秒以内で読み込まれるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p118">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="cf42f-168">すべてのユーザー操作が 1 秒以内で応答することを確認します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-168">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="cf42f-169">長時間実行する操作には、読み込みインジケーターを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-169">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="cf42f-p119">画像、リソース、および一般的なライブラリを CDN を使用してホストします。可能な限り多くのものを 1 つの場所から読み込みます。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p119">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="cf42f-p120">Web ページを最適化するには、標準的な Web の慣習に従います。運用環境では、ライブラリの縮小バージョンのみを使用します。必要なリソースのみを読み込み、リソースが読み込まれる方法を最適化します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p120">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="cf42f-p121">操作の実行に時間がかかる場合は、ユーザーにフィードバックを提供します。次の表のしきい値に注意してください。追加情報については、「[Office アドインのリソースの制限とパフォーマンスの最適化](../concepts/resource-limits-and-performance-optimization.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p121">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="cf42f-178">**インタラクション クラス**</span><span class="sxs-lookup"><span data-stu-id="cf42f-178">**Interaction class**</span></span>|<span data-ttu-id="cf42f-179">**ターゲット**</span><span class="sxs-lookup"><span data-stu-id="cf42f-179">**Target**</span></span>|<span data-ttu-id="cf42f-180">**上限値**</span><span class="sxs-lookup"><span data-stu-id="cf42f-180">**Upper bound**</span></span>|<span data-ttu-id="cf42f-181">**人間の知覚**</span><span class="sxs-lookup"><span data-stu-id="cf42f-181">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="cf42f-182">即時</span><span class="sxs-lookup"><span data-stu-id="cf42f-182">Instant</span></span>|<span data-ttu-id="cf42f-183">50 ミリ秒以下</span><span class="sxs-lookup"><span data-stu-id="cf42f-183"><=50 ms</span></span>|<span data-ttu-id="cf42f-184">100 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-184">100 ms</span></span>|<span data-ttu-id="cf42f-185">顕著な遅延はない。</span><span class="sxs-lookup"><span data-stu-id="cf42f-185">No noticeable delay.</span></span>|
  |<span data-ttu-id="cf42f-186">速く</span><span class="sxs-lookup"><span data-stu-id="cf42f-186">Fast</span></span>|<span data-ttu-id="cf42f-187">50 から 100 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-187">50-100 ms</span></span>|<span data-ttu-id="cf42f-188">200 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-188">200 ms</span></span>|<span data-ttu-id="cf42f-p122">最低限知覚される遅延。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p122">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="cf42f-191">普通</span><span class="sxs-lookup"><span data-stu-id="cf42f-191">Typical</span></span>|<span data-ttu-id="cf42f-192">100-300 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-192">100-300 ms</span></span>|<span data-ttu-id="cf42f-193">500 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-193">500 ms</span></span>|<span data-ttu-id="cf42f-p123">速い。しかし、高速とまではいかない。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p123">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="cf42f-196">速い</span><span class="sxs-lookup"><span data-stu-id="cf42f-196">Responsive</span></span>|<span data-ttu-id="cf42f-197">300-500 ミリ秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-197">300-500 ms</span></span>|<span data-ttu-id="cf42f-198">1 秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-198">1 second</span></span>|<span data-ttu-id="cf42f-p124">高速ではないが、速いという実感はある。フィードバックは不要。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p124">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="cf42f-201">連続</span><span class="sxs-lookup"><span data-stu-id="cf42f-201">Continuous</span></span>|<span data-ttu-id="cf42f-202">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-202">>500 ms</span></span>|<span data-ttu-id="cf42f-203">5 秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-203">5 seconds</span></span>|<span data-ttu-id="cf42f-p125">中程度の待ち時間。速いという実感はない。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p125">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="cf42f-206">キャプティブ</span><span class="sxs-lookup"><span data-stu-id="cf42f-206">Captive</span></span>|<span data-ttu-id="cf42f-207">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-207">>500 ms</span></span>|<span data-ttu-id="cf42f-208">10 秒</span><span class="sxs-lookup"><span data-stu-id="cf42f-208">10 seconds</span></span>|<span data-ttu-id="cf42f-p126">長い。しかし、何か他のことを行えるほどの長さではない。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p126">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="cf42f-211">拡張</span><span class="sxs-lookup"><span data-stu-id="cf42f-211">Extended</span></span>|<span data-ttu-id="cf42f-212">500 ミリ秒より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-212">>500 ms</span></span>|<span data-ttu-id="cf42f-213">10 秒より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-213">>10 seconds</span></span>|<span data-ttu-id="cf42f-p127">待機中に他のことを行うのに十分な長さ。フィードバックが必要な可能性あり。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p127">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="cf42f-216">長時間実行</span><span class="sxs-lookup"><span data-stu-id="cf42f-216">Long running</span></span>|<span data-ttu-id="cf42f-217">5 秒より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-217">>5 seconds</span></span>|<span data-ttu-id="cf42f-218">1 分より長い</span><span class="sxs-lookup"><span data-stu-id="cf42f-218">>1 minute</span></span>|<span data-ttu-id="cf42f-219">ユーザーは確実に別のことを行えます。</span><span class="sxs-lookup"><span data-stu-id="cf42f-219">Users will certainly do something else.</span></span>|

- <span data-ttu-id="cf42f-220">サービスの正常性を監視し、テレメトリを使用して、ユーザーが正常に完了したか監視します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-220">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="cf42f-221">アドインと Office ドキュメント間のデータ交換を最小限にします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-221">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="cf42f-222">詳細については、「 [ループでのコンテキストの同期方法の使用を避ける](correlated-objects-pattern.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-222">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="cf42f-223">アドインを売り込む</span><span class="sxs-lookup"><span data-stu-id="cf42f-223">Market your add-in</span></span>

- <span data-ttu-id="cf42f-p129">アドインを [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) に公開して、Web サイトで[それを宣伝](/office/dev/store/promote-your-office-store-solution)します。[効果的な AppSource リストを作成します](/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p129">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="cf42f-p130">アドイン タイトルを簡潔でわかりやすいものにします。128 文字以下にします。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p130">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="cf42f-p131">アドインの短くて魅力的な説明を作成します。「このアドインでどんな問題が解決しますか？」という質問への答えを提供します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p131">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="cf42f-p132">タイトルと説明でアドインの価値提案を行います。ブランドに依存しないでください。</span><span class="sxs-lookup"><span data-stu-id="cf42f-p132">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="cf42f-232">ユーザーがアドインを見つけて使うことができる Web サイトを作成します。</span><span class="sxs-lookup"><span data-stu-id="cf42f-232">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="cf42f-233">Internet Explorer をサポートする JavaScript の使用</span><span class="sxs-lookup"><span data-stu-id="cf42f-233">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="cf42f-234">関連項目</span><span class="sxs-lookup"><span data-stu-id="cf42f-234">See also</span></span>

- [<span data-ttu-id="cf42f-235">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="cf42f-235">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
