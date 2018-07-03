# <a name="first-run-experience-patterns"></a><span data-ttu-id="082f9-101">最初の実行エクスペリエンスパターン</span><span class="sxs-lookup"><span data-stu-id="082f9-101">First-run experience patterns</span></span>

<span data-ttu-id="082f9-102">最初の実行エクスペリエンス（FRE）は、ユーザーに対するアドインの紹介です。</span><span class="sxs-lookup"><span data-stu-id="082f9-102">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="082f9-103">FRE は、ユーザーが初めてアドインを開いた時に表示され、新機能、特徴、および/またはアドインのメリットがわかります。</span><span class="sxs-lookup"><span data-stu-id="082f9-103">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="082f9-104">このエクスペリエンスは、ユーザーのアドインの印象を形作るのを助け、また戻ってくる、および継続的にアドオンを使用するなどの可能性に強く影響します。</span><span class="sxs-lookup"><span data-stu-id="082f9-104">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="082f9-105">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="082f9-105">Best practices</span></span>


<span data-ttu-id="082f9-106">最初の実行エクスペリエンスを作成する際、次のベストプラクティスに従ってください。</span><span class="sxs-lookup"><span data-stu-id="082f9-106">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="082f9-107">するべきこと</span><span class="sxs-lookup"><span data-stu-id="082f9-107">Do</span></span>|<span data-ttu-id="082f9-108">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="082f9-108">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="082f9-109">アドインの主な操作を簡単に、短く紹介します。</span><span class="sxs-lookup"><span data-stu-id="082f9-109">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="082f9-110">開始するのに関係のない情報やコールアウトを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="082f9-110">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="082f9-111">アドインの使用にプラスの影響を与える操作を完了する機会をユーザーに与えます。</span><span class="sxs-lookup"><span data-stu-id="082f9-111">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="082f9-112">ユーザーが一度にすべてを覚えるとは思わないでください。</span><span class="sxs-lookup"><span data-stu-id="082f9-112">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="082f9-113">最も価値を提供する操作に焦点を当てます。</span><span class="sxs-lookup"><span data-stu-id="082f9-113">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="082f9-114">ユーザーが完了したいと思うような、魅力的な体験を作成します。</span><span class="sxs-lookup"><span data-stu-id="082f9-114">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="082f9-115">ユーザーを強制的に最初の実行エクスペリエンスへ導かないでください。</span><span class="sxs-lookup"><span data-stu-id="082f9-115">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="082f9-116">ユーザーには、最初の実行エクスペリエンスを迂回する選択肢を与えます。</span><span class="sxs-lookup"><span data-stu-id="082f9-116">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="082f9-117">ユーザーに最初の実行エクスペリエンスを 1 回示すか、定期的に示すかを検討することがシナリオにとって重要かどうかを検討します。</span><span class="sxs-lookup"><span data-stu-id="082f9-117">Consider whether showing users the first-run experience once or many times is important to your scenario.</span></span> <span data-ttu-id="082f9-118">例えば、アドインが定期的にのみ活用される場合、ユーザーはアドインにあまり親しんでいない可能性があり、最初の実行エクスペリエンスでもう一度体験することにメリットがある場合があります。</span><span class="sxs-lookup"><span data-stu-id="082f9-118">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="082f9-119">該当する場合、次のパターンを適用してアドインの最初の実行エクスペリエンスを作成し、向上させましょう。</span><span class="sxs-lookup"><span data-stu-id="082f9-119">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="082f9-120">カルーセル</span><span class="sxs-lookup"><span data-stu-id="082f9-120">Carousel</span></span>


<span data-ttu-id="082f9-121">カルーセルは、ユーザーがアドインを使用する前に、ユーザーに一連の特徴や情報ページを表示します。</span><span class="sxs-lookup"><span data-stu-id="082f9-121">Walkthrough takes users through a series of features or information before they start using the add-in. (PDF, code)</span></span>

<span data-ttu-id="082f9-122">*図 1: カルーセルフローで先に進む、または最初のページを飛ばすことができるようにします。*
![最初の実行 - カルーセル - デスクトップ 作業ウィンドウの仕様](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="082f9-122">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="082f9-123">*図 2: ユーザーに表示するカルーセル画面の数を効率的にページを伝えるのに必要なだけの最小限にます*
![最初の実行 - カルーセル - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="082f9-123">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="082f9-124">*図 3: 最初の実行エクスペリエンスを終了するための、明確な操作のきっかけを提供します。*
![最初の実行 - カルーセル - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="082f9-124">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="082f9-125">バリュープレイスマット</span><span class="sxs-lookup"><span data-stu-id="082f9-125">Value Placemat</span></span>

<span data-ttu-id="082f9-126">バリュー プレイスマットは、アドインノ価値提案をロゴを配置することで、価値提案、機能のハイライトまたは概要、およびコール トゥ アクションを明確に伝えます。</span><span class="sxs-lookup"><span data-stu-id="082f9-126">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="082f9-127">![最初の実行 - バリュープレイスマット - デスクトップ作業ウィンドウの仕様](../images/add-in-FRE-value.png)
*ロゴ付きのバリュープレイスマット、明確な価値提案、機能の概要、およびコール トゥ アクション。*</span><span class="sxs-lookup"><span data-stu-id="082f9-127">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="082f9-128">ビデオプレイスマット</span><span class="sxs-lookup"><span data-stu-id="082f9-128">Video Placemat</span></span>

<span data-ttu-id="082f9-129">ビデオプレイスマットは、ユーザーがアドインを使用し始める前にビデオを表示します。</span><span class="sxs-lookup"><span data-stu-id="082f9-129">Video shows users a video before they start using your add-in. (spec, code)</span></span>


<span data-ttu-id="082f9-130">*図 1: 最初の実行プレイスマット - 画面にはビデオからの静止画と、再生ボタンならびに明確なコール トゥ アクションボタンが含まれています。*![ビデオプレイスマット - デスクトップ作業ウィンドウ仕様](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="082f9-130">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="082f9-131">*図 2: ビデオプレーヤー - ユーザーには、ダイアログウィンドウの中にビデオを表示されます。*
![ビデオプレイスマット - デスクトップ作業ウィンドウ仕様](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="082f9-131">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
