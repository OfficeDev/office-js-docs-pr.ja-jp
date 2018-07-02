# <a name="branding-patterns"></a><span data-ttu-id="a52f4-101">ブランディング パターン</span><span class="sxs-lookup"><span data-stu-id="a52f4-101">Branding patterns</span></span>

<span data-ttu-id="a52f4-102">これらのパターンは、アドイン ユーザーにブランドの視認性とコンテキストを提供します。</span><span class="sxs-lookup"><span data-stu-id="a52f4-102">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="a52f4-103">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="a52f4-103">Best practices</span></span>

|<span data-ttu-id="a52f4-104">するべきこと</span><span class="sxs-lookup"><span data-stu-id="a52f4-104">Do</span></span> |<span data-ttu-id="a52f4-105">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="a52f4-105">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="a52f4-106">文字体裁や色など、ブランディング アクセントを適用した使い慣れた UI コンポーネントを使用してください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-106">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="a52f4-107">確立された Office UI と矛盾する新しい UI コンポーネントを考案しないでください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-107">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="a52f4-108">アドインのブランディングを UI の下部にあるブランドバーのフッターに配置してください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-108">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="a52f4-109">UI 上部にあるすぐ隣のブランドバーで作業ウィンドウ名を繰り返さないでください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-109">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="a52f4-110">ブランド要素を多用しないでください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-110">Use brand elements sparingly.</span></span> <span data-ttu-id="a52f4-111">ソリューションが補完的なものとなるように Office に合わせてください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-111">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="a52f4-112">顧客にとって紛らわしく混乱するような、過剰にブランド化した要素を Office UI に挿入しないでください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-112">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="a52f4-113">ソリューションを認識可能にし、一貫性のあるビジュアル要素で画面を接続してください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-113">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="a52f4-114">認識できない、また一貫性なく適用されたビジュアル要素でソリューションを隠さないでください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-114">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="a52f4-115">親サービスまたはビジネスとの接続を構築して、顧客がソリューションを把握し、信頼できるようにしてください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-115">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="a52f4-116">信頼と価値を築くために活用できる有用で理解しやすい関係がある場合は、顧客が新しいブランドコンセプトを習得しないようにしてください。</span><span class="sxs-lookup"><span data-stu-id="a52f4-116">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="a52f4-117">ユーザーがアドインの完全なユーティリティを使用できるように、以下のパターンとコンポーネントを適用します。</span><span class="sxs-lookup"><span data-stu-id="a52f4-117">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="a52f4-118">ブランド バー</span><span class="sxs-lookup"><span data-stu-id="a52f4-118">Brand Bar</span></span>

<span data-ttu-id="a52f4-119">ブランドバーは、ブランド名とロゴを入れるためのフッターにあるスペースです。</span><span class="sxs-lookup"><span data-stu-id="a52f4-119">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="a52f4-120">また、ブランドのウェブサイトへのリンクとオプションのアクセス場所としても機能します。</span><span class="sxs-lookup"><span data-stu-id="a52f4-120">It also serves as a link to your brand's website and an optional access location.</span></span>

![ブランド バー - デスクトップ作業ウィンドウの仕様](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="a52f4-122">スプラッシュ スクリーン</span><span class="sxs-lookup"><span data-stu-id="a52f4-122">Splash Screen</span></span>

<span data-ttu-id="a52f4-123">この画面を使用して、アドインの読み込み中または UI から UI への移行中にブランディングを表示します。</span><span class="sxs-lookup"><span data-stu-id="a52f4-123">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![ブランド スプラッシュ スクリーン - デスクトップ作業ウィンドウの仕様](../images/add-in-splash-screen.png)