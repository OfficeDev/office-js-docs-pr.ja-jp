# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="8db1a-101">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="8db1a-101">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="8db1a-102">Office アドインのユーザーエクスペリエンスの設計では、Office ユーザーが優れたエクスペリエンスを得られるとともに、既定の Office UI 内にシームレスに合致することで、Office の全体的なエクスペリエンスが拡張するようにします。</span><span class="sxs-lookup"><span data-stu-id="8db1a-102">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="8db1a-103">当社の UX パターンはコンポーネントで構成されます。</span><span class="sxs-lookup"><span data-stu-id="8db1a-103">Our UX patterns are composed of components.</span></span> <span data-ttu-id="8db1a-104">コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="8db1a-104">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="8db1a-105">ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。</span><span class="sxs-lookup"><span data-stu-id="8db1a-105">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="8db1a-106">Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。</span><span class="sxs-lookup"><span data-stu-id="8db1a-106">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="8db1a-107">Fabric を活用して、Office とシームレスに統合します。</span><span class="sxs-lookup"><span data-stu-id="8db1a-107">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="8db1a-108">アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="8db1a-108">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="8db1a-109">Office と統合する際に、それを保持する機会を探します。</span><span class="sxs-lookup"><span data-stu-id="8db1a-109">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="8db1a-110">スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルやと動作の採用を行う方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="8db1a-110">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="8db1a-111">規定のパターンは、共通の顧客シナリオとユーザー エクスペリエンスについての調査に基づくベスト プラクティスのソリューションです。</span><span class="sxs-lookup"><span data-stu-id="8db1a-111">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="8db1a-112">このようなパターンにより、アドインの設計と開発を素早く始められるとともに、Microsoft とブランド要素の間のバランスを取るためのガイダンスとしても役立ちます。</span><span class="sxs-lookup"><span data-stu-id="8db1a-112">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="8db1a-113">Microsoft の Fabric デザイン言語のデザイン要素とパートナー固有のブランドの独自性の間のバランスを取る、すっきりしてモダンなユーザー エクスペリエンスによって、ユーザー定着率とアドイン導入率を高められます。</span><span class="sxs-lookup"><span data-stu-id="8db1a-113">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="8db1a-114">UX パターン テンプレートを使用して、次の作業を行います。</span><span class="sxs-lookup"><span data-stu-id="8db1a-114">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="8db1a-115">よくある顧客のシナリオにソリューションとして適用する。</span><span class="sxs-lookup"><span data-stu-id="8db1a-115">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="8db1a-116">設計のベスト プラクティスとして適用する。</span><span class="sxs-lookup"><span data-stu-id="8db1a-116">Apply design best practices.</span></span>
* <span data-ttu-id="8db1a-117">[Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/get-started) のコンポーネントとスタイルを組み込む。</span><span class="sxs-lookup"><span data-stu-id="8db1a-117">Incorporate [Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="8db1a-118">Office の既定の UI に視覚的に溶け込むアドインをビルドする。</span><span class="sxs-lookup"><span data-stu-id="8db1a-118">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="8db1a-119">UX を概念化し、視覚化する。</span><span class="sxs-lookup"><span data-stu-id="8db1a-119">Ideate and visualize UX.</span></span>


## <a name="getting-started"></a><span data-ttu-id="8db1a-120">はじめに</span><span class="sxs-lookup"><span data-stu-id="8db1a-120">Getting started</span></span>

<span data-ttu-id="8db1a-121">パターンは、アドインで共通する主要な操作やエクスペリエンスによって整理されています。</span><span class="sxs-lookup"><span data-stu-id="8db1a-121">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="8db1a-122">主なグループは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8db1a-122">The main differences are:</span></span>

* [<span data-ttu-id="8db1a-123">最初の実行エクスペリエンス  (FRE)</span><span class="sxs-lookup"><span data-stu-id="8db1a-123">First run experience</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="8db1a-124">認証</span><span class="sxs-lookup"><span data-stu-id="8db1a-124">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="8db1a-125">ナビゲーション</span><span class="sxs-lookup"><span data-stu-id="8db1a-125">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="8db1a-126">ブランド化デザイン</span><span class="sxs-lookup"><span data-stu-id="8db1a-126">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="8db1a-127">各グループを確認して、ベスト プラクティスを使用してアドインを設計する方法を理解してください。</span><span class="sxs-lookup"><span data-stu-id="8db1a-127">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>



><span data-ttu-id="8db1a-128">注記: この文書全体で示されている画面の例は、**1366x768** の解像度で設計し、表示されています。</span><span class="sxs-lookup"><span data-stu-id="8db1a-128">NOTE: The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**</span></span>




## <a name="see-also"></a><span data-ttu-id="8db1a-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="8db1a-129">See also</span></span>
* [<span data-ttu-id="8db1a-130">デザイン ツールキット</span><span class="sxs-lookup"><span data-stu-id="8db1a-130">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="8db1a-131">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="8db1a-131">Office UI Fabric</span></span>](https://developer.microsoft.com/en-us/fabric)
* [<span data-ttu-id="8db1a-132">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="8db1a-132">Best practices for developing Office Add-ins</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="8db1a-133">Fabric React の使用を開始する</span><span class="sxs-lookup"><span data-stu-id="8db1a-133">name: Get started using Fabric React href: design/using-office-ui-fabric-react.md</span></span>](https://docs.microsoft.com/en-us/office/dev/add-ins/design/using-office-ui-fabric-react)
