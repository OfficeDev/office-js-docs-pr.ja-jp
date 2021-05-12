---
title: Office アドインのレイアウト ガイドライン
description: アドインで作業ウィンドウまたはダイアログをレイアウトする方法のOfficeを取得します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 1eea665028abc90b2361edae45e81bc85481a429
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330011"
---
# <a name="layout"></a><span data-ttu-id="61f24-103">レイアウト</span><span class="sxs-lookup"><span data-stu-id="61f24-103">Layout</span></span>

<span data-ttu-id="61f24-p101">Office に埋め込まれている各 HTML コンテナーは、レイアウトを持つことになります。これらのレイアウトは、アドインのメイン画面です。そこでは、お客様による操作の開始、設定の変更、表示、スクロール、コンテンツ間の移動を可能にするエクスペリエンスを作成します。エクスペリエンスの継続性を保証するために、画面全体のレイアウトが整合性のあるアドインを設計します。お客様が使い慣れている既存の Web サイトがある場合、既存の Web ページのレイアウトを再利用することを検討してください。Office HTML コンテナー内に調和よく収まるようにそれらを適合させます。</span><span class="sxs-lookup"><span data-stu-id="61f24-p101">Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you will create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.</span></span>

<span data-ttu-id="61f24-110">レイアウトのガイドラインについては、[作業ウィンドウ](task-pane-add-ins.md)、[コンテンツ](content-add-ins.md)、[ダイアログ ボックス](dialog-boxes.md)に関する記事をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="61f24-110">For guidelines on layout, see [Task pane](task-pane-add-ins.md), [Content](content-add-ins.md), and [Dialog box](dialog-boxes.md).</span></span> <span data-ttu-id="61f24-111">Fluent UI React [、または](using-office-ui-fabric-react.md) [Office UI Fabric JS](fabric-core.md)、コンポーネントを一般的なレイアウトとユーザー エクスペリエンス フローに組み立てる方法の詳細については、「UX デザイン パターン テンプレート」を[参照してください](ux-design-pattern-templates.md)。</span><span class="sxs-lookup"><span data-stu-id="61f24-111">For more information about how to assemble [Fluent UI React](using-office-ui-fabric-react.md), or [Office UI Fabric JS](fabric-core.md), components into common layouts and user experience flows, see [UX design patterns templates](ux-design-pattern-templates.md).</span></span>

<span data-ttu-id="61f24-112">レイアウトについて、次の一般的なガイドラインが適用されます。</span><span class="sxs-lookup"><span data-stu-id="61f24-112">Apply the following general guidelines for layouts:</span></span>

*   <span data-ttu-id="61f24-p103">HTML コンテナーでは、狭い余白や広い余白は使用しないでください。20 ピクセルが最適な既定値です。</span><span class="sxs-lookup"><span data-stu-id="61f24-p103">Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default.</span></span>
*   <span data-ttu-id="61f24-p104">要素を意図的に配置します。追加のインデントや新しい配置点は視覚的な階層をより明確にするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="61f24-p104">Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.</span></span>
*   <span data-ttu-id="61f24-p105">Office インターフェイスは、4 px グリッド線上にあります。要素間のスペースを 4 の倍数に維持することを目標にします。</span><span class="sxs-lookup"><span data-stu-id="61f24-p105">Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4.</span></span>
*   <span data-ttu-id="61f24-119">インタ フェースに詰め込みすぎると、混乱を招き、タッチ操作が使いにくくなります。</span><span class="sxs-lookup"><span data-stu-id="61f24-119">Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions.</span></span>
*   <span data-ttu-id="61f24-p106">画面全体でレイアウトの整合性を保ちます。予期しないレイアウトの変更は、ソリューションの信頼と信用を失う一因となるビジュアルのバグのようになります。</span><span class="sxs-lookup"><span data-stu-id="61f24-p106">Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution.</span></span>
*   <span data-ttu-id="61f24-p107">一般的なレイアウトのパターンに従います。規則は、ユーザーがインターフェイスの使用方法を理解するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="61f24-p107">Follow common layout patterns. Conventions help users understand how to use an interface.</span></span>
*   <span data-ttu-id="61f24-124">ブランドやコマンドのような要素を冗長に使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="61f24-124">Avoid redundant elements like branding or commands.</span></span>
*   <span data-ttu-id="61f24-125">コントロールとビューを統合して、マウスを移動させる回数を減らします。</span><span class="sxs-lookup"><span data-stu-id="61f24-125">Consolidate controls and views to avoid requiring too much mouse movement.</span></span>
*   <span data-ttu-id="61f24-126">HTML コンテナーの幅と高さに適合する応答性の高いエクスペリエンスを作成します。</span><span class="sxs-lookup"><span data-stu-id="61f24-126">Create responsive experiences that adapt to HTML container widths and heights.</span></span>
