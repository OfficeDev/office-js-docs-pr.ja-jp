---
title: Office アドインのデータ可視化のスタイル ガイドライン
description: Office アドインでデータを表示する方法について、適切な方法を紹介します。
ms.date: 01/14/2019
localization_priority: Normal
ms.openlocfilehash: f3fa2a6cc5a9d27135ad4290eded838dfaecb7d6
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132075"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="a8246-103">Office アドインのデータ可視化のスタイル ガイドライン</span><span class="sxs-lookup"><span data-stu-id="a8246-103">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="a8246-p101">データ可視化が良好なら、ユーザーはデータから洞察が得やすくなります。ユーザーは、これらの洞察を使って通知や説得の話ができます。この記事では、Excel やその他の Office アプリ用のアドインで効果的なデータ可視化を設計するためのガイドラインを示します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="a8246-p102">データ可視化のクロムを作成するには、[Office UI Fabric](https://developer.microsoft.com/fabric) を使用することをお勧めします。Office UI Fabric には、Office の外観とシームレスに統合するスタイルとコンポーネントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a8246-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span>
<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="a8246-109">データ可視化の要素</span><span class="sxs-lookup"><span data-stu-id="a8246-109">Data visualization elements</span></span>

<span data-ttu-id="a8246-110">データビジュアライゼーションは、次の図に示すように、一般的なフレームワークと、タイトル、ラベル、データプロットなどの一般的なビジュアル要素と対話要素を共有します。</span><span class="sxs-lookup"><span data-stu-id="a8246-110">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.</span></span>

![タイトル、軸、凡例、ラベル付きプロットエリアが付いた折れ線グラフ](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a><span data-ttu-id="a8246-112">グラフのタイトル</span><span class="sxs-lookup"><span data-stu-id="a8246-112">Chart titles</span></span>

<span data-ttu-id="a8246-113">グラフのタイトルに関する次のガイドラインに従います。</span><span class="sxs-lookup"><span data-stu-id="a8246-113">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="a8246-p103">グラフのタイトルを見やすくします。グラフの残りの部分との階層関係を視覚ではっきり示すように配置します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="a8246-p104">一般に、文頭だけを大文字にします (最初の単語の最初の文字を大文字にします)。コントラストを付けたり、階層を明確にしたりするには、すべて大文字を使用できますが、控えめに使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8246-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="a8246-p105">[Office UI Fabric の文字体裁](https://developer.microsoft.com/fabric#/styles/typography)を組み込み、グラフを Segoe を使用する Office UI と一貫性をもたせます。グラフのコンテンツを UI と区別するために、異なる書体を使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="a8246-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="a8246-120">カウンターの大きい sans-serif 書体を使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-120">Use sans-serif typefaces with large counters.</span></span>

### <a name="axis-labels"></a><span data-ttu-id="a8246-121">軸ラベル</span><span class="sxs-lookup"><span data-stu-id="a8246-121">Axis labels</span></span>

<span data-ttu-id="a8246-p106">テキスト色と背景色のコントラスト比を適正に保ちつつ、軸ラベルをはっきり読める程度にまで濃くします。データ インクと張り合うほど濃くしません。</span><span class="sxs-lookup"><span data-stu-id="a8246-p106">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="a8246-124">軸のラベルには明るいグレーが最も効果的です。</span><span class="sxs-lookup"><span data-stu-id="a8246-124">Light grays are most effective for axis labels.</span></span> <span data-ttu-id="a8246-125">Fabric を使用している場合は、[ [ニュートラルカラー] パレット](https://developer.microsoft.com/fabric#/styles/colors)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8246-125">If you're using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="a8246-126">データ インク</span><span class="sxs-lookup"><span data-stu-id="a8246-126">Data ink</span></span>

<span data-ttu-id="a8246-p108">グラフの実際のデータを表すピクセルをデータ インクと言います。これは可視化で最も重点が置かれるものです。影付き、太いアウトライン、またはデータをゆがめたり、データと張り合ったりする不要なデザイン要素の使用は避けてください。グラデーションを使用するのは、データ値が色の値と関連する場合だけにします。測定可能な対象値が三次元に結び付けられていない限り、三次元のグラフは避けてください。</span><span class="sxs-lookup"><span data-stu-id="a8246-p108">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="a8246-132">色</span><span class="sxs-lookup"><span data-stu-id="a8246-132">Color</span></span>

<span data-ttu-id="a8246-p109">ハードコードされた色ではなく、オペレーティング システムまたはアプリケーションのテーマに沿った色を選びます。同時に、適用する色がデータをゆがめないようにします。データ可視化で誤って色を使用すると、データがゆがめられて、情報が間違って伝わることがあります。</span><span class="sxs-lookup"><span data-stu-id="a8246-p109">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="a8246-136">データ可視化における色の使用のベスト プラクティスについては、次をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="a8246-136">For best practices for use of color in data visualizations, see the following:</span></span>

- [<span data-ttu-id="a8246-137">なぜ虹色はデータの可視化に適していないか</span><span class="sxs-lookup"><span data-stu-id="a8246-137">Why rainbow colors aren't the best option for data visualizations</span></span>](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="a8246-138">Color Brewer 2.0:地図作成の色のアドバイス</span><span class="sxs-lookup"><span data-stu-id="a8246-138">Color Brewer 2.0: Color Advice for Cartography</span></span>](https://colorbrewer2.org/)
- [<span data-ttu-id="a8246-139">色相が必要だ</span><span class="sxs-lookup"><span data-stu-id="a8246-139">I Want Hue</span></span>](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="a8246-140">枠線</span><span class="sxs-lookup"><span data-stu-id="a8246-140">Gridlines</span></span>

<span data-ttu-id="a8246-p110">グラフを正確に読み取るために目盛線が必要な場合もありますが、データ インクを引き立てる (データ インクと競合しない) 2 次的なビジュアル要素でなければなりません。静的な目盛線は特にハイ コントラスト用にデザインされたものでなければ、細く明るい色にします。また、ユーザーがグラフを対話的に使用するときにコンテキストに沿って現れる、その場限りの動的な目盛線を対話的操作によって作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="a8246-p110">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="a8246-144">目盛線には明るいグレーが最も効果的です。</span><span class="sxs-lookup"><span data-stu-id="a8246-144">Light grays are most effective for gridlines.</span></span> <span data-ttu-id="a8246-145">Fabric を使用している場合は、[ [ニュートラルカラー] パレット](https://developer.microsoft.com/fabric#/styles/colors)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8246-145">If you're using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="a8246-146">次の図は、目盛線のあるデータ可視化を示しています。</span><span class="sxs-lookup"><span data-stu-id="a8246-146">The following image shows a data visualization with gridlines.</span></span>

![グリッド線付き折れ線グラフのデータビジュアライゼーション](../images/data-visualization.png)

### <a name="legends"></a><span data-ttu-id="a8246-148">凡例</span><span class="sxs-lookup"><span data-stu-id="a8246-148">Legends</span></span>

<span data-ttu-id="a8246-149">次が必要な場合は、凡例を追加します。</span><span class="sxs-lookup"><span data-stu-id="a8246-149">Add legends if necessary to:</span></span>

- <span data-ttu-id="a8246-150">データ系列を区別する</span><span class="sxs-lookup"><span data-stu-id="a8246-150">Distinguish between series</span></span>
- <span data-ttu-id="a8246-151">目盛または値の変化を示す</span><span class="sxs-lookup"><span data-stu-id="a8246-151">Present scale or value changes</span></span>

<span data-ttu-id="a8246-p112">凡例がデータ インクを引き立てるようにし、データ インクと競合しないようにしてください。次のように凡例を配置します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p112">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="a8246-154">凡例項目がすべてグラフの上に収まる場合は、プロット エリアを既定で左揃えにします。</span><span class="sxs-lookup"><span data-stu-id="a8246-154">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="a8246-155">一部の凡例項目がグラフの上に収まらない場合は、プロット エリアの右上に配置し、必要に応じてスクロール可能にします。</span><span class="sxs-lookup"><span data-stu-id="a8246-155">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="a8246-p113">読みやすさとアクセシビリティを最適化するには、凡例のマーカーを関連するグラフの図形に合わせます。たとえば、散布図とバブルチャートの凡例には円形の凡例マーカーを使用します。折れ線グラフには線分の凡例マーカーを使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p113">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="a8246-159">データ ラベルとヒント</span><span class="sxs-lookup"><span data-stu-id="a8246-159">Data labels and tooltips</span></span>

<span data-ttu-id="a8246-p114">データ ラベルとヒントの空白スペースと活字バリエーションが十分であることを確認します。オクルージョンと競合を最小限にするアルゴリズムを使用します。たとえば、既定ではデータ ポイントの右側にヒントを表示するものの、右端が検出された場合は左側に表示するなどです。</span><span class="sxs-lookup"><span data-stu-id="a8246-p114">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="a8246-163">デザインの原則</span><span class="sxs-lookup"><span data-stu-id="a8246-163">Design principles</span></span>

<span data-ttu-id="a8246-164">次に示す一連のデザインの原則は Office の設計チームによって作成されたものであり、Office 製品スイートのデータ可視化を新たに設計するときに使用されているものです。</span><span class="sxs-lookup"><span data-stu-id="a8246-164">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="a8246-165">ビジュアル デザインの原則</span><span class="sxs-lookup"><span data-stu-id="a8246-165">Visual design principles</span></span>

- <span data-ttu-id="a8246-p115">可視化では、データを優先し、これを引き立てて理解しやすくする必要があります。コンテキストを示すために必要な分だけサポート要素を追加し、データを強調します。不要な装飾 (影付きやアウトラインなど) や無意味なグラフ、データの歪みは避けます。</span><span class="sxs-lookup"><span data-stu-id="a8246-p115">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="a8246-p116">可視化は、調査を促す十分な視覚的フィードバックを返す必要があります。確立した対話的操作のパターン、インターフェイスのコントロール、明確なシステム フィードバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p116">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="a8246-p117">古くからあるデザイン原則を具体化します。形式、読みやすさ、意味を強化するため、文字体裁と視覚伝達のための定評あるデザイン原則を使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p117">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="a8246-173">対話的操作のデザイン原則</span><span class="sxs-lookup"><span data-stu-id="a8246-173">Interaction design principles</span></span>

- <span data-ttu-id="a8246-174">調査を考慮に入れてデザインします。</span><span class="sxs-lookup"><span data-stu-id="a8246-174">Design to allow for exploration.</span></span>
- <span data-ttu-id="a8246-175">新しい洞察をもたらす、オブジェクトとの直接の対話的操作 (たとえばドラッグで並べ替え) を考慮に入れます。</span><span class="sxs-lookup"><span data-stu-id="a8246-175">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="a8246-176">単純で直接的な、慣れ親しんだ対話的操作モデルを使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-176">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="a8246-177">使いやすい対話型のデータ可視化をデザインする方法については、「[UI の原則と落とし穴](https://uitraps.com/)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="a8246-177">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="a8246-178">モーション デザインの原則</span><span class="sxs-lookup"><span data-stu-id="a8246-178">Motion design principles</span></span>

<span data-ttu-id="a8246-p118">モーションは外部からの操作に従います。ビジュアル要素は、同じ方向に同じ速度で移動する必要があります。適用対象は以下のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a8246-p118">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="a8246-182">チャートの作成</span><span class="sxs-lookup"><span data-stu-id="a8246-182">Chart creation</span></span>
- <span data-ttu-id="a8246-183">1 つのグラフの種類から別のグラフの種類への移行</span><span class="sxs-lookup"><span data-stu-id="a8246-183">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="a8246-184">フィルター</span><span class="sxs-lookup"><span data-stu-id="a8246-184">Filtering</span></span>
- <span data-ttu-id="a8246-185">並べ替え</span><span class="sxs-lookup"><span data-stu-id="a8246-185">Sorting</span></span>
- <span data-ttu-id="a8246-186">データの追加または削除</span><span class="sxs-lookup"><span data-stu-id="a8246-186">Adding or subtracting data</span></span>
- <span data-ttu-id="a8246-187">データのブラッシングまたはスライス</span><span class="sxs-lookup"><span data-stu-id="a8246-187">Brushing or slicing data</span></span>
- <span data-ttu-id="a8246-188">グラフのサイズ変更</span><span class="sxs-lookup"><span data-stu-id="a8246-188">Resizing a chart</span></span>

<span data-ttu-id="a8246-p119">因果関係を知覚できるようにします。アニメーションをステージングする場合には、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="a8246-p119">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="a8246-191">一度に 1 つだけステージングします。</span><span class="sxs-lookup"><span data-stu-id="a8246-191">Stage one thing at a time.</span></span>
- <span data-ttu-id="a8246-192">データ インクの変更より前に、軸の変更をステージングします。</span><span class="sxs-lookup"><span data-stu-id="a8246-192">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="a8246-193">複数のオブジェクトが同じ速度で同じ方向に向かって移動している場合は、グループとしてステージングおよびアニメーション処理します。</span><span class="sxs-lookup"><span data-stu-id="a8246-193">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="a8246-p120">データ要素をステージングするグループのオブジェクト数はせいぜい 4 から 5 個とします。4 から 5 個を超えると、見る人がオブジェクトを個別に追跡しにくくなります。</span><span class="sxs-lookup"><span data-stu-id="a8246-p120">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="a8246-196">モーションは意味を付け加えます。</span><span class="sxs-lookup"><span data-stu-id="a8246-196">Motion adds meaning.</span></span>

- <span data-ttu-id="a8246-197">アニメーションは、ユーザーがデータの変化をより良く理解できるようにしたり、コンテキストを示したり、言語によらない注釈層として機能したりします。</span><span class="sxs-lookup"><span data-stu-id="a8246-197">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="a8246-198">モーションは、意味のある可視化の座標空間で行わなければなりません。</span><span class="sxs-lookup"><span data-stu-id="a8246-198">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="a8246-199">アニメーションはビジュアルに合わせます。</span><span class="sxs-lookup"><span data-stu-id="a8246-199">Tailor the animation to the visual.</span></span>
- <span data-ttu-id="a8246-200">余計なアニメーションは避けてください。</span><span class="sxs-lookup"><span data-stu-id="a8246-200">Avoid gratuitous animations.</span></span>

<span data-ttu-id="a8246-201">モーションはデータに従います。</span><span class="sxs-lookup"><span data-stu-id="a8246-201">Motion follows data.</span></span>

- <span data-ttu-id="a8246-p121">データのマッピングを保持します。測定単位に関係する領域があるなら、切り替え中にその領域を保持します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p121">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="a8246-p122">一貫性のあるアニメーション デザインの言語を保持します。できれば、データ可視化アニメーションを既存の Office モーション デザイン言語にマップします。類似するグラフ タイプには、類似のアニメーションを使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p122">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="a8246-207">データ可視化におけるアクセシビリティ</span><span class="sxs-lookup"><span data-stu-id="a8246-207">Accessibility in data visualizations</span></span>

- <span data-ttu-id="a8246-p123">情報を伝達する唯一の手段として色を使用することはしないでください。色覚異常がある場合、結果がわからなくなってしまいます。できれば、色だけでなく、形状、サイズ、テクスチャを情報の伝達に使用します。</span><span class="sxs-lookup"><span data-stu-id="a8246-p123">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="a8246-211">プッシュ ボタンやピック リストなど、すべての対話型要素をキーボードからアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="a8246-211">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="a8246-212">フォーカスの変更、ヒントなどを通知するため、アクセシビリティ イベントをスクリーン リーダーに送信します。</span><span class="sxs-lookup"><span data-stu-id="a8246-212">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="a8246-213">関連項目</span><span class="sxs-lookup"><span data-stu-id="a8246-213">See also</span></span>

- [<span data-ttu-id="a8246-214">データ可視化を構築するための 5 つの最適なライブラリ</span><span class="sxs-lookup"><span data-stu-id="a8246-214">The Five Best Libraries for Building Data Visualizations</span></span>](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="a8246-215">定量的情報のビジュアル表示</span><span class="sxs-lookup"><span data-stu-id="a8246-215">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
