---
title: アドイン コマンドのアイコンをデザインする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: a3dc7837bdc95df9576a5fc4a6c1840e64afacb6
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437480"
---
# <a name="design-icons-for-add-in-commands"></a><span data-ttu-id="1e50a-102">アドイン コマンドのアイコンをデザインする</span><span class="sxs-lookup"><span data-stu-id="1e50a-102">Design icons for add-in commands</span></span>

<span data-ttu-id="1e50a-p101">[アドイン コマンド](add-in-commands.md)は、Office UI にボタン、テキスト、およびアイコンを追加します。アドイン コマンドのボタンには、ユーザーがコマンドを使うときに、実行しようとするアクションを明確に識別できる、分かりやすいアイコンとラベルをつける必要があります。この記事では、Office とシームレスに統合するアイコンをデザインするための、スタイルと運用に関するガイドラインを提示します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p101">[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI. Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command. This article provides stylistic and production guidelines that help you design icons that integrate seamlessly with Office.</span></span> 

## <a name="office-icon-design-principles"></a><span data-ttu-id="1e50a-106">Office のアイコン デザインの原則</span><span class="sxs-lookup"><span data-stu-id="1e50a-106">Office icon design principles</span></span>

<span data-ttu-id="1e50a-p102">Office のデスクトップ クライアントの Office 2013 リリースでは、図像が更新されています。スタイルについての大きな変更は、内容が削減されたことです。新しいアイコンには、コミュニケーションに関する不可欠の要素だけが含まれてます。遠近法、グラデーション、および光源など、重要でない要素が削除されています。アイコンが簡略化されたことで、コマンドやコントロールの解析をより高速に行うことができるようになっています。このスタイルは、Office に最適です。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p102">The Office 2013 release of the Office desktop clients includes refreshed iconography. The overriding stylistic change is reduction. The new icons include only essential communicative elements. Non-essential elements including perspective, gradients, and light source are removed. The simplified icons support faster parsing of commands and controls. Follow this style to best fit with Office.</span></span>

<span data-ttu-id="1e50a-113">Office のアイコンは、次のデザインの原則に基づいています。</span><span class="sxs-lookup"><span data-stu-id="1e50a-113">Office icons are based on the following design principles:</span></span> 

- <span data-ttu-id="1e50a-114">Office のアイコン コレクションに関する現在の採用方針</span><span class="sxs-lookup"><span data-stu-id="1e50a-114">Modern interpretation of Office icon collection</span></span> 
- <span data-ttu-id="1e50a-115">新鮮で、かつなじみ深いもの</span><span class="sxs-lookup"><span data-stu-id="1e50a-115">Fresh yet familiar</span></span>  
- <span data-ttu-id="1e50a-116">シンプルで、わかりやすく、直接的</span><span class="sxs-lookup"><span data-stu-id="1e50a-116">Simple, clear, and direct</span></span> 

<span data-ttu-id="1e50a-117">以下の図は、現在のデザインの原則を適用したアイコンです。</span><span class="sxs-lookup"><span data-stu-id="1e50a-117">The following image shows icons that apply the modern design principles.</span></span>

![Office の古いアイコンと、現在の採用方針に基づいて更新されたアイコンを示す図](../images/icons-images.png)

## <a name="icon-guidelines"></a><span data-ttu-id="1e50a-119">アイコン ガイドライン</span><span class="sxs-lookup"><span data-stu-id="1e50a-119">Icon guidelines</span></span>
<span data-ttu-id="1e50a-120">アイコンを作成するときは、以下のガイドラインに従ってください。</span><span class="sxs-lookup"><span data-stu-id="1e50a-120">Follow these guidelines when you create your icons:</span></span> 

- <span data-ttu-id="1e50a-121">最適な状態に仕上げるため、必ず 1px グリッドにし、ビットマップ編集ツールを使用します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-121">Stick to the 1px grid and use a bitmap editing tool for best results.</span></span>  
- <span data-ttu-id="1e50a-p103">サイズ変更ではなく、再描画してください。アイコンのサイズを大きくしたり小さくしたりする場合は、カットアウト、角、および丸角のエッジの線をできる限り明瞭に出すために、再描画を行う手間を省かないでください。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p103">Redraw, don't resize. As you resize your icons for larger or smaller sizes, take the time to redraw cutouts, corners, and rounded edges to maximize line clarity.</span></span> 
- <span data-ttu-id="1e50a-124">アイコンを乱雑に見せる成果物は削除します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-124">Remove artifacts that make your icon look messy.</span></span>
- <span data-ttu-id="1e50a-p104">Office のリボンまたはコンテキスト メニューにある Office UI Fabric のアイコンは、再利用しないでください。Fabric のアイコンはスタイルが異なるので、適合しません。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p104">Don't reuse Office UI Fabric icons in the Office ribbon or contextual menu. Fabric icons are stylistically different and will not match.</span></span> 
- <span data-ttu-id="1e50a-p105">アドイン コマンドで何をするかを伝えるために、ロゴやブランドに頼らないようにします。ブランド マークは、サイズの小さいアイコンにしたり、修飾子を適用したりすると、しばしば認識不可能になります。ブランド マークは、多くの場合、Office のリボン アイコンのスタイルと競合し、アイコンがたくさんある環境ではユーザーの関心を奪い合うおそれがあります。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p105">Avoid relying on your logo or brand to communicate what an add-in command does. Brand marks aren't always recognizable at smaller icon sizes and when modifiers are applied. Brand marks often conflict with Office ribbon icon styles, and can compete for user attention in a saturated environment.</span></span>
- <span data-ttu-id="1e50a-p106">アクセスしやすくするために、白の塗りつぶしを使います。アイコンのオブジェクトは、Office UI のテーマのハイ コントラスト モードで読みやすさを保つために、たいていは背景を白にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p106">Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.</span></span>  
- <span data-ttu-id="1e50a-132">透明背景の PNG 形式を使用します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-132">Use the PNG format with a transparent background.</span></span> 
- <span data-ttu-id="1e50a-133">アイコンに、表記文字、段落のラグ、および疑問符などの、ローカライズ可能なコンテンツを含めないようにします。</span><span class="sxs-lookup"><span data-stu-id="1e50a-133">Avoid localizable content in your icons, including typographic characters, indications of paragraph rags, and question marks.</span></span> 
- <span data-ttu-id="1e50a-p107">さまざまなコマンドで、同じ視覚的メタファーを再利用しないようにします。さまざまなアクションに同じアイコンを使用すると、混乱を招くおそれがあります。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p107">Don't reuse visual metaphors for different commands. Using the same icon for different actions can cause confusion.</span></span> 
- <span data-ttu-id="1e50a-p108">ボタンのラベルを明確で簡潔なものにします。意味を伝えるには、ビジュアルとテキストの情報を組み合わせます。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p108">Make your button labels clear and succinct. Use a combination of visual and textual information to convey meaning.</span></span> 


## <a name="icon-size-recommendations-and-requirements"></a><span data-ttu-id="1e50a-138">アイコン サイズについて推奨と要件</span><span class="sxs-lookup"><span data-stu-id="1e50a-138">Icon size recommendations and requirements</span></span>

<span data-ttu-id="1e50a-p109">Office 2016 のデスクトップ アイコンは、ビットマップ画像です。ユーザーの DPI 設定やタッチ モードによって異なるサイズで表示されます。サポートされている 8 つのサイズすべてを組み込んで、すべての解像度とコンテキストで最高のエクスペリエンスを提供します。以下のサイズがサポートされています (うち 3 つは必須)：</span><span class="sxs-lookup"><span data-stu-id="1e50a-p109">Office 2016 desktop icons are bitmap images. Different sizes will render depending on the user's DPI setting and touch mode. Include all eight supported sizes to create the best experience in all supported resolutions and contexts. The following are the supported sizes - three are required:</span></span>

- <span data-ttu-id="1e50a-143">16 px (必須)</span><span class="sxs-lookup"><span data-stu-id="1e50a-143">16 px (Required)</span></span>
- <span data-ttu-id="1e50a-144">20 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-144">20 px</span></span>
- <span data-ttu-id="1e50a-145">24 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-145">24 px</span></span>
- <span data-ttu-id="1e50a-146">32 px (必須)</span><span class="sxs-lookup"><span data-stu-id="1e50a-146">32 px (Required)</span></span>
- <span data-ttu-id="1e50a-147">40 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-147">40 px</span></span>
- <span data-ttu-id="1e50a-148">48 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-148">48 px</span></span>
- <span data-ttu-id="1e50a-149">64 ピクセル (推奨、Mac に最適)</span><span class="sxs-lookup"><span data-stu-id="1e50a-149">64 px (Recommended, best for Mac)</span></span>
- <span data-ttu-id="1e50a-150">80 px (必須)</span><span class="sxs-lookup"><span data-stu-id="1e50a-150">80 px (Required)</span></span>  

<span data-ttu-id="1e50a-151">それぞれのアイコンを、サイズに合わせて縮小するのではなく再描画します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-151">Make sure to redraw your icons for each size rather than shrink them to fit.</span></span>

![アイコンの縮小ではなく、アイコンのサイズ変更が推奨されていることを示す図](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a><span data-ttu-id="1e50a-153">アイコンの構造とレイアウト</span><span class="sxs-lookup"><span data-stu-id="1e50a-153">Icon anatomy and layout</span></span>

<span data-ttu-id="1e50a-p110">Office のアイコンは、基本要素に、アクション修飾子と概念的修飾子を重ね合わせた構成になっています。 アクション修飾子は、追加、開く、新規、閉じるなどの概念を表します。概念的修飾子は、ステータス、変更、またはアイコンの説明を表します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p110">Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.</span></span> 

<span data-ttu-id="1e50a-p111">Office UI と協調するコマンドを作成するために、基本要素と修飾子のレイアウト ガイドラインに従ってください。これにより、コマンドがプロフェッショナルな仕上がりになり、アドインに対する顧客の信頼度もあがります。場合によっては、意図的にこれらのガイドラインに対して例外を設けることもできます。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p111">To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.</span></span>

<span data-ttu-id="1e50a-160">以下の図は、Office のアイコンの基本要素と修飾子のレイアウトを表しています。</span><span class="sxs-lookup"><span data-stu-id="1e50a-160">The following image shows the layout of base elements and modifiers in an Office icon.</span></span>

![中央にアイコンの基本要素、右下に修飾子、左上にアクション修飾子を配した画像](../images/icon-layouts.png)

- <span data-ttu-id="1e50a-162">基本要素をピクセル フレームの中央に配置し、周囲に余白をとります。</span><span class="sxs-lookup"><span data-stu-id="1e50a-162">Center base elements in the pixel frame with empty padding all around.</span></span>
- <span data-ttu-id="1e50a-163">アクション修飾子は、左上に配置します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-163">Place action modifiers on the top left.</span></span> 
- <span data-ttu-id="1e50a-164">概念的修飾子は、右下に配置します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-164">Place conceptual modifiers on the bottom right.</span></span>
- <span data-ttu-id="1e50a-p112">アイコン内の要素の数を制限します。32px では、修飾子の数を最大 2 つまでに制限します。16px では、修飾子の数を 1 つに制限します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p112">Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.</span></span>

<span data-ttu-id="1e50a-p113">基本要素は、どのサイズでも同じ配置にします。基本要素をフレームの中央に配置できない場合は、左上にそろえ、余分のピクセルは右下に残します。最良の結果を得るために、次の表に示すパディングのガイドラインを適用してください。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p113">Place base elements consistently across sizes. If base elements can't be centered in the frame, align them to the top left, leaving the extra pixels on the bottom right. For best results, apply the padding guidelines listed in the following table.</span></span>

|<span data-ttu-id="1e50a-171">**アイコンのサイズ**</span><span class="sxs-lookup"><span data-stu-id="1e50a-171">**Icon size**</span></span>|<span data-ttu-id="1e50a-172">**基本要素の周囲のパディング**</span><span class="sxs-lookup"><span data-stu-id="1e50a-172">**Padding around base element**</span></span>|
|:---|:---|
|<span data-ttu-id="1e50a-173">16 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-173">16px</span></span>|<span data-ttu-id="1e50a-174">0</span><span class="sxs-lookup"><span data-stu-id="1e50a-174">{0}</span></span>|
|<span data-ttu-id="1e50a-175">20 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-175">20px</span></span>|<span data-ttu-id="1e50a-176">1 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-176">1px</span></span>|
|<span data-ttu-id="1e50a-177">24 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-177">24px</span></span>|<span data-ttu-id="1e50a-178">1 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-178">1px</span></span>|
|<span data-ttu-id="1e50a-179">32 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-179">32px</span></span>|<span data-ttu-id="1e50a-180">2 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-180">2px</span></span>|
|<span data-ttu-id="1e50a-181">40 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-181">40px</span></span>|<span data-ttu-id="1e50a-182">2 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-182">2px</span></span>|
|<span data-ttu-id="1e50a-183">48 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-183">48px</span></span>|<span data-ttu-id="1e50a-184">3 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-184">3px</span></span>|
|<span data-ttu-id="1e50a-185">64 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-185">64px</span></span>|<span data-ttu-id="1e50a-186">5 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-186">5px</span></span>|
|<span data-ttu-id="1e50a-187">80 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-187">80px</span></span>|<span data-ttu-id="1e50a-188">5 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-188">5px</span></span>|

<span data-ttu-id="1e50a-p114">すべての修飾子には、背景を含め、各要素の間に 1px の透明なカットアウトが必要です。要素が直接重ならないようにします。ルールとエッジの間に余白を作ります。修飾子はサイズが少しずつ異なっている場合がありますが、開始点としてこれらのサイズを使用します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p114">All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.</span></span>

|<span data-ttu-id="1e50a-193">**アイコンのサイズ**</span><span class="sxs-lookup"><span data-stu-id="1e50a-193">**Icon size**</span></span>|<span data-ttu-id="1e50a-194">**修飾子のサイズ**</span><span class="sxs-lookup"><span data-stu-id="1e50a-194">**Modifier size**</span></span>|
|:---|:---|
|<span data-ttu-id="1e50a-195">16 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-195">16px</span></span>|<span data-ttu-id="1e50a-196">9 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-196">9px</span></span>|
|<span data-ttu-id="1e50a-197">20 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-197">20px</span></span>|<span data-ttu-id="1e50a-198">10 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-198">10px</span></span>|
|<span data-ttu-id="1e50a-199">24 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-199">24px</span></span>|<span data-ttu-id="1e50a-200">12 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-200">12px</span></span>|
|<span data-ttu-id="1e50a-201">32 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-201">32px</span></span>|<span data-ttu-id="1e50a-202">14 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-202">14px</span></span>|
|<span data-ttu-id="1e50a-203">40 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-203">40px</span></span>|<span data-ttu-id="1e50a-204">20 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-204">20px</span></span>|
|<span data-ttu-id="1e50a-205">48 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-205">48px</span></span>|<span data-ttu-id="1e50a-206">22 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-206">22px</span></span>|
|<span data-ttu-id="1e50a-207">64 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-207">64px</span></span>|<span data-ttu-id="1e50a-208">29 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-208">29px</span></span>|
|<span data-ttu-id="1e50a-209">80 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-209">80px</span></span>|<span data-ttu-id="1e50a-210">38 px</span><span class="sxs-lookup"><span data-stu-id="1e50a-210">38px</span></span>|

## <a name="icon-colors"></a><span data-ttu-id="1e50a-211">アイコンの色</span><span class="sxs-lookup"><span data-stu-id="1e50a-211">Icon colors</span></span>

<span data-ttu-id="1e50a-p115">Office のアイコンには、限定されたカラー パレットがあります。Office UI とのシームレスな統合を保証するために、以下の表に記載されている色を使用してください。色の使用について、以下のガイドラインに従ってください。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p115">Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color:</span></span> 

- <span data-ttu-id="1e50a-p116">色は、装飾のためというよりも、意味を伝える目的のために使用します。アクション、ステータス、または明示的にマークを区別する要素を、色によってハイライトまたは強調します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p116">Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.</span></span>  
- <span data-ttu-id="1e50a-p117">可能であれば、グレー以外の 1 色のみを追加で使用します。追加する色は最大 2 色までに制限します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p117">If possible, use only one additional color beyond gray. Limit additional colors to two at the most.</span></span>
- <span data-ttu-id="1e50a-p118">すべてのサイズのアイコンで、色を統一する必要があります。Office のアイコンのカラー パレットは、アイコンのサイズによってわずかな違いがあります。16px 以下のアイコンでは少し濃く、32px 以上のアイコンではより鮮やかな色になっています。これらの微妙な調整をしないと、サイズによって色の見え方が変わってしまいます。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p118">Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.</span></span>   

|<span data-ttu-id="1e50a-223">**色の名前**</span><span class="sxs-lookup"><span data-stu-id="1e50a-223">**Color name**</span></span>|<span data-ttu-id="1e50a-224">**RGB**</span><span class="sxs-lookup"><span data-stu-id="1e50a-224">**RGB**</span></span>|<span data-ttu-id="1e50a-225">**16 進数**</span><span class="sxs-lookup"><span data-stu-id="1e50a-225">**Hex**</span></span>|<span data-ttu-id="1e50a-226">**色**</span><span class="sxs-lookup"><span data-stu-id="1e50a-226">**Color**</span></span>|<span data-ttu-id="1e50a-227">**分類**</span><span class="sxs-lookup"><span data-stu-id="1e50a-227">**Category**</span></span>|
|:---|:---|:---|:---|:---|
|<span data-ttu-id="1e50a-228">テキスト グレー (80)</span><span class="sxs-lookup"><span data-stu-id="1e50a-228">Text Gray (80)</span></span>|<span data-ttu-id="1e50a-229">80、80、80</span><span class="sxs-lookup"><span data-stu-id="1e50a-229">80, 80, 80</span></span>|<span data-ttu-id="1e50a-230">#505050</span><span class="sxs-lookup"><span data-stu-id="1e50a-230">#505050</span></span>| ![テキスト グレー 80 のカラー イメージ](../images/color-text-gray-80.png) |<span data-ttu-id="1e50a-232">テキスト</span><span class="sxs-lookup"><span data-stu-id="1e50a-232">Text</span></span>|
|<span data-ttu-id="1e50a-233">テキスト グレー (95)</span><span class="sxs-lookup"><span data-stu-id="1e50a-233">Text Gray (95)</span></span>|<span data-ttu-id="1e50a-234">95、95、95</span><span class="sxs-lookup"><span data-stu-id="1e50a-234">95, 95, 95</span></span>|<span data-ttu-id="1e50a-235">#5F5F5F</span><span class="sxs-lookup"><span data-stu-id="1e50a-235">#5F5F5F</span></span>| ![テキスト グレー 95 のカラー イメージ](../images/color-text-gray-95.png) |<span data-ttu-id="1e50a-237">テキスト</span><span class="sxs-lookup"><span data-stu-id="1e50a-237">Text</span></span>|
|<span data-ttu-id="1e50a-238">テキスト グレー (105)</span><span class="sxs-lookup"><span data-stu-id="1e50a-238">Text Gray (105)</span></span>|<span data-ttu-id="1e50a-239">105, 105, 105</span><span class="sxs-lookup"><span data-stu-id="1e50a-239">105, 105, 105</span></span>|<span data-ttu-id="1e50a-240">#696969</span><span class="sxs-lookup"><span data-stu-id="1e50a-240">#696969</span></span>| ![テキスト グレー 105 のカラー イメージ](../images/color-text-gray-105.png) |<span data-ttu-id="1e50a-242">テキスト</span><span class="sxs-lookup"><span data-stu-id="1e50a-242">Text</span></span>|
|<span data-ttu-id="1e50a-243">ダーク グレー 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-243">Dark Gray 32</span></span>|<span data-ttu-id="1e50a-244">128、128、128</span><span class="sxs-lookup"><span data-stu-id="1e50a-244">128, 128, 128</span></span>|<span data-ttu-id="1e50a-245">#808080</span><span class="sxs-lookup"><span data-stu-id="1e50a-245">#808080</span></span>| ![ダーク グレー 32 のカラー イメージ](../images/color-dark-gray-32.png) |<span data-ttu-id="1e50a-247">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-247">32 and above</span></span>|
|<span data-ttu-id="1e50a-248">ミディアム グレー 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-248">Medium Gray 32</span></span>|<span data-ttu-id="1e50a-249">158、158、158</span><span class="sxs-lookup"><span data-stu-id="1e50a-249">158, 158, 158</span></span>|<span data-ttu-id="1e50a-250">#9E9E9E</span><span class="sxs-lookup"><span data-stu-id="1e50a-250">#9E9E9E</span></span>| ![ミディアム グレー 32 のカラー イメージ](../images/color-medium-gray-32.png) |<span data-ttu-id="1e50a-252">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-252">32 and above</span></span>|
|<span data-ttu-id="1e50a-253">ライト グレー オール</span><span class="sxs-lookup"><span data-stu-id="1e50a-253">Light Gray ALL</span></span>|<span data-ttu-id="1e50a-254">179、179、179</span><span class="sxs-lookup"><span data-stu-id="1e50a-254">179, 179, 179</span></span>|<span data-ttu-id="1e50a-255">#B3B3B3</span><span class="sxs-lookup"><span data-stu-id="1e50a-255">#B3B3B3</span></span>| ![ライト グレー オールのカラー イメージ](../images/color-light-gray-all.png) |<span data-ttu-id="1e50a-257">すべてのサイズ</span><span class="sxs-lookup"><span data-stu-id="1e50a-257">All sizes</span></span>|
|<span data-ttu-id="1e50a-258">ダーク グレー 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-258">Dark Gray 16</span></span>|<span data-ttu-id="1e50a-259">114、114、114</span><span class="sxs-lookup"><span data-stu-id="1e50a-259">114, 114, 114</span></span>|<span data-ttu-id="1e50a-260">#727272</span><span class="sxs-lookup"><span data-stu-id="1e50a-260">#727272</span></span>| ![ダーク グレー 16 のカラー イメージ](../images/color-dark-gray-16.png) |<span data-ttu-id="1e50a-262">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-262">16 and below</span></span>|
|<span data-ttu-id="1e50a-263">ミディアム グレー 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-263">Medium Gray 16</span></span>|<span data-ttu-id="1e50a-264">144、144、144</span><span class="sxs-lookup"><span data-stu-id="1e50a-264">144, 144, 144</span></span>|<span data-ttu-id="1e50a-265">#909090</span><span class="sxs-lookup"><span data-stu-id="1e50a-265">#909090</span></span>| ![ミディアム グレー 16 のカラー イメージ](../images/color-medium-gray-16.png) |<span data-ttu-id="1e50a-267">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-267">16 and below</span></span>|
|<span data-ttu-id="1e50a-268">ブルー 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-268">Blue 32</span></span>|<span data-ttu-id="1e50a-269">77、130、184</span><span class="sxs-lookup"><span data-stu-id="1e50a-269">77, 130, 184</span></span>|<span data-ttu-id="1e50a-270">#4d82B8</span><span class="sxs-lookup"><span data-stu-id="1e50a-270">#4d82B8</span></span>| ![ブルー 32 のカラー イメージ](../images/color-blue-32.png) |<span data-ttu-id="1e50a-272">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-272">32 and above</span></span>|
|<span data-ttu-id="1e50a-273">ブルー 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-273">Blue 16</span></span>|<span data-ttu-id="1e50a-274">74、125、177</span><span class="sxs-lookup"><span data-stu-id="1e50a-274">74, 125, 177</span></span>|<span data-ttu-id="1e50a-275">#4A7DB1</span><span class="sxs-lookup"><span data-stu-id="1e50a-275">#4A7DB1</span></span>| ![ブルー 16 のカラー イメージ](../images/color-blue-16.png) |<span data-ttu-id="1e50a-277">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-277">16 and below</span></span>|
|<span data-ttu-id="1e50a-278">イエロー オール</span><span class="sxs-lookup"><span data-stu-id="1e50a-278">Yellow ALL</span></span>|<span data-ttu-id="1e50a-279">234、194、130</span><span class="sxs-lookup"><span data-stu-id="1e50a-279">234, 194, 130</span></span>|<span data-ttu-id="1e50a-280">#EAC282</span><span class="sxs-lookup"><span data-stu-id="1e50a-280">#EAC282</span></span>| ![イエロー オールのカラー イメージ](../images/color-yellow-all.png) |<span data-ttu-id="1e50a-282">すべてのサイズ</span><span class="sxs-lookup"><span data-stu-id="1e50a-282">All sizes</span></span>|
|<span data-ttu-id="1e50a-283">オレンジ 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-283">Orange 32</span></span>|<span data-ttu-id="1e50a-284">231、142、70</span><span class="sxs-lookup"><span data-stu-id="1e50a-284">231, 142, 70</span></span>|<span data-ttu-id="1e50a-285">#E78E46</span><span class="sxs-lookup"><span data-stu-id="1e50a-285">#E78E46</span></span>| ![オレンジ 32 のカラー イメージ](../images/color-orange-32.png) |<span data-ttu-id="1e50a-287">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-287">32 and above</span></span>|
|<span data-ttu-id="1e50a-288">オレンジ 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-288">Orange 16</span></span>|<span data-ttu-id="1e50a-289">227、142、70</span><span class="sxs-lookup"><span data-stu-id="1e50a-289">227, 142, 70</span></span>|<span data-ttu-id="1e50a-290">#E3751C</span><span class="sxs-lookup"><span data-stu-id="1e50a-290">#E3751C</span></span>| ![オレンジ 16 のカラー イメージ](../images/color-orange-16.png) |<span data-ttu-id="1e50a-292">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-292">16 and below</span></span>|
|<span data-ttu-id="1e50a-293">ピンク オール</span><span class="sxs-lookup"><span data-stu-id="1e50a-293">Pink ALL</span></span>|<span data-ttu-id="1e50a-294">230、132、151</span><span class="sxs-lookup"><span data-stu-id="1e50a-294">230, 132, 151</span></span>|<span data-ttu-id="1e50a-295">#E68497</span><span class="sxs-lookup"><span data-stu-id="1e50a-295">#E68497</span></span>| ![ピンク オールのカラー イメージ](../images/color-pink-all.png) |<span data-ttu-id="1e50a-297">すべてのサイズ</span><span class="sxs-lookup"><span data-stu-id="1e50a-297">All sizes</span></span>|
|<span data-ttu-id="1e50a-298">グリーン 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-298">Green 32</span></span>|<span data-ttu-id="1e50a-299">118、167、151</span><span class="sxs-lookup"><span data-stu-id="1e50a-299">118, 167, 151</span></span>|<span data-ttu-id="1e50a-300">#76A797</span><span class="sxs-lookup"><span data-stu-id="1e50a-300">#76A797</span></span>| ![グリーン 32 のカラー イメージ](../images/color-green-32.png) |<span data-ttu-id="1e50a-302">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-302">32 and above</span></span>|
|<span data-ttu-id="1e50a-303">グリーン 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-303">Green 16</span></span>|<span data-ttu-id="1e50a-304">104、164、144</span><span class="sxs-lookup"><span data-stu-id="1e50a-304">104, 164, 144</span></span>|<span data-ttu-id="1e50a-305">#68A490</span><span class="sxs-lookup"><span data-stu-id="1e50a-305">#68A490</span></span>| ![グリーン 16 のカラー イメージ](../images/color-green-16.png) |<span data-ttu-id="1e50a-307">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-307">16 and below</span></span>|
|<span data-ttu-id="1e50a-308">レッド 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-308">Red 32</span></span>|<span data-ttu-id="1e50a-309">216、99、68</span><span class="sxs-lookup"><span data-stu-id="1e50a-309">216, 99, 68</span></span>|<span data-ttu-id="1e50a-310">#D86344</span><span class="sxs-lookup"><span data-stu-id="1e50a-310">#D86344</span></span>| ![レッド 32 のカラー イメージ](../images/color-red-32.png) |<span data-ttu-id="1e50a-312">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-312">32 and above</span></span>|
|<span data-ttu-id="1e50a-313">レッド 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-313">Red 16</span></span>|<span data-ttu-id="1e50a-314">214、85、50</span><span class="sxs-lookup"><span data-stu-id="1e50a-314">214, 85, 50</span></span>|<span data-ttu-id="1e50a-315">#D65532</span><span class="sxs-lookup"><span data-stu-id="1e50a-315">#D65532</span></span>| ![レッド 16 のカラー イメージ](../images/color-red-16.png) |<span data-ttu-id="1e50a-317">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-317">16 and below</span></span>|
|<span data-ttu-id="1e50a-318">パープル 32</span><span class="sxs-lookup"><span data-stu-id="1e50a-318">Purple 32</span></span>|<span data-ttu-id="1e50a-319">152、104、185</span><span class="sxs-lookup"><span data-stu-id="1e50a-319">152, 104, 185</span></span>|<span data-ttu-id="1e50a-320">#9868B9</span><span class="sxs-lookup"><span data-stu-id="1e50a-320">#9868B9</span></span>| ![パープル 32 のカラー イメージ](../images/color-purple-32.png) |<span data-ttu-id="1e50a-322">32 以上</span><span class="sxs-lookup"><span data-stu-id="1e50a-322">32 and above</span></span>|
|<span data-ttu-id="1e50a-323">パープル 16</span><span class="sxs-lookup"><span data-stu-id="1e50a-323">Purple 16</span></span>|<span data-ttu-id="1e50a-324">137、89、171</span><span class="sxs-lookup"><span data-stu-id="1e50a-324">137, 89, 171</span></span>|<span data-ttu-id="1e50a-325">#8959AB</span><span class="sxs-lookup"><span data-stu-id="1e50a-325">#8959AB</span></span>| ![パープル 16 のカラー イメージ](../images/color-purple-16.png) |<span data-ttu-id="1e50a-327">16 以下</span><span class="sxs-lookup"><span data-stu-id="1e50a-327">16 and below</span></span>|


## <a name="icons-in-high-contrast-modes"></a><span data-ttu-id="1e50a-328">ハイコントラスト モードのアイコン</span><span class="sxs-lookup"><span data-stu-id="1e50a-328">Icons in high contrast modes</span></span>

<span data-ttu-id="1e50a-p119">Office のアイコンは、ハイコントラスト モードで適切に表示されるように設計されています。前景の要素は背景と区別され、読みやすさを最大限に高め、色の変更を可能にします。ハイコントラスト モードでは、Office は赤、緑、または青の値が 190 未満のアイコンのすべてのピクセルを、完全な黒に変更します。それ以外のピクセルは、すべて白になります。つまり、各 RGB チャンネルは 0 から 189 の値が黒、190 から 255 の値が白と評価されます。その他のハイコントラスト テーマも同じ 190 値のしきい値を使用して色の変更が行われますが、ルールは異なります。たとえば、白のハイコントラスト テーマでは、190 よりも大きい不透明のピクセルすべての色を変更しますが、その他のピクセルはすべて透明になります。次のガイドラインを適用して、ハイコントラスト設定で読みやすさを最大限にします。</span><span class="sxs-lookup"><span data-stu-id="1e50a-p119">Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:</span></span>

- <span data-ttu-id="1e50a-337">190 値のしきい値に沿って、前景と背景の要素を区別するようにします。</span><span class="sxs-lookup"><span data-stu-id="1e50a-337">Aim to differentiate foreground and background elements along the 190 value threshold.</span></span>
- <span data-ttu-id="1e50a-338">Office アイコンの表示スタイルに従います。</span><span class="sxs-lookup"><span data-stu-id="1e50a-338">Follow Office icon visual styles.</span></span>
- <span data-ttu-id="1e50a-339">色はアイコン パレットから使用します。</span><span class="sxs-lookup"><span data-stu-id="1e50a-339">Use colors from our icon palette.</span></span>
- <span data-ttu-id="1e50a-340">グラデーションの使用を避けます。</span><span class="sxs-lookup"><span data-stu-id="1e50a-340">Avoid the use of gradients.</span></span>
- <span data-ttu-id="1e50a-341">同じ様な値を持つ大きな色のブロックを避けます。</span><span class="sxs-lookup"><span data-stu-id="1e50a-341">Avoid large blocks of color with similar values.</span></span>

## <a name="see-also"></a><span data-ttu-id="1e50a-342">関連項目</span><span class="sxs-lookup"><span data-stu-id="1e50a-342">See also</span></span>

- [<span data-ttu-id="1e50a-343">アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1e50a-343">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="1e50a-344">Excel、Word、PowerPoint のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="1e50a-344">Add-in commands for Excel, Word, and PowerPoint</span></span>](../design/add-in-commands.md)
