---
title: Office アドインの Monoline スタイルのアイコンガイドライン
description: Office アドインで Monoline スタイルアイコンアイコンを使用するためのガイドラインを取得します。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 264aa9e01bd70924cfee01a864c515c8c7a4d138
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132201"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="44d19-103">Office アドインの Monoline スタイルのアイコンガイドライン</span><span class="sxs-lookup"><span data-stu-id="44d19-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="44d19-104">Monoline style 図像は Office 365 で使用されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-104">Monoline style iconography are used in Office 365.</span></span> <span data-ttu-id="44d19-105">アイコンがサブスクリプション以外の Office 2013 以降の新しいスタイルに一致するようにする場合は、「 [Office アドインの新しいスタイルのアイコンガイドライン](add-in-icons-fresh.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44d19-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="44d19-106">Office Monoline の視覚スタイル</span><span class="sxs-lookup"><span data-stu-id="44d19-106">Office Monoline visual style</span></span>

<span data-ttu-id="44d19-107">一貫性があり、わかりやすく、アクセス可能な図像を持つ Monoline スタイルの目的は、シンプルなビジュアルを使用して操作と機能を伝えるために、アイコンがすべてのユーザーに対してアクセス可能であること、および Windows の他の場所で使用されているものと一致するスタイルを持つことを示します。</span><span class="sxs-lookup"><span data-stu-id="44d19-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="44d19-108">次のガイドラインは、サードパーティの開発者が、既にインストールされているアイコンと一貫性のある機能のアイコンを作成することを希望しています。</span><span class="sxs-lookup"><span data-stu-id="44d19-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="44d19-109">設計原則</span><span class="sxs-lookup"><span data-stu-id="44d19-109">Design principles</span></span>

- <span data-ttu-id="44d19-110">シンプル、クリーン、クリア。</span><span class="sxs-lookup"><span data-stu-id="44d19-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="44d19-111">必要な要素のみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="44d19-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="44d19-112">ウィンドウアイコンのスタイル。</span><span class="sxs-lookup"><span data-stu-id="44d19-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="44d19-113">すべてのユーザーがアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="44d19-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="44d19-114">意味を伝える</span><span class="sxs-lookup"><span data-stu-id="44d19-114">Conveying meaning</span></span>

- <span data-ttu-id="44d19-115">ページなどの説明的な要素を使用して、メールを表すドキュメントまたはエンベロープを表します。</span><span class="sxs-lookup"><span data-stu-id="44d19-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="44d19-116">同じ概念を表すのと同じ要素を使用します。つまり、メールは常に、スタンプではなく封筒で表されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="44d19-117">概念開発時にコアメタファを使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="44d19-118">要素の削減</span><span class="sxs-lookup"><span data-stu-id="44d19-118">Reduction of Elements</span></span>

- <span data-ttu-id="44d19-119">アイコンは、メタファに不可欠な要素のみを使用して、中心となる意味を小さくします。</span><span class="sxs-lookup"><span data-stu-id="44d19-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="44d19-120">アイコンのサイズに関係なく、アイコンの要素の数を2に制限します。</span><span class="sxs-lookup"><span data-stu-id="44d19-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="44d19-121">一貫</span><span class="sxs-lookup"><span data-stu-id="44d19-121">Consistency</span></span>

<span data-ttu-id="44d19-122">アイコンのサイズ、配置、色は一貫している必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="44d19-123">スタイル</span><span class="sxs-lookup"><span data-stu-id="44d19-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="44d19-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="44d19-124">Perspective</span></span>

<span data-ttu-id="44d19-125">既定では、Monoline アイコンは前方に向きます。</span><span class="sxs-lookup"><span data-stu-id="44d19-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="44d19-126">キューブなどの遠近や回転を必要とする一部の要素は許可されますが、例外は最小限に抑える必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="44d19-127">装飾記号</span><span class="sxs-lookup"><span data-stu-id="44d19-127">Embellishment</span></span>

<span data-ttu-id="44d19-128">Monoline は、完全な最小のスタイルです。</span><span class="sxs-lookup"><span data-stu-id="44d19-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="44d19-129">すべてがフラットな色を使用しているため、グラデーション、テクスチャ、または光源がないことを意味します。</span><span class="sxs-lookup"><span data-stu-id="44d19-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="44d19-130">設計</span><span class="sxs-lookup"><span data-stu-id="44d19-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="44d19-131">フェース</span><span class="sxs-lookup"><span data-stu-id="44d19-131">Sizes</span></span>

<span data-ttu-id="44d19-132">高 DPI デバイスをサポートするには、これらのサイズで各アイコンを生成することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="44d19-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="44d19-133">絶対に *必要な* サイズは、100% のサイズである 16 px、20 px、および 32 px です。</span><span class="sxs-lookup"><span data-stu-id="44d19-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="44d19-134">**16 px、20 px、24ピクセル、32 px、40 px、48 px、64 px、80 px、96 px**</span><span class="sxs-lookup"><span data-stu-id="44d19-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

### <a name="layout"></a><span data-ttu-id="44d19-135">レイアウト</span><span class="sxs-lookup"><span data-stu-id="44d19-135">Layout</span></span>

<span data-ttu-id="44d19-136">次に、修飾子付きのアイコンレイアウトの例を示します。</span><span class="sxs-lookup"><span data-stu-id="44d19-136">The following is an example of icon layout with a modifier.</span></span>

![修飾子が右下のアイコンの図](../images/monolineicon1.png)  ![ベース、修飾子、スペース、およびカットアウトのグリッドの背景と吹き出しが追加された同じアイコンの図](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="44d19-139">要素</span><span class="sxs-lookup"><span data-stu-id="44d19-139">Elements</span></span>

- <span data-ttu-id="44d19-140">**Base**: アイコンが表す主な概念です。</span><span class="sxs-lookup"><span data-stu-id="44d19-140">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="44d19-141">これは通常、アイコンに必要なビジュアルだけですが、第2の要素 (修飾子) を使用して主な概念を拡張することもできます。</span><span class="sxs-lookup"><span data-stu-id="44d19-141">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="44d19-142">**修飾子** ベースをオーバーレイする任意の要素。これは、通常、アクションまたはステータスを表す修飾子です。</span><span class="sxs-lookup"><span data-stu-id="44d19-142">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="44d19-143">追加、変更、または記述子として機能することによって、基本要素を変更します。</span><span class="sxs-lookup"><span data-stu-id="44d19-143">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![ベースおよび修飾子が "out" と呼ばれるグリッドの図](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="44d19-145">建設</span><span class="sxs-lookup"><span data-stu-id="44d19-145">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="44d19-146">要素の配置</span><span class="sxs-lookup"><span data-stu-id="44d19-146">Element placement</span></span>

<span data-ttu-id="44d19-147">Base 要素は、スペース内のアイコンの中央に配置されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-147">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="44d19-148">中心を完全に配置できない場合は、上から右にエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-148">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="44d19-149">次の例では、アイコンは完全に中央揃えになっています。</span><span class="sxs-lookup"><span data-stu-id="44d19-149">In the following example, the icon is perfectly centered.</span></span>

![完全に中央揃えアイコンが表示されている図](../images/monolineicon4.png)

<span data-ttu-id="44d19-151">次の例では、アイコンは左に erring ます。</span><span class="sxs-lookup"><span data-stu-id="44d19-151">In the following example, the icon is erring to the left.</span></span>

![Errs の左に1ピクセルのアイコンを示す図](../images/monolineicon5.png)

<span data-ttu-id="44d19-153">修飾子は、ほとんどの場合、アイコンキャンバスの右下隅に配置されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-153">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="44d19-154">まれに、修飾子が異なる隅に配置される場合があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-154">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="44d19-155">たとえば、底要素が右下隅の修飾子で認識されない場合は、左上隅に配置することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="44d19-155">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![右下に修飾子が表示された4つのアイコンと、右上にモディファイアがあるアイコンが1つある図](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="44d19-157">Padding</span><span class="sxs-lookup"><span data-stu-id="44d19-157">Padding</span></span>

<span data-ttu-id="44d19-158">各サイズアイコンには、アイコンの周囲に指定された余白があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-158">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="44d19-159">Base 要素は埋め込みの範囲内に残りますが、補助線はキャンバスの端までの辺までになり、アイコンの境界線の外側まで拡張されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-159">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="44d19-160">次の画像は、アイコンのサイズごとに推奨される埋め込みを示しています。</span><span class="sxs-lookup"><span data-stu-id="44d19-160">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="44d19-161">**16px**</span><span class="sxs-lookup"><span data-stu-id="44d19-161">**16px**</span></span>|<span data-ttu-id="44d19-162">**20px**</span><span class="sxs-lookup"><span data-stu-id="44d19-162">**20px**</span></span>|<span data-ttu-id="44d19-163">**24px**</span><span class="sxs-lookup"><span data-stu-id="44d19-163">**24px**</span></span>|<span data-ttu-id="44d19-164">**32px**</span><span class="sxs-lookup"><span data-stu-id="44d19-164">**32px**</span></span>|<span data-ttu-id="44d19-165">**40px**</span><span class="sxs-lookup"><span data-stu-id="44d19-165">**40px**</span></span>|<span data-ttu-id="44d19-166">**48px**</span><span class="sxs-lookup"><span data-stu-id="44d19-166">**48px**</span></span>|<span data-ttu-id="44d19-167">**64px**</span><span class="sxs-lookup"><span data-stu-id="44d19-167">**64px**</span></span>|<span data-ttu-id="44d19-168">**80px**</span><span class="sxs-lookup"><span data-stu-id="44d19-168">**80px**</span></span>|<span data-ttu-id="44d19-169">**96px**</span><span class="sxs-lookup"><span data-stu-id="44d19-169">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![0px のパディング付きの 16 px アイコン](../images/monolineicon7.png)|![1px padding 付きの 20 px アイコン](../images/monolineicon8.png)|![1px padding を含む 24 px アイコン](../images/monolineicon9.png)|![32 px、2px のパディング付きのアイコン](../images/monolineicon10.png)|![40 px、2px のパディング付きのアイコン](../images/monolineicon11.png)|![48 px アイコンと3ピクセルのパディング](../images/monolineicon12.png)|![64ピクセル、4ピクセルのパディング付きのアイコン](../images/monolineicon13.png)|![80ピクセル、5ピクセルのパディング付きのアイコン](../images/monolineicon14.png)|![96 px、および6px のパディング付きのアイコン](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="44d19-179">線の太さ</span><span class="sxs-lookup"><span data-stu-id="44d19-179">Line weights</span></span>

<span data-ttu-id="44d19-180">Monoline は、線とアウトライン付きの図形のスタイルです。</span><span class="sxs-lookup"><span data-stu-id="44d19-180">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="44d19-181">アイコンのサイズに応じて、次の線の太さを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-181">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="44d19-182">アイコンのサイズ:</span><span class="sxs-lookup"><span data-stu-id="44d19-182">Icon Size:</span></span>|<span data-ttu-id="44d19-183">16px</span><span class="sxs-lookup"><span data-stu-id="44d19-183">16px</span></span>|<span data-ttu-id="44d19-184">20px</span><span class="sxs-lookup"><span data-stu-id="44d19-184">20px</span></span>|<span data-ttu-id="44d19-185">24px</span><span class="sxs-lookup"><span data-stu-id="44d19-185">24px</span></span>|<span data-ttu-id="44d19-186">32px</span><span class="sxs-lookup"><span data-stu-id="44d19-186">32px</span></span>|<span data-ttu-id="44d19-187">40px</span><span class="sxs-lookup"><span data-stu-id="44d19-187">40px</span></span>|<span data-ttu-id="44d19-188">48px</span><span class="sxs-lookup"><span data-stu-id="44d19-188">48px</span></span>|<span data-ttu-id="44d19-189">64px</span><span class="sxs-lookup"><span data-stu-id="44d19-189">64px</span></span>|<span data-ttu-id="44d19-190">80px</span><span class="sxs-lookup"><span data-stu-id="44d19-190">80px</span></span>|<span data-ttu-id="44d19-191">96px</span><span class="sxs-lookup"><span data-stu-id="44d19-191">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="44d19-192">**線の太さ:**</span><span class="sxs-lookup"><span data-stu-id="44d19-192">**Line Weight:**</span></span>|<span data-ttu-id="44d19-193">1px</span><span class="sxs-lookup"><span data-stu-id="44d19-193">1px</span></span>|<span data-ttu-id="44d19-194">1px</span><span class="sxs-lookup"><span data-stu-id="44d19-194">1px</span></span>|<span data-ttu-id="44d19-195">1px</span><span class="sxs-lookup"><span data-stu-id="44d19-195">1px</span></span>|<span data-ttu-id="44d19-196">1px</span><span class="sxs-lookup"><span data-stu-id="44d19-196">1px</span></span>|<span data-ttu-id="44d19-197">2px</span><span class="sxs-lookup"><span data-stu-id="44d19-197">2px</span></span>|<span data-ttu-id="44d19-198">2px</span><span class="sxs-lookup"><span data-stu-id="44d19-198">2px</span></span>|<span data-ttu-id="44d19-199">2px</span><span class="sxs-lookup"><span data-stu-id="44d19-199">2px</span></span>|<span data-ttu-id="44d19-200">2px</span><span class="sxs-lookup"><span data-stu-id="44d19-200">2px</span></span>|<span data-ttu-id="44d19-201">3px</span><span class="sxs-lookup"><span data-stu-id="44d19-201">3px</span></span>|
|<span data-ttu-id="44d19-202">**アイコンの例:**</span><span class="sxs-lookup"><span data-stu-id="44d19-202">**Example icon:**</span></span>|![16ピクセルアイコン](../images/monolineicon16.png)|![20ピクセルアイコン](../images/monolineicon17.png)|![24 px アイコン](../images/monolineicon18.png)|![32 px アイコン](../images/monolineicon19.png)|![40 px アイコン](../images/monolineicon20.png)|![48 px アイコン](../images/monolineicon21.png)|![64 px アイコン](../images/monolineicon22.png)|![80 px アイコン](../images/monolineicon23.png)|![96 px アイコン](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="44d19-212">切り抜き</span><span class="sxs-lookup"><span data-stu-id="44d19-212">Cutouts</span></span>

<span data-ttu-id="44d19-213">Icon 要素が別の要素の上に配置されている場合は、主に読みやすくするために、(bottom 要素の) 切り抜きを使用して、2つの要素の間にスペースを提供します。</span><span class="sxs-lookup"><span data-stu-id="44d19-213">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="44d19-214">これは通常、修飾子が基本要素の上に配置されている場合に、いずれの要素も修飾子ではない場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="44d19-214">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="44d19-215">これらの2つの要素間の切り抜きは、「gap」と呼ばれることもあります。</span><span class="sxs-lookup"><span data-stu-id="44d19-215">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="44d19-216">間隔のサイズは、そのサイズに対して使用される線の太さと同じである必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-216">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="44d19-217">16ピクセルのアイコンを作成する場合は、間隔を1px に設定し、それが48ピクセルのアイコンの場合、間隔を2px にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-217">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="44d19-218">次の例では、1px と基になるベースとの間にギャップがある 32 px アイコンを示しています。</span><span class="sxs-lookup"><span data-stu-id="44d19-218">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![1px と基になる基準の間にギャップがある 32 px アイコン](../images/monolineicon25.png)

<span data-ttu-id="44d19-220">場合によっては、境界線が斜めまたは曲線の場合、間隔を1/2 ピクセルに増やすことができる場合があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-220">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="44d19-221">この場合は、1px 線の太さが 16 px、20 px、24ピクセル、および 32 px のアイコンのみに影響する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-221">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="44d19-222">背景の塗りつぶし</span><span class="sxs-lookup"><span data-stu-id="44d19-222">Background fills</span></span>

<span data-ttu-id="44d19-223">Monoline アイコンセットのほとんどのアイコンは、背景の塗りつぶしを必要とします。</span><span class="sxs-lookup"><span data-stu-id="44d19-223">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="44d19-224">ただし、オブジェクトが自然に塗りつぶされていない場合は、塗りつぶしを適用する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="44d19-224">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="44d19-225">次のアイコンには白の塗りつぶしが設定されています。</span><span class="sxs-lookup"><span data-stu-id="44d19-225">The following icons have a white fill.</span></span>

![白塗りつぶしが設定される5つのアイコンのコンパイル](../images/monolineicon26.png)

<span data-ttu-id="44d19-227">次のアイコンには塗りつぶしがありません。</span><span class="sxs-lookup"><span data-stu-id="44d19-227">The following icons have no fill.</span></span> <span data-ttu-id="44d19-228">(中央の穴が塗りつぶされていないことを示す歯車アイコンが含まれています)。</span><span class="sxs-lookup"><span data-stu-id="44d19-228">(The gear icon is included to show that the center hole is not filled.)</span></span>

![塗りつぶしなしの5つのアイコンのコンパイル](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="44d19-230">塗りつぶしのベストプラクティス</span><span class="sxs-lookup"><span data-stu-id="44d19-230">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="44d19-231">Do</span><span class="sxs-lookup"><span data-stu-id="44d19-231">Dos:</span></span>

- <span data-ttu-id="44d19-232">境界が定義されている任意の要素を塗りつぶします。塗りつぶしがあります。</span><span class="sxs-lookup"><span data-stu-id="44d19-232">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="44d19-233">背景の塗りつぶしを作成するには、別の図形を使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-233">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="44d19-234">[カラーパレット](#color)から **背景の塗りつぶし** を使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-234">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="44d19-235">重なり合う要素間のピクセルの間隔を維持します。</span><span class="sxs-lookup"><span data-stu-id="44d19-235">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="44d19-236">複数のオブジェクト間での塗りつぶし。</span><span class="sxs-lookup"><span data-stu-id="44d19-236">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="44d19-237">注意事項</span><span class="sxs-lookup"><span data-stu-id="44d19-237">Don'ts:</span></span>

- <span data-ttu-id="44d19-238">自然に入力されていないオブジェクトを塗りつぶすことはできません。たとえば、クリップになります。</span><span class="sxs-lookup"><span data-stu-id="44d19-238">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="44d19-239">角かっこを入力しません。</span><span class="sxs-lookup"><span data-stu-id="44d19-239">Don't fill brackets.</span></span>
- <span data-ttu-id="44d19-240">数字または英字の後ろには入力しないでください。</span><span class="sxs-lookup"><span data-stu-id="44d19-240">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="44d19-241">色</span><span class="sxs-lookup"><span data-stu-id="44d19-241">Color</span></span>

<span data-ttu-id="44d19-242">カラーパレットは、シンプルかつアクセシビリティを目的として設計されています。</span><span class="sxs-lookup"><span data-stu-id="44d19-242">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="44d19-243">この中には、4つの中間色の色と、青、緑、黄色、赤、および紫の2つのバリエーションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="44d19-243">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="44d19-244">オレンジ色は、意図的に Monoline アイコンカラーパレットに含まれていません。</span><span class="sxs-lookup"><span data-stu-id="44d19-244">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="44d19-245">各色は、このセクションで説明する特定の方法で使用することを目的としています。</span><span class="sxs-lookup"><span data-stu-id="44d19-245">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="44d19-246">カラー</span><span class="sxs-lookup"><span data-stu-id="44d19-246">Palette</span></span>

![Monoline の4つの灰色の網掛け: スタンドアロンまたはアウトライン用の濃い灰色、アウトラインまたはコンテンツ用の淡い灰色、背景の塗りつぶしに対して明るい灰色、および塗りつぶしに薄い灰色](../images/monoline-grayshades.png)

![Monoline のカラーパレットには、スタンドアロン、アウトライン、および塗りつぶし用の青、緑、黄色、赤、紫の影が含まれています。](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="44d19-249">色の使用方法</span><span class="sxs-lookup"><span data-stu-id="44d19-249">How to use color</span></span>

<span data-ttu-id="44d19-250">Monoline カラーパレットでは、すべての色に、スタンドアロン、アウトライン、および塗りつぶしのバリエーションがあります。</span><span class="sxs-lookup"><span data-stu-id="44d19-250">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="44d19-251">通常、要素は塗りつぶしと輪郭で構成されています。</span><span class="sxs-lookup"><span data-stu-id="44d19-251">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="44d19-252">色は、次のいずれかのパターンで適用されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-252">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="44d19-253">塗りつぶしが設定されていないオブジェクトのスタンドアロンカラーのみ。</span><span class="sxs-lookup"><span data-stu-id="44d19-253">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="44d19-254">罫線はアウトラインの色を使用し、塗りつぶしには塗りつぶしの色が使用されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-254">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="44d19-255">罫線は、スタンドアロンの色を使用し、塗りつぶしには背景の塗りつぶしの色が使用されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-255">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="44d19-256">次に色を使用する例を示します。</span><span class="sxs-lookup"><span data-stu-id="44d19-256">The following are examples of using color.</span></span>

![境界線、塗りつぶし、またはその両方に色を設定した3つのアイコンのコンパイル](../images/monolineicon28.png)

<span data-ttu-id="44d19-258">最も一般的な状況は、要素に背景の塗りつぶしを使用して濃い灰色のスタンドアロンを使用することです。</span><span class="sxs-lookup"><span data-stu-id="44d19-258">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="44d19-259">色の塗りつぶしを使用する場合は、常に、対応する輪郭の色で表示される必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-259">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="44d19-260">たとえば、青い塗りつぶしは青い輪郭線でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-260">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="44d19-261">ただし、この一般的なルールには次の2つの例外があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-261">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="44d19-262">背景の塗りつぶしは、任意の色のスタンドアロンで使用できます。</span><span class="sxs-lookup"><span data-stu-id="44d19-262">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="44d19-263">明るい灰色の塗りつぶしは、2つの異なる輪郭の色 (濃い灰色または淡い灰色) で使用できます。</span><span class="sxs-lookup"><span data-stu-id="44d19-263">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="44d19-264">色を使用する場合</span><span class="sxs-lookup"><span data-stu-id="44d19-264">When to use color</span></span>

<span data-ttu-id="44d19-265">装飾記号ではなく、アイコンの意味を伝えるために色を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-265">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="44d19-266">ユーザーに対する **アクションを強調表示** する必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-266">It should **highlight the action** to the user.</span></span> <span data-ttu-id="44d19-267">色が設定された基本要素に修飾子を追加すると、通常は、次のセットの左端のアイコンにある "X" 修飾子が picture base に追加されている場合など、基本要素が暗い灰色および背景の塗りつぶしに変更されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-267">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![色を使用する5つのアイコンのコンパイル](../images/monolineicon29.png)

<span data-ttu-id="44d19-269">アイコンは、前に説明したアウトラインと塗りつぶし以外の **1 つ** の追加色に制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-269">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="44d19-270">ただし、メタファにとって重要であり、灰色以外の2つの追加の色の制限がある場合は、より多くの色を使用できます。</span><span class="sxs-lookup"><span data-stu-id="44d19-270">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="44d19-271">その他の色を必要とする場合、例外が発生することがまれにあります。</span><span class="sxs-lookup"><span data-stu-id="44d19-271">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="44d19-272">次に示すのは、1つの色のみを使用するアイコンの適切な例です。</span><span class="sxs-lookup"><span data-stu-id="44d19-272">The following are good examples of icons that use just one color.</span></span>

  ![それぞれが1つの色を使用する5つのアイコンのコンパイル](../images/monolineicon30.png)

<span data-ttu-id="44d19-274">しかし、次のアイコンは多くの色を使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-274">But the following icons use too many colors.</span></span>

  ![それぞれが複数の色を使用する5つのアイコンのコンパイル](../images/monolineicon31.png)

<span data-ttu-id="44d19-276">内部の "コンテンツ" (スプレッドシートのアイコン内のグリッド線など) に **淡い灰色** を使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-276">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="44d19-277">コンテンツにコントロールの動作を表示する必要がある場合は、追加の内部色が使用されます。</span><span class="sxs-lookup"><span data-stu-id="44d19-277">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![淡い灰色の内部要素を含む5つのアイコンのコンパイル](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="44d19-279">テキスト行</span><span class="sxs-lookup"><span data-stu-id="44d19-279">Text lines</span></span>

<span data-ttu-id="44d19-280">テキスト行が "container" (たとえば、ドキュメントのテキスト) にある場合は、淡い灰色を使用します。</span><span class="sxs-lookup"><span data-stu-id="44d19-280">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="44d19-281">コンテナー内にないテキスト行は、 **濃い灰色** にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-281">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="44d19-282">テキスト</span><span class="sxs-lookup"><span data-stu-id="44d19-282">Text</span></span>

<span data-ttu-id="44d19-283">アイコンでテキスト文字を使用することは避けてください。</span><span class="sxs-lookup"><span data-stu-id="44d19-283">Avoid using text characters in icons.</span></span> <span data-ttu-id="44d19-284">Office 製品は世界中で使用されているため、アイコンをできるだけニュートラルにしておきたいと考えています。</span><span class="sxs-lookup"><span data-stu-id="44d19-284">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="44d19-285">生産</span><span class="sxs-lookup"><span data-stu-id="44d19-285">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="44d19-286">アイコンファイルの形式</span><span class="sxs-lookup"><span data-stu-id="44d19-286">Icon file format</span></span>

<span data-ttu-id="44d19-287">最終的なアイコンは .png 画像ファイルとして保存する必要があります。</span><span class="sxs-lookup"><span data-stu-id="44d19-287">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="44d19-288">透明な背景で PNG 形式を使用し、32ビットの深さを設定します。</span><span class="sxs-lookup"><span data-stu-id="44d19-288">Use PNG format with a transparent background and have 32-bit depth.</span></span>
