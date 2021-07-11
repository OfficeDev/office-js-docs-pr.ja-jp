---
title: アドインのモノOfficeアイコンのガイドライン
description: アドインで Monoline スタイル アイコンを使用Officeガイドライン。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: a59574f1f49ccb8b7b6fd485d08f83e39d760a48
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349344"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="b070c-103">アドインのモノOfficeアイコンのガイドライン</span><span class="sxs-lookup"><span data-stu-id="b070c-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="b070c-104">モノライン スタイルのアイコンは、アプリでOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="b070c-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="b070c-105">2013+ のサブスクリプション以外の新しいスタイルとアイコンが一致することを希望する場合は、「Office アドインのフレッシュ スタイル[アイコンガイドライン](add-in-icons-fresh.md)」を参照Officeしてください。</span><span class="sxs-lookup"><span data-stu-id="b070c-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="b070c-106">Officeモノラインの表示スタイル</span><span class="sxs-lookup"><span data-stu-id="b070c-106">Office Monoline visual style</span></span>

<span data-ttu-id="b070c-107">モノライン スタイルの目的は、アクションと機能を簡単なビジュアルで通信し、アイコンがすべてのユーザーにアクセス可能で、Windows の他の場所で使用されているスタイルと一致するスタイルを持つ、一貫性のある、明確でアクセスしやすい図ノグラフィを持つという目的です。</span><span class="sxs-lookup"><span data-stu-id="b070c-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="b070c-108">次のガイドラインは、既に製品に表示されているアイコンと一致する機能のアイコンを作成するサードパーティの開発者Officeです。</span><span class="sxs-lookup"><span data-stu-id="b070c-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="b070c-109">デザインの原則</span><span class="sxs-lookup"><span data-stu-id="b070c-109">Design principles</span></span>

- <span data-ttu-id="b070c-110">シンプル、クリーン、クリア。</span><span class="sxs-lookup"><span data-stu-id="b070c-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="b070c-111">必要な要素のみを含む。</span><span class="sxs-lookup"><span data-stu-id="b070c-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="b070c-112">アイコンのスタイルWindowsにインスパイアされています。</span><span class="sxs-lookup"><span data-stu-id="b070c-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="b070c-113">すべてのユーザーがアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="b070c-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="b070c-114">意味を伝える</span><span class="sxs-lookup"><span data-stu-id="b070c-114">Conveying meaning</span></span>

- <span data-ttu-id="b070c-115">ページなどの説明的な要素を使用して、文書や封筒を表してメールを表します。</span><span class="sxs-lookup"><span data-stu-id="b070c-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="b070c-116">同じ要素を使用して同じ概念を表します。つまり、メールは常にスタンプではなく封筒で表されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="b070c-117">概念開発中にコアメタファーを使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="b070c-118">要素の削減</span><span class="sxs-lookup"><span data-stu-id="b070c-118">Reduction of Elements</span></span>

- <span data-ttu-id="b070c-119">比喩に不可欠な要素のみを使用して、アイコンを中心の意味に減らします。</span><span class="sxs-lookup"><span data-stu-id="b070c-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="b070c-120">アイコンのサイズに関係なく、アイコン内の要素の数を 2 に制限します。</span><span class="sxs-lookup"><span data-stu-id="b070c-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="b070c-121">整合性</span><span class="sxs-lookup"><span data-stu-id="b070c-121">Consistency</span></span>

<span data-ttu-id="b070c-122">アイコンのサイズ、配置、色は一貫している必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="b070c-123">スタイル設定</span><span class="sxs-lookup"><span data-stu-id="b070c-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="b070c-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="b070c-124">Perspective</span></span>

<span data-ttu-id="b070c-125">モノライン アイコンは既定で前方に向かっています。</span><span class="sxs-lookup"><span data-stu-id="b070c-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="b070c-126">キューブなどの視点や回転を必要とする特定の要素は許可されますが、例外は最小限に抑える必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="b070c-127">装飾</span><span class="sxs-lookup"><span data-stu-id="b070c-127">Embellishment</span></span>

<span data-ttu-id="b070c-128">モノラインはクリーンな最小限のスタイルです。</span><span class="sxs-lookup"><span data-stu-id="b070c-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="b070c-129">すべてはフラットカラーを使用します。つまり、グラデーション、テクスチャ、光源はありません。</span><span class="sxs-lookup"><span data-stu-id="b070c-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="b070c-130">設計</span><span class="sxs-lookup"><span data-stu-id="b070c-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="b070c-131">サイズ</span><span class="sxs-lookup"><span data-stu-id="b070c-131">Sizes</span></span>

<span data-ttu-id="b070c-132">高 DPI デバイスをサポートするために、これらのすべてのサイズで各アイコンを作成することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b070c-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="b070c-133">絶対に必要 *なサイズ* は 16 px、20 px、および 32 px で、サイズは 100% です。</span><span class="sxs-lookup"><span data-stu-id="b070c-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="b070c-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span><span class="sxs-lookup"><span data-stu-id="b070c-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b070c-135">アドインの代表的なアイコンである画像については、「サイズなどの要件については[、「AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)および Office 内で効果的なリストを作成する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b070c-135">For an image that is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

### <a name="layout"></a><span data-ttu-id="b070c-136">レイアウト</span><span class="sxs-lookup"><span data-stu-id="b070c-136">Layout</span></span>

<span data-ttu-id="b070c-137">修飾子付きアイコン レイアウトの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b070c-137">The following is an example of icon layout with a modifier.</span></span>

![右下に修飾子が付いたアイコンの図。](../images/monolineicon1.png)  ![ベース、修飾子、パディング、およびカットアウトにグリッドの背景と吹き出しが追加された同じアイコンの図。](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="b070c-140">要素</span><span class="sxs-lookup"><span data-stu-id="b070c-140">Elements</span></span>

- <span data-ttu-id="b070c-141">**基本**: アイコンが表す主な概念。</span><span class="sxs-lookup"><span data-stu-id="b070c-141">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="b070c-142">これは通常、アイコンに必要な唯一のビジュアルですが、セカンダリ要素である修飾子を使用して主な概念を拡張できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-142">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="b070c-143">**修飾子** ベースをオーバーレイする要素。つまり、通常はアクションまたは状態を表す修飾子です。</span><span class="sxs-lookup"><span data-stu-id="b070c-143">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="b070c-144">追加、変更、または記述子として機能することで、基本要素を変更します。</span><span class="sxs-lookup"><span data-stu-id="b070c-144">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![呼び出された基本領域と修飾子領域を持つグリッドの図。](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="b070c-146">建設</span><span class="sxs-lookup"><span data-stu-id="b070c-146">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="b070c-147">要素の配置</span><span class="sxs-lookup"><span data-stu-id="b070c-147">Element placement</span></span>

<span data-ttu-id="b070c-148">基本要素は、パディング内のアイコンの中央に配置されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-148">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="b070c-149">完全に中央に配置できない場合は、基部の位置が一番上の位置に誤りがある必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-149">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="b070c-150">次の例では、アイコンは完全に中央に表示されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-150">In the following example, the icon is perfectly centered.</span></span>

![完全に中央に位置するアイコンを示す図。](../images/monolineicon4.png)

<span data-ttu-id="b070c-152">次の例では、アイコンが左側にエラーが発生しています。</span><span class="sxs-lookup"><span data-stu-id="b070c-152">In the following example, the icon is erring to the left.</span></span>

![左に 1 px のエラーが表示されるアイコンを示す図。](../images/monolineicon5.png)

<span data-ttu-id="b070c-154">修飾子は、ほとんどの場合、アイコン キャンバスの右下隅に配置されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-154">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="b070c-155">まれに、修飾子が別のコーナーに配置される場合があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-155">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="b070c-156">たとえば、右下隅の修飾子で基本要素を認識できない場合は、左上隅に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-156">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![右下に修飾子を持つ 4 つのアイコンと、左上に修飾子を持つ 1 つのアイコンを示す図。](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="b070c-158">Padding</span><span class="sxs-lookup"><span data-stu-id="b070c-158">Padding</span></span>

<span data-ttu-id="b070c-159">各サイズ アイコンには、アイコンの周囲に指定した量のパディングがあります。</span><span class="sxs-lookup"><span data-stu-id="b070c-159">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="b070c-160">基本要素はパディング内に残りますが、修飾子はキャンバスの端まで突き合わせ、パディングの外側をアイコンの境界線の端まで拡張する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-160">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="b070c-161">次の図は、各アイコン サイズに使用する推奨されるパディングを示しています。</span><span class="sxs-lookup"><span data-stu-id="b070c-161">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="b070c-162">**16px**</span><span class="sxs-lookup"><span data-stu-id="b070c-162">**16px**</span></span>|<span data-ttu-id="b070c-163">**20px**</span><span class="sxs-lookup"><span data-stu-id="b070c-163">**20px**</span></span>|<span data-ttu-id="b070c-164">**24px**</span><span class="sxs-lookup"><span data-stu-id="b070c-164">**24px**</span></span>|<span data-ttu-id="b070c-165">**32px**</span><span class="sxs-lookup"><span data-stu-id="b070c-165">**32px**</span></span>|<span data-ttu-id="b070c-166">**40px**</span><span class="sxs-lookup"><span data-stu-id="b070c-166">**40px**</span></span>|<span data-ttu-id="b070c-167">**48px**</span><span class="sxs-lookup"><span data-stu-id="b070c-167">**48px**</span></span>|<span data-ttu-id="b070c-168">**64px**</span><span class="sxs-lookup"><span data-stu-id="b070c-168">**64px**</span></span>|<span data-ttu-id="b070c-169">**80px**</span><span class="sxs-lookup"><span data-stu-id="b070c-169">**80px**</span></span>|<span data-ttu-id="b070c-170">**96px**</span><span class="sxs-lookup"><span data-stu-id="b070c-170">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![0px パディングを含む 16 px アイコン。](../images/monolineicon7.png)|![1px パディングを含む 20 px アイコン。](../images/monolineicon8.png)|![1px パディングを含む 24 px アイコン。](../images/monolineicon9.png)|![2px パディングを含む 32 px アイコン。](../images/monolineicon10.png)|![2px パディングを含む 40 px アイコン。](../images/monolineicon11.png)|![3px パディングを含む 48 px アイコン。](../images/monolineicon12.png)|![4px パディングを含む 64 px アイコン。](../images/monolineicon13.png)|![5px パディングを含む 80 px アイコン。](../images/monolineicon14.png)|![6px パディング付き 96 px アイコン。](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="b070c-180">線の太さ</span><span class="sxs-lookup"><span data-stu-id="b070c-180">Line weights</span></span>

<span data-ttu-id="b070c-181">モノラインは、線とアウトラインの図形で支配されるスタイルです。</span><span class="sxs-lookup"><span data-stu-id="b070c-181">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="b070c-182">アイコンを作成するサイズに応じて、次の線の太みを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-182">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="b070c-183">アイコンのサイズ:</span><span class="sxs-lookup"><span data-stu-id="b070c-183">Icon Size:</span></span>|<span data-ttu-id="b070c-184">16px</span><span class="sxs-lookup"><span data-stu-id="b070c-184">16px</span></span>|<span data-ttu-id="b070c-185">20px</span><span class="sxs-lookup"><span data-stu-id="b070c-185">20px</span></span>|<span data-ttu-id="b070c-186">24px</span><span class="sxs-lookup"><span data-stu-id="b070c-186">24px</span></span>|<span data-ttu-id="b070c-187">32px</span><span class="sxs-lookup"><span data-stu-id="b070c-187">32px</span></span>|<span data-ttu-id="b070c-188">40px</span><span class="sxs-lookup"><span data-stu-id="b070c-188">40px</span></span>|<span data-ttu-id="b070c-189">48px</span><span class="sxs-lookup"><span data-stu-id="b070c-189">48px</span></span>|<span data-ttu-id="b070c-190">64px</span><span class="sxs-lookup"><span data-stu-id="b070c-190">64px</span></span>|<span data-ttu-id="b070c-191">80px</span><span class="sxs-lookup"><span data-stu-id="b070c-191">80px</span></span>|<span data-ttu-id="b070c-192">96px</span><span class="sxs-lookup"><span data-stu-id="b070c-192">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="b070c-193">**線の太さ:**</span><span class="sxs-lookup"><span data-stu-id="b070c-193">**Line Weight:**</span></span>|<span data-ttu-id="b070c-194">1px</span><span class="sxs-lookup"><span data-stu-id="b070c-194">1px</span></span>|<span data-ttu-id="b070c-195">1px</span><span class="sxs-lookup"><span data-stu-id="b070c-195">1px</span></span>|<span data-ttu-id="b070c-196">1px</span><span class="sxs-lookup"><span data-stu-id="b070c-196">1px</span></span>|<span data-ttu-id="b070c-197">1px</span><span class="sxs-lookup"><span data-stu-id="b070c-197">1px</span></span>|<span data-ttu-id="b070c-198">2px</span><span class="sxs-lookup"><span data-stu-id="b070c-198">2px</span></span>|<span data-ttu-id="b070c-199">2px</span><span class="sxs-lookup"><span data-stu-id="b070c-199">2px</span></span>|<span data-ttu-id="b070c-200">2px</span><span class="sxs-lookup"><span data-stu-id="b070c-200">2px</span></span>|<span data-ttu-id="b070c-201">2px</span><span class="sxs-lookup"><span data-stu-id="b070c-201">2px</span></span>|<span data-ttu-id="b070c-202">3px</span><span class="sxs-lookup"><span data-stu-id="b070c-202">3px</span></span>|
|<span data-ttu-id="b070c-203">**アイコンの例:**</span><span class="sxs-lookup"><span data-stu-id="b070c-203">**Example icon:**</span></span>|![16 px アイコン。](../images/monolineicon16.png)|![20 px アイコン。](../images/monolineicon17.png)|![24 px アイコン。](../images/monolineicon18.png)|![32 px アイコン。](../images/monolineicon19.png)|![40 px アイコン。](../images/monolineicon20.png)|![48 px アイコン。](../images/monolineicon21.png)|![64 px アイコン。](../images/monolineicon22.png)|![80 px アイコン。](../images/monolineicon23.png)|![96 px アイコン。](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="b070c-213">カットアウト</span><span class="sxs-lookup"><span data-stu-id="b070c-213">Cutouts</span></span>

<span data-ttu-id="b070c-214">icon 要素を別の要素の上に配置すると、主に読みやすさを目的として、2 つの要素の間にスペースを提供するために(下の要素の) 切り抜きが使用されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-214">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="b070c-215">これは通常、修飾子が基本要素の上に配置されている場合に発生しますが、どちらの要素も修飾子でもない場合があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-215">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="b070c-216">2 つの要素間のこれらの切り抜きは、"ギャップ" と呼ばれる場合があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-216">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="b070c-217">ギャップのサイズは、そのサイズで使用される線の太さと同じ幅である必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-217">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="b070c-218">16 px アイコンを作成する場合、ギャップ幅は 1px で、48 px アイコンの場合、ギャップは 2px になります。</span><span class="sxs-lookup"><span data-stu-id="b070c-218">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="b070c-219">次の使用例は、修飾子と基になるベースの間に 1px の間隔を持つ 32 px アイコンを示しています。</span><span class="sxs-lookup"><span data-stu-id="b070c-219">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![修飾子と基になるベースの間に 1px の間隔を持つ 32 px アイコン。](../images/monolineicon25.png)

<span data-ttu-id="b070c-221">場合によっては、修飾子に対角線または曲線のエッジが存在し、標準のギャップが十分に分離できない場合は、ギャップを 1/2 ピクセル増やします。</span><span class="sxs-lookup"><span data-stu-id="b070c-221">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="b070c-222">これは、線の太さ 16 px、20 px、24 px、および 32 px のアイコンにのみ影響します。</span><span class="sxs-lookup"><span data-stu-id="b070c-222">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="b070c-223">背景の塗りつぶし</span><span class="sxs-lookup"><span data-stu-id="b070c-223">Background fills</span></span>

<span data-ttu-id="b070c-224">モノライン アイコン セットのほとんどのアイコンでは、背景の塗りつぶしが必要です。</span><span class="sxs-lookup"><span data-stu-id="b070c-224">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="b070c-225">ただし、オブジェクトが自然に塗りつぶしを持たないので、塗りつぶしを適用しない場合があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-225">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="b070c-226">次のアイコンは、白い塗りつぶしを持っています。</span><span class="sxs-lookup"><span data-stu-id="b070c-226">The following icons have a white fill.</span></span>

![白塗りつぶしの 5 つのアイコンのコンパイル。](../images/monolineicon26.png)

<span data-ttu-id="b070c-228">次のアイコンには塗りつぶしはありません。</span><span class="sxs-lookup"><span data-stu-id="b070c-228">The following icons have no fill.</span></span> <span data-ttu-id="b070c-229">(中央の穴が塗りつぶされていないのを示す歯車アイコンが含まれています)。</span><span class="sxs-lookup"><span data-stu-id="b070c-229">(The gear icon is included to show that the center hole is not filled.)</span></span>

![塗りつぶしがない 5 つのアイコンのコンパイル。](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="b070c-231">塗りつぶしのベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="b070c-231">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="b070c-232">Dos:</span><span class="sxs-lookup"><span data-stu-id="b070c-232">Dos:</span></span>

- <span data-ttu-id="b070c-233">定義された境界を持ち、自然に塗りつぶしを持つ要素を塗りつぶします。</span><span class="sxs-lookup"><span data-stu-id="b070c-233">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="b070c-234">背景塗りつぶしを作成するには、別の図形を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-234">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="b070c-235">カラー \*\*パレットの [背景の\*\* 塗りつぶし [] を使用します](#color)。</span><span class="sxs-lookup"><span data-stu-id="b070c-235">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="b070c-236">重なり合う要素間のピクセルの分離を維持します。</span><span class="sxs-lookup"><span data-stu-id="b070c-236">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="b070c-237">複数のオブジェクト間で塗りつぶしを行います。</span><span class="sxs-lookup"><span data-stu-id="b070c-237">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="b070c-238">T'ts:</span><span class="sxs-lookup"><span data-stu-id="b070c-238">Don'ts:</span></span>

- <span data-ttu-id="b070c-239">自然に塗りつぶされないオブジェクトを塗りつぶす必要があります。たとえば、ペーパークリップなどです。</span><span class="sxs-lookup"><span data-stu-id="b070c-239">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="b070c-240">角かっこを埋めない。</span><span class="sxs-lookup"><span data-stu-id="b070c-240">Don't fill brackets.</span></span>
- <span data-ttu-id="b070c-241">数字やアルファ文字の後ろに埋め込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-241">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="b070c-242">色</span><span class="sxs-lookup"><span data-stu-id="b070c-242">Color</span></span>

<span data-ttu-id="b070c-243">カラー パレットは、わかりやすく、アクセシビリティを考慮して設計されています。</span><span class="sxs-lookup"><span data-stu-id="b070c-243">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="b070c-244">青、緑、黄、赤、紫の 4 色と 2 種類のバリエーションが含まれる。</span><span class="sxs-lookup"><span data-stu-id="b070c-244">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="b070c-245">オレンジ色は、モノライン アイコンのカラー パレットに意図的に含まれません。</span><span class="sxs-lookup"><span data-stu-id="b070c-245">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="b070c-246">各色は、このセクションで説明されている特定の方法で使用することを意図しています。</span><span class="sxs-lookup"><span data-stu-id="b070c-246">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="b070c-247">パレット</span><span class="sxs-lookup"><span data-stu-id="b070c-247">Palette</span></span>

![モノラインの灰色の 4 つの網掛け: スタンドアロンまたはアウトラインの濃い灰色、アウトラインまたはコンテンツの中灰色、背景塗りつぶしの非常に薄い灰色、塗りつぶしの淡い灰色。](../images/monoline-grayshades.png)

![モノラインのカラー パレットには、スタンドアロン、アウトライン、塗りつぶしの青、緑、黄色、赤、紫の色合いが含まれます。](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="b070c-250">色の使い方</span><span class="sxs-lookup"><span data-stu-id="b070c-250">How to use color</span></span>

<span data-ttu-id="b070c-251">モノライン カラー パレットでは、すべての色にスタンドアロン、アウトライン、塗りつぶしのバリエーションがあります。</span><span class="sxs-lookup"><span data-stu-id="b070c-251">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="b070c-252">一般に、要素は塗りつぶしと罫線で構成されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-252">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="b070c-253">色は、次のいずれかのパターンで適用されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-253">The colors are applied in one of the following patterns.</span></span>

- <span data-ttu-id="b070c-254">塗りつぶしがないオブジェクトのスタンドアロン色。</span><span class="sxs-lookup"><span data-stu-id="b070c-254">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="b070c-255">罫線はアウトライン色を使用し、塗りつぶしは塗りつぶしの色を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-255">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="b070c-256">枠線はスタンドアロン色を使用し、塗りつぶしは背景塗りつぶしの色を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-256">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="b070c-257">色の使用例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b070c-257">The following are examples of using color.</span></span>

![境界線または塗りつぶしまたは両方の色を持つ 3 つのアイコンのコンパイル。](../images/monolineicon28.png)

<span data-ttu-id="b070c-259">最も一般的な状況は、要素が背景塗りつぶしを持つ濃い灰色のスタンドアロンを使用する場合です。</span><span class="sxs-lookup"><span data-stu-id="b070c-259">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="b070c-260">塗りつぶしを使用する場合は、常に対応するアウトラインの色を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-260">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="b070c-261">たとえば、青い塗りつぶしは青いアウトラインでのみ使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-261">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="b070c-262">ただし、この一般的なルールには 2 つの例外があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-262">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="b070c-263">背景の塗りつぶしは、任意の色スタンドアロンで使用できます。</span><span class="sxs-lookup"><span data-stu-id="b070c-263">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="b070c-264">薄い灰色の塗りつぶしは、2 つの異なるアウトライン色 (濃い灰色または中灰色) で使用できます。</span><span class="sxs-lookup"><span data-stu-id="b070c-264">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="b070c-265">色を使用する場合</span><span class="sxs-lookup"><span data-stu-id="b070c-265">When to use color</span></span>

<span data-ttu-id="b070c-266">装飾ではなく、アイコンの意味を伝えるために色を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-266">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="b070c-267">ユーザーに **対するアクションを強調表示** する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-267">It should **highlight the action** to the user.</span></span> <span data-ttu-id="b070c-268">色を持つ基本要素に修飾子を追加すると、通常、基本要素は濃い灰色と背景塗りつぶしに変換され、次のセットの左端のアイコンの図のベースに "X" 修飾子が追加されている場合など、修飾子を色の要素にできます。</span><span class="sxs-lookup"><span data-stu-id="b070c-268">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![色を使用する 5 つのアイコンのコンパイル。](../images/monolineicon29.png)

<span data-ttu-id="b070c-270">上記のアウトラインと塗りつぶし以外 **の 1** つの追加の色にアイコンを制限する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-270">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="b070c-271">ただし、比喩に不可欠な場合は、灰色以外の 2 つの追加の色を制限して、より多くの色を使用できます。</span><span class="sxs-lookup"><span data-stu-id="b070c-271">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="b070c-272">まれに、より多くの色が必要な場合は例外があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-272">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="b070c-273">1 つの色を使用するアイコンの良い例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b070c-273">The following are good examples of icons that use just one color.</span></span>

  ![それぞれ 1 色を使用する 5 つのアイコンのコンパイル。](../images/monolineicon30.png)

<span data-ttu-id="b070c-275">ただし、次のアイコンでは色が多すぎます。</span><span class="sxs-lookup"><span data-stu-id="b070c-275">But the following icons use too many colors.</span></span>

  ![それぞれ複数の色を使用する 5 つのアイコンのコンパイル。](../images/monolineicon31.png)

<span data-ttu-id="b070c-277">スプレッドシート **のアイコンの** グリッド線など、内部の "コンテンツ" には中灰色を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-277">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="b070c-278">追加の内部色は、コンテンツがコントロールの動作を表示する必要がある場合に使用されます。</span><span class="sxs-lookup"><span data-stu-id="b070c-278">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![中灰色の内部要素を持つ 5 つのアイコンのコンパイル。](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="b070c-280">テキスト行</span><span class="sxs-lookup"><span data-stu-id="b070c-280">Text lines</span></span>

<span data-ttu-id="b070c-281">テキスト行が "コンテナー" (ドキュメント上のテキストなど) にある場合は、中灰色を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-281">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="b070c-282">コンテナー内に含めないテキスト行は濃い **灰色である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="b070c-282">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="b070c-283">テキスト</span><span class="sxs-lookup"><span data-stu-id="b070c-283">Text</span></span>

<span data-ttu-id="b070c-284">アイコンでテキスト文字を使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="b070c-284">Avoid using text characters in icons.</span></span> <span data-ttu-id="b070c-285">世界中Office製品が使用されていますので、アイコンは可能な限り言語に依存しない状態にしておきたいと考えています。</span><span class="sxs-lookup"><span data-stu-id="b070c-285">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="b070c-286">生産</span><span class="sxs-lookup"><span data-stu-id="b070c-286">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="b070c-287">アイコン ファイル形式</span><span class="sxs-lookup"><span data-stu-id="b070c-287">Icon file format</span></span>

<span data-ttu-id="b070c-288">最後のアイコンは、イメージ ファイルとして.pngする必要があります。</span><span class="sxs-lookup"><span data-stu-id="b070c-288">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="b070c-289">背景が透明で、奥行きが 32 ビットの PNG 形式を使用します。</span><span class="sxs-lookup"><span data-stu-id="b070c-289">Use PNG format with a transparent background and have 32-bit depth.</span></span>

## <a name="see-also"></a><span data-ttu-id="b070c-290">関連項目</span><span class="sxs-lookup"><span data-stu-id="b070c-290">See also</span></span>

- [<span data-ttu-id="b070c-291">Icon manifest 要素</span><span class="sxs-lookup"><span data-stu-id="b070c-291">Icon manifest element</span></span>](../reference/manifest/icon.md)
- [<span data-ttu-id="b070c-292">IconUrl マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="b070c-292">IconUrl manifest element</span></span>](../reference/manifest/iconurl.md)
- [<span data-ttu-id="b070c-293">HighResolutionIconUrl マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="b070c-293">HighResolutionIconUrl manifest element</span></span>](../reference/manifest/highresolutioniconurl.md)
- [<span data-ttu-id="b070c-294">アドイン用のアイコンの作成</span><span class="sxs-lookup"><span data-stu-id="b070c-294">Create an icon for your add-in</span></span>](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
