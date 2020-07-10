---
title: Office アドインの新しいスタイルのアイコンガイドライン
description: Office アドインでの新しいスタイルアイコンアイコンの使用に関するガイドラインを取得します。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 7f29de70712448e9ee7458db864fb40746412153
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093932"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Office アドインの新しいスタイルのアイコンガイドライン

Office 2013 + (サブスクリプション以外の) バージョンの Office では、Microsoft の新しいスタイル図像が使用されます。 アイコンが Microsoft 365 の Monoline スタイルに一致するようにする場合は、「 [Office アドインの Monoline スタイルアイコンのガイドライン](add-in-icons-monoline.md)」を参照してください。

## <a name="office-fresh-visual-style"></a>Office の新しい視覚スタイル

新しいアイコンには、重要な communicative 要素のみが含まれています。 遠近法、グラデーション、および光源など、重要でない要素が削除されています。 アイコンが簡略化されたことで、コマンドやコントロールの解析をより高速に行うことができるようになっています。 このスタイルは、Office 以外のサブスクリプションクライアントに最適なものにするために使用します。

## <a name="best-practices"></a>ベスト プラクティス

アイコンを作成するときは、以下のガイドラインに従ってください。

|するべきこと|してはいけないこと|
|:---|:---|
|コミュニケーションの主要な要素に重点を置いて、ビジュアルをシンプルかつ明瞭にします。| アイコンが乱雑に見える成果物は使用しないでください。|
|Office アイコンの言語を使用して、動作や概念を示します。|Office アプリのリボンまたはコンテキストメニューでは、アドインコマンドの Office UI Fabric グリフを転用しないようにしてください。 Fabric のアイコンはスタイルが異なるので、適合しません。|
|書式設定用のペイントブラシや検索用の虫眼鏡など、一般的な Office の視覚的メタファーを再利用します。|異なるコマンドで、同じ視覚的メタファーを再利用しないようにします。 異なる動作や概念に同じアイコンを使用すると、混乱が生じる可能性があります。 |
|アイコンを小さくしたり大きくしたりするために、アイコンを再描画します。 カットアウト、角、および丸角のエッジの線をできる限り明瞭にするために、再描画を行う手間を省かないでください。 |縮小または拡大してアイコンのサイズを変更しないでください。 これにより、視覚的品質が低くなり、動作が不明瞭になることがあります。 再描画せずにサイズを小さくすると、より大きなサイズで作成された複雑なアイコンから明瞭さが失われることがあります。 |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |アドイン コマンドで何をするかを伝えるために、ロゴやブランドに頼らないようにします。 ブランド マークは、サイズの小さいアイコンにしたり、修飾子を適用したりすると、しばしば認識不可能になります。 ブランドマークは、多くの場合、Office アプリのリボンアイコンのスタイルと競合し、飽和環境ではユーザーの注意を引くことができます。 |
|透明背景の PNG 形式を使用します。 ||
|アイコンに、表記文字、段落のラグ、および疑問符などの、ローカライズ可能なコンテンツを含めないようにします。 ||

## <a name="icon-size-recommendations-and-requirements"></a>アイコン サイズについて推奨事項と要件

Office のデスクトップ アイコンは、ビットマップ画像です。 ユーザーの DPI 設定やタッチ モードに応じて異なるサイズで表示されます。 サポートされている 8 つのサイズすべてを組み込んで、すべての解像度とコンテキストで最高のエクスペリエンスを提供します。 以下のサイズがサポートされています (うち 3 つは必須)：

- 16 px (必須)
- 20 px
- 24 px
- 32 px (必須)
- 40 px
- 48 px
- 64 ピクセル (推奨、Mac に最適)
- 80 px (必須)

それぞれのアイコンを、サイズに合わせて縮小するのではなく再描画します。

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

## <a name="icon-anatomy-and-layout"></a>アイコンの構造とレイアウト

Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

以下の図は、Office のアイコンの基本要素と修飾子のレイアウトを表しています。

![中央にアイコンの基本要素、右下に修飾子、左上にアクション修飾子を配した画像](../images/icon-layouts.png)

- 基本要素をピクセル フレームの中央に配置し、周囲に余白をとります。
- アクション修飾子は、左上に配置します。
- 概念的修飾子は、右下に配置します。
- Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.

### <a name="base-element-padding"></a>基本要素のパディング

基本要素は、どのサイズでも同じ配置にします。 基本要素をフレームの中央に配置できない場合は、左上にそろえ、余分のピクセルは右下に残します。 最良の結果を得るには、次のセクションの表に示すように、パディングのガイドラインを適用します。

### <a name="modifiers"></a>修飾子

All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.

|**アイコンのサイズ**|**基本要素の周囲のパディング**|**修飾子のサイズ**|
|:---|:---|:---|
|16px|.0|9px|
|20px|1px|10px|
|24px|1px|12px|
|32px|2px|14px|
|40px|2px|20px|
|48px|3px|22px|
|64px|5px|29px|
|80px|5px|38px|

## <a name="icon-colors"></a>アイコンの色

> [!NOTE]
> これらの色のガイドラインは、[アドイン コマンド](add-in-commands.md)で使用されるリボン アイコン用です。 これらのアイコンは Microsoft UI Fabric では表示されず、色のパレットは「[Microsoft UI Fabric | 色 | 共有](https://fluentfabric.azurewebsites.net/#/color/shared)」に記載されているパレットとは異なります。

Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color:

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark. 
- If possible, use only one additional color beyond gray. Limit additional colors to two at the most.
- Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.

|**色の名前**|**RGB**|**16 進数**|**色**|**分類**|
|:---|:---|:---|:---|:---|
|テキスト グレー (80)|80、80、80|#505050| ![テキスト グレー 80 のカラー イメージ](../images/color-text-gray-80.png) |テキスト|
|テキスト グレー (95)|95、95、95|#5F5F5F| ![テキスト グレー 95 のカラー イメージ](../images/color-text-gray-95.png) |テキスト|
|テキスト グレー (105)|105, 105, 105|#696969| ![テキスト グレー 105 のカラー イメージ](../images/color-text-gray-105.png) |テキスト|
|ダーク グレー 32|128、128、128|#808080| ![ダーク グレー 32 のカラー イメージ](../images/color-dark-gray-32.png) |32 以上|
|ミディアム グレー 32|158、158、158|#9E9E9E| ![ミディアム グレー 32 のカラー イメージ](../images/color-medium-gray-32.png) |32 以上|
|ライト グレー オール|179、179、179|#B3B3B3| ![ライト グレー オールのカラー イメージ](../images/color-light-gray-all.png) |すべてのサイズ|
|ダーク グレー 16|114、114、114|#727272| ![ダーク グレー 16 のカラー イメージ](../images/color-dark-gray-16.png) |16 以下|
|ミディアム グレー 16|144、144、144|#909090| ![ミディアム グレー 16 のカラー イメージ](../images/color-medium-gray-16.png) |16 以下|
|ブルー 32|77、130、184|#4d82B8| ![ブルー 32 のカラー イメージ](../images/color-blue-32.png) |32 以上|
|ブルー 16|74、125、177|#4A7DB1| ![ブルー 16 のカラー イメージ](../images/color-blue-16.png) |16 以下|
|イエロー オール|234、194、130|#EAC282| ![イエロー オールのカラー イメージ](../images/color-yellow-all.png) |すべてのサイズ|
|オレンジ 32|231、142、70|#E78E46| ![オレンジ 32 のカラー イメージ](../images/color-orange-32.png) |32 以上|
|オレンジ 16|227、142、70|#E3751C| ![オレンジ 16 のカラー イメージ](../images/color-orange-16.png) |16 以下|
|ピンク オール|230、132、151|#E68497| ![ピンク オールのカラー イメージ](../images/color-pink-all.png) |すべてのサイズ|
|グリーン 32|118、167、151|#76A797| ![グリーン 32 のカラー イメージ](../images/color-green-32.png) |32 以上|
|グリーン 16|104、164、144|#68A490| ![グリーン 16 のカラー イメージ](../images/color-green-16.png) |16 以下|
|レッド 32|216、99、68|#D86344| ![レッド 32 のカラー イメージ](../images/color-red-32.png) |32 以上|
|レッド 16|214、85、50|#D65532| ![レッド 16 のカラー イメージ](../images/color-red-16.png) |16 以下|
|パープル 32|152、104、185|#9868B9| ![パープル 32 のカラー イメージ](../images/color-purple-32.png) |32 以上|
|パープル 16|137、89、171|#8959AB| ![パープル 16 のカラー イメージ](../images/color-purple-16.png) |16 以下|

## <a name="icons-in-high-contrast-modes"></a>ハイコントラスト モードのアイコン

Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:

- 190 値のしきい値に沿って、前景と背景の要素を区別するようにします。
- Office アイコンの表示スタイルに従います。
- 色はアイコン パレットから使用します。
- グラデーションの使用を避けます。
- 同じ様な値を持つ大きな色のブロックを避けます。
