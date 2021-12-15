---
title: PowerPoint JavaScript プレビュー API
description: JavaScript API のPowerPoint詳細。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 406808b4b4ff16df72d9c37468696525c8be642f
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513992"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint JavaScript プレビュー API

JavaScript api PowerPoint最初に "プレビュー" で導入され、後で十分なテストが行われるとユーザーフィードバックが取得された後、特定の番号付き要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| スライドの管理 | スライドの追加とスライド レイアウトとスライド マスターの管理のサポートを追加します。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| 図形 | スライド内の図形への参照を取得するサポートを追加します。 | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中PowerPoint JavaScript API の一覧を示します。 すべての JavaScript API (プレビュー API PowerPoint以前にリリースされた API を含む) の完全な一覧については[、「JavaScript API Excel参照してください](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|段落内の箇条書きが表示される場合に指定します。|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|段落の箇条書き形式を表します。|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|段落の水平方向の配置を表します。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|この図形の塗りつぶしの書式設定を返します。|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|図形の高さをポイントで指定します。|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|図形の左側からスライドの左側までの距離をポイントで指定します。|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|この図形の線の書式設定を返します。|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|この図形の名前を指定します。|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|この図形のテキスト フレーム オブジェクトを返します。|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|図形の上端からスライドの上端までの距離をポイントで指定します。|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|この図形の種類を返します。|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|図形の幅をポイント単位で指定します。|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|図形の高さをポイントで指定します。|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|図形の左側からスライドの左側までの距離をポイントで指定します。|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|図形の上端からスライドの上端までの距離をポイントで指定します。|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|図形の幅をポイント単位で指定します。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType、オプション?: PowerPoint。ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|スライドに幾何学的な図形を追加します。|
||[addLine(connectorType?: PowerPoint.ConnectorType、オプション?: PowerPoint。ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|スライドに線を追加します。|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|指定したテキストをコンテンツとしてスライドにテキスト ボックスを追加します。|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|図形塗りつぶしの前景色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500"など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|図形の塗りつぶしの書式設定を均一な色に設定します。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|塗りつぶしの透明度の割合を 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|図形の塗りつぶしの種類を返します。|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|テキストの色の HTML カラー コード表現 (例: "#FF0000" は赤を表します)。|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|フォント名 ("Calibri" など) を表します。|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|フォント サイズをポイント (11 など) で表します。|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|フォントに適用する下線の種類。|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|線の色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500" など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|線のダッシュ スタイルを表します。|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|図形の線スタイルを表します。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|行の透明度のパーセンテージを 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|図形要素の線の書式設定が表示される場合に指定します。|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|線の太さ (ポイント数) を表します。|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|テキスト フレームの自動サイズ設定。|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|テキスト フレーム内のテキストをすべて削除します。|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|テキスト フレームにテキストが含まれている場合に指定します。|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|テキスト フレームの垂直方向の配置を表します。|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|図形内のテキストに合わせて線が自動的に折れるかどうかを指定します。|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|テキスト範囲の `ShapeFont` フォント属性を表すオブジェクトを返します。|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|指定した範囲 `TextRange` の部分文字列のオブジェクトを返します。|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|テキスト範囲の段落形式を表します。|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|

## <a name="see-also"></a>関連項目

- [PowerPoint JavaScript API リファレンス ドキュメント](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API の要件セット](powerpoint-api-requirement-sets.md)