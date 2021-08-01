---
title: ExcelJavaScript API 要件セット 1.9
description: ExcelApi 1.9 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 41f6eb2dd329a2ab82981cb3ee8e11a784e23591
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671885"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>JavaScript API 1.9 Excel新機能

1.9 の要件セットにより、500 件を超える新しい Excel API が 導入されました。 最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | 画像、幾何学な図形、テキスト ボックスを挿入、位置変更、書式設定します。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape) [Image](/javascript/api/excel/excel.image) |
| [オート フィルター](../../excel/excel-add-ins-worksheets.md#filter-data) | 範囲にフィルターを追加します。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [エリア](../../excel/excel-add-ins-multiple-ranges.md) | 連続していない範囲をサポートします。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特別なセル](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | ある範囲内に日付、コメント、数式を含むセルを取得します。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [検索](../../excel/excel-add-ins-ranges-string-match.md) | ある範囲またはワークシート内で値や数式を見つけます。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [コピーと貼り付け](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | 範囲間で値、書式、数式をコピーします。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [計算](../../excel/performance.md#suspend-calculation-temporarily) | Excel 計算エンジンを細かく操作できます。 | [アプリケーション](/javascript/api/excel/excel.application) |
| 新しいグラフ | 新しくサポートされたグラフである、マップ、箱ひげ図、ウォーターフォール、サンバースト、パレート、 じょうごをお試しください。 | [Chart](/javascript/api/excel/excel.charttype) |
| 範囲の形式 | 範囲の形式の新しい機能です。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.9 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.9 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|アプリケーションの計算の状態を返します。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|反復計算の設定を返します。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|次の呼び出しが呼び出されるまで、画面の `context.sync()` 更新を中断します。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|範囲にオートフィルターを適用します。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|オートフィルターのフィルター条件がクリアされます。|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|`Range`オートフィルターを適用する範囲を表すオブジェクトを返します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|`Range`オートフィルターを適用する範囲を表すオブジェクトを返します。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|オートフィルターが有効になっている場合に指定します。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|オートフィルターにフィルター条件がある場合に指定します。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|その範囲で現在指定されている Autofilter オブジェクトを適用します。|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|範囲の AutoFilter を削除します。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|1 つの境界線の `color` プロパティを表します。|
||[style](/javascript/api/excel/excel.cellborder#style)|1 つの境界線の `style` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|1 つの境界線の `tintAndShade` プロパティを表します。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|1 つの境界線の `weight` プロパティを表します。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|`format.borders.bottom` プロパティを表します。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|`format.borders.diagonalDown` プロパティを表します。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|`format.borders.diagonalUp` プロパティを表します。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|`format.borders.horizontal` プロパティを表します。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|`format.borders.left` プロパティを表します。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|`format.borders.right` プロパティを表します。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|`format.borders.top` プロパティを表します。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|`format.borders.vertical` プロパティを表します。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|`addressLocal` プロパティを表します。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|`hidden` プロパティを表します。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|`format.fill.color` プロパティを表します。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|`format.fill.pattern` プロパティを表します。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|`format.fill.patternColor` プロパティを表します。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|`format.fill.patternTintAndShade` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|`format.fill.tintAndShade` プロパティを表します。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|`format.font.bold` プロパティを表します。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|`format.font.color` プロパティを表します。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|`format.font.italic` プロパティを表します。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|`format.font.name` プロパティを表します。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|`format.font.size` プロパティを表します。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|`format.font.strikethrough` プロパティを表します。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|`format.font.subscript` プロパティを表します。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|`format.font.superscript` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|`format.font.tintAndShade` プロパティを表します。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|`format.font.underline` プロパティを表します。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|`autoIndent` プロパティを表します。|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|`borders` プロパティを表します。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|`fill` プロパティを表します。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|`font` プロパティを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|`horizontalAlignment` プロパティを表します。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|`indentLevel` プロパティを表します。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|`protection` プロパティを表します。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|`readingOrder` プロパティを表します。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|`shrinkToFit` プロパティを表します。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|`textOrientation` プロパティを表します。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|`useStandardHeight` プロパティを表します。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|`useStandardWidth` プロパティを表します。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|`verticalAlignment` プロパティを表します。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|`wrapText` プロパティを表します。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|`format.protection.formulaHidden` プロパティを表します。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|`format.protection.locked` プロパティを表します。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|変更後の値を表します。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|変更前の値を表します。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|変更後の値の種類を表します。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|変更前の値の種類を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|Excel UI でグラフをアクティブにします。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|ピボット グラフのオプションをカプセル化します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|グラフの配色を指定します。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|グラフのグラフ領域の角が丸い場合に指定します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|数値の形式がセルにリンクされている場合に指定します。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|ヒストグラム グラフまたはパレート グラフでビン オーバーフローが有効になっている場合に指定します。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|ヒストグラム グラフまたはパレート グラフでビンアンダーフローが有効になっている場合に指定します。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|ヒストグラム グラフまたはパレート グラフのビン数を指定します。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|ヒストグラム グラフまたはパレート グラフのビン オーバーフロー値を指定します。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|ヒストグラム グラフまたはパレート グラフのビンの種類を指定します。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|ヒストグラム グラフまたはパレート グラフのビンアンダーフロー値を指定します。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|ヒストグラム グラフまたはパレート グラフのビン幅の値を指定します。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|ボックスグラフとひげグラフの四分位計算の種類を指定します。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|ボックスとひげグラフに内側の点を表示する場合に指定します。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|平均線をボックスとひげグラフに表示する場合に指定します。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|平均マーカーをボックスとひげグラフに表示する場合に指定します。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|ボックスとひげグラフに外れ値ポイントを表示する場合に指定します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|セルに番号の書式をリンクする (セル内でラベルが変更された場合に数値の書式が変更される) 場合に指定します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|数値の形式がセルにリンクされている場合に指定します。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|エラー バーに終了スタイル の上限が設定されている場合に指定します。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|誤差範囲のどの部分を含めるかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|誤差範囲の書式の種類を指定します。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|エラー バーを表示するかどうかを指定します。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|グラフの線の書式設定を表します。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|地域マップ グラフの系列マップ ラベル戦略を指定します。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|地域マップ グラフの系列マッピング レベルを指定します。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|地域マップ グラフの系列投影の種類を指定します。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|軸フィールド ボタンをウィンドウに表示するかどうかを指定ピボットグラフ。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|凡例フィールド ボタンを凡例フィールド ボタンで表示するかどうかを指定ピボットグラフ。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|レポート にレポート フィルター フィールド ボタンを表示するかどうかを指定ピボットグラフ。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|フィールドの [値の表示] ボタンを表示するかどうかを指定ピボットグラフ。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|地域マップ グラフ系列の最大値の色を指定します。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|地域マップ グラフ系列の最大値の種類を指定します。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|地域マップ グラフ系列の最大値を指定します。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|地域マップ グラフ系列の中点の値の色を指定します。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|地域マップ グラフ系列の中点値の種類を指定します。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|地域マップ グラフ系列の中点の値を指定します。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|地域マップ グラフ系列の最小値の色を指定します。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|地域マップ グラフ系列の最小値の種類を指定します。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|地域マップ グラフ系列の最小値を指定します。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|地域マップ グラフの系列グラデーション スタイルを指定します。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|系列内の負のデータ ポイントの塗りつぶしの色を指定します。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|ツリーマップ グラフの系列の親ラベル戦略領域を指定します。|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|箱ひげ図グラフのオプションをカプセル化します。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|リージョン マップ グラフのオプションをカプセル化します。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|ウォーターフォール グラフにコネクタ線を表示するかどうかを指定します。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|系列内のデータ ラベルごとに引き出し線を表示するかどうかを指定します。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|円グラフまたは棒グラフの 2 つのセクションを分割するしきい値を指定します。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|セルに番号の書式をリンクする (セル内でラベルが変更された場合に数値の書式が変更される) 場合に指定します。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|`addressLocal` プロパティを表します。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|`columnIndex` プロパティを表します。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|1 つ以上の四角形の範囲で構成され、その範囲に対 `RangeAreas` して同じ形式が適用される値を返します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|無効なセル値を持つ 1 つ以上の四角形の範囲を含む `RangeAreas` オブジェクトを返します。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|無効なセル値を持つ 1 つ以上の四角形の範囲を含む `RangeAreas` オブジェクトを返します。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|豊富な値に対してリッチ フィルターを実行するためにフィルターで使用されるプロパティ。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|図形 ID を返します。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|幾何学的図形の `Shape` オブジェクトを返します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|図形グループの図形の数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|コレクション内の位置に基づいて図形を取得します。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|ワークシートの中央フッター。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|ワークシートの中央ヘッダー。|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|ワークシートの左側のフッター。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|ワークシートの左側のヘッダー。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|ワークシートの右側のフッター。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|ワークシートの右側のヘッダー。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|ヘッダー/フッターが設定されている状態。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|画像の形式を返します。|
||[id](/javascript/api/excel/excel.image#id)|イメージ オブジェクトの図形識別子を指定します。|
||[shape](/javascript/api/excel/excel.image#shape)|イメージに関連 `Shape` 付けられているオブジェクトを返します。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|循環参照を解決するために、各反復間のExcelを指定します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|循環参照の解決に使用Excel繰り返しの最大数を指定します。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|指定された線の始点の矢印の幅を表します。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|指定されたコネクタの始点を指定された図形に接続します。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|指定されたコネクタの終点を指定された図形に接続します。|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|線のコネクタの種類を表します。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|指定されたコネクタの始点を図形から切り離します。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|指定されたコネクタの終点を図形から切り離します。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|指定された線の終点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|指定された線の始点が接続されている図形を表します。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|コネクタの始点が接続されている結合点を表します。|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|指定された線の終点が接続されている図形を表します。|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|コネクタの終点が接続されている結合点を表します。|
||[id](/javascript/api/excel/excel.line#id)|図形識別子を指定します。|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|指定した線の先頭が図形に接続される場合に指定します。|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|指定した線の端が図形に接続される場合に指定します。|
||[shape](/javascript/api/excel/excel.line#shape)|行に関連 `Shape` 付けられたオブジェクトを返します。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|改ページ オブジェクトを削除します。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|改ページの後の最初のセルを取得します。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|ページブレークの列インデックスを指定します。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|ページブレークの行インデックスを指定します。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|指定された範囲の左上セルの前に改ページを追加します。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|コレクション内の改ページの数を取得します。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|インデックス経由で改ページ オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|コレクション内の手動改ページをすべてリセットします。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|ワークシートの白黒印刷オプション。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|ポイントでの印刷に使用するワークシートの下部ページ余白。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|ワークシートの中央に水平方向にフラグを設定します。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|ワークシートの中央に垂直フラグを設定します。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|ワークシートの下書きモード オプション。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|印刷するワークシートの最初のページ番号。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|印刷時に使用するワークシートのフッター余白をポイントで指定します。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|ワークシートの印刷領域を表す 1 つ以上の四角形の範囲を含む `RangeAreas` オブジェクトを取得します。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|ワークシートの印刷領域を表す 1 つ以上の四角形の範囲を含む `RangeAreas` オブジェクトを取得します。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|タイトル行を表す範囲オブジェクトを取得します。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|タイトル行を表す範囲オブジェクトを取得します。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|印刷時に使用するワークシートのヘッダー余白をポイントで指定します。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|印刷時に使用するワークシートの左余白をポイントで指定します。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|ワークシートのページの向き。|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|ワークシートのページの用紙サイズ。|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|印刷時にワークシートのコメントを表示する必要がある場合に指定します。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|ワークシートの印刷エラー オプション。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|ワークシートの枠線を印刷する場合に指定します。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|ワークシートの見出しを印刷する場合に指定します。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|ワークシートのページ印刷順序オプション。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|ワークシートのヘッダーとフッターの構成。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|印刷時に使用するワークシートの右余白をポイントで指定します。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|ワークシートの印刷範囲を設定します。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|ワークシートのページ余白を単位で設定します。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|セルを含む列を、印刷時、ワークシートの各ページの左で繰り返すように設定します。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|セルを含む行を、印刷時、ワークシートの各ページの上で繰り返すように設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|印刷時に使用するワークシートの上余白をポイントで指定します。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|ワークシートの印刷ズーム オプション。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|印刷に使用する単位でページ レイアウトの下部余白を指定します。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|印刷に使用する単位のページ レイアウト フッター余白を指定します。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|印刷に使用する単位のページ レイアウト ヘッダー余白を指定します。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|印刷に使用する単位のページ レイアウト左余白を指定します。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|印刷に使用する単位のページ レイアウト右余白を指定します。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|印刷に使用する単位でページ レイアウトの上余白を指定します。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|横方向に合わせるページ数。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|印刷ページのスケール値は 10 から 400 までです。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|縦方向に合わせるページ数。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|書式設定が更新時またはフィールドの移動時に自動的に書式設定される場合を指定します。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|PivotTable 内で指定された範囲の値を計算するために使用される DataHierarchy を取得します。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算される場合に書式設定を保持する場合に指定します。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|必要なすべての条件とコンテキストを自動的に選択するため、指定したセルを使用して自動的に並べ替えを実行するようピボットテーブルを設定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|ピボットテーブルでデータ本文の値をユーザーが編集できる場合を指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|ピボットテーブルが並べ替え時にカスタム リストを使用する場合に指定します。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|指定した AutoFill ロジックを使用して、現在の範囲から移動先の範囲の範囲を塗りつぶしします。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|データ型を持つ範囲セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|ワークシート内の範囲セルをリンクされたデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|セル データまたは書式をソース範囲または現在の `RangeAreas` 範囲にコピーします。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|指定された条件に基づいて指定された文字列を見つけます。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|指定された条件に基づいて指定された文字列を見つけます。|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|現在の範囲にフラッシュ塗りつぶしを実行します。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|2D 配列を返します。各セルのフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|一次元配列を返します。各列のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|一次元配列を返します。各行のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|指定した種類と値に一致するセルを表す、1 つ以上の四角形の範囲を含むオブジェクト `RangeAreas` を取得します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|指定した種類と値に一致するセルを表す 1 つ以上の範囲を含むオブジェクト `RangeAreas` を取得します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|範囲と重なるテーブルの集まりを範囲限定で取得します。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|各セルのデータ型の状態を表します。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|列によって指定される範囲から重複する値を削除します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|現在の範囲内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|セル プロパティの 2D 配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|列プロパティの 1 次元配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|次の再計算が発生したときに再計算する範囲を設定します。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|行プロパティの 1 次元配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|内のすべてのセルを計算します `RangeAreas` 。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|このオブジェクトを構成する各領域の値、書式、塗りつぶし、罫線、その他のプロパティをクリア `RangeAreas` します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|データ型を含むすべての `RangeAreas` セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|リンクされたデータ型に、すべての `RangeAreas` セルを変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|セル データまたは書式をソース範囲または現在の `RangeAreas` 範囲からコピーします `RangeAreas` 。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|列全体を表すオブジェクトを返します (たとえば、カレントがセル `RangeAreas` `RangeAreas` `RangeAreas` "B4:E11, H2" を表す場合は、列 `RangeAreas` "B:E, H:H" を表す a を返します)。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|行全体を表すオブジェクトを返します (たとえば、カレントがセル "B4:E11" を表す場合は、行 `RangeAreas` `RangeAreas` `RangeAreas` `RangeAreas` "4:11" を表す a を返します)。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|指定した範囲 `RangeAreas` の交差部分を表すオブジェクトを返します `RangeAreas` 。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|指定した範囲 `RangeAreas` の交差部分を表すオブジェクトを返します `RangeAreas` 。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|特定の行 `RangeAreas` と列のオフセットによってシフトされるオブジェクトを返します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|指定した型 `RangeAreas` と値に一致するすべてのセルを表すオブジェクトを返します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|指定した型 `RangeAreas` と値に一致するすべてのセルを表すオブジェクトを返します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|このオブジェクト内の任意の範囲と重なるテーブルのスコープ付きコレクションを返 `RangeAreas` します。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|オブジェクト内の個々の四角形範囲のすべての使用領域を含む `RangeAreas` 使用される領域を返 `RangeAreas` します。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|オブジェクト内の個々の四角形範囲のすべての使用領域を含む `RangeAreas` 使用される領域を返 `RangeAreas` します。|
||[address](/javascript/api/excel/excel.rangeareas#address)|`RangeAreas`A1 スタイルの参照を返します。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|ユーザー ロケール内 `RangeAreas` の参照を返します。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|このオブジェクトを構成する四角形の範囲の数を返 `RangeAreas` します。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|このオブジェクトを構成する四角形の範囲のコレクションを返 `RangeAreas` します。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|オブジェクト内のセルの数を返し、個々の四角形のすべての範囲のセル数 `RangeAreas` を合計します。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|このオブジェクト内のセルと交差する条件付き書式のコレクションを返 `RangeAreas` します。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|内のすべての範囲のデータ検証オブジェクトを返します `RangeAreas` 。|
||[format](/javascript/api/excel/excel.rangeareas#format)|オブジェクト内のすべての範囲のフォント、塗りつぶし、罫線、配置、その他のプロパティをカプセル化するオブジェクト `RangeFormat` を返 `RangeAreas` します。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|このオブジェクトのすべての範囲が列全体を表す `RangeAreas` ("A:C、Q:Z"など) を指定します。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|このオブジェクトのすべての範囲が行全体を表す (例: `RangeAreas` "1:3, 5:7") を指定します。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|現在のワークシートを返します `RangeAreas` 。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|次の `RangeAreas` 再計算が行われるときに再計算されるように設定します。|
||[style](/javascript/api/excel/excel.rangeareas#style)|このオブジェクトのすべての範囲のスタイルを表 `RangeAreas` します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|範囲の境界線の色を明るくまたは暗くする倍数を指定します。値は -1 (最も暗い) ~ 1 (最も明るい) で、元の色の場合は 0 です。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|範囲の境界線の色を明るくまたは暗くする倍数を指定します。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|内の範囲の数を返します `RangeCollection` 。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|内の位置に基づいて range オブジェクトを返します `RangeCollection` 。|
||[items](/javascript/api/excel/excel.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|範囲のパターン。|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|範囲パターンの色を表す HTML カラー コードは、#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("オレンジ色" など) として表されます。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|範囲塗りつぶしのパターンの色を明るくまたは暗くする倍数を指定します。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|範囲塗りつぶしの色を明るくまたは暗くする倍数を指定します。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|フォントの取り消し線の状態を指定します。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|フォントの下付き文字の状態を指定します。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|フォントの上付き文字の状態を指定します。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|範囲フォントの色を明るくまたは暗くする倍数を指定します。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|テキストの配置が等しい分布に設定されている場合に、テキストが自動的にインデントされる場合に指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|使用可能な列の幅に収まるテキストを自動的に縮小する場合に指定します。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|一致で大文字と小文字が区別される場合を指定します。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|`addressLocal` プロパティを表します。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|`rowIndex` プロパティを表します。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|一致で大文字と小文字が区別される場合を指定します。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|検索の方向を指定します。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|`format` プロパティを表します。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|`hyperlink` プロパティを表します。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|`style` プロパティを表します。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|`columnHidden` プロパティを表します。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[format: Excel.CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|`format` プロパティを表します。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)|`format` プロパティを表します。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|`rowHidden` プロパティを表します。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|オブジェクトの代替説明テキストを指定 `Shape` します。|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|オブジェクトの代替タイトル テキストを指定 `Shape` します。|
||[delete()](/javascript/api/excel/excel.shape#delete__)|ワークシートから図形を削除します。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|この幾何学的な図形の幾何学的な図形の種類を指定します。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。|
||[height](/javascript/api/excel/excel.shape#height)|図形の高さをポイントで指定します。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|指定したポイント数だけ水平方向に図形を移動します。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|z 軸を中心に、指定された度数だけ、図形を時計回りに回転します。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|指定したポイント数だけ垂直方向に図形を移動します。|
||[left](/javascript/api/excel/excel.shape#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|この図形の縦横比をロックする場合に指定します。|
||[name](/javascript/api/excel/excel.shape#name)|図形の名前を指定します。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|この図形の結合点の数を返します。|
||[fill](/javascript/api/excel/excel.shape#fill)|この図形の塗りつぶしの書式設定を返します。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|図形に関連付けられた幾何学的図形を返します。|
||[group](/javascript/api/excel/excel.shape#group)|図形に関連付けられた図形グループを返します。|
||[id](/javascript/api/excel/excel.shape#id)|図形識別子を指定します。|
||[image](/javascript/api/excel/excel.shape#image)|図形に関連付けられた画像を返します。|
||[level](/javascript/api/excel/excel.shape#level)|指定した図形のレベルを指定します。|
||[line](/javascript/api/excel/excel.shape#line)|図形に関連付けられた線を返します。|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|この図形の線の書式設定を返します。|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|図形がアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|図形が非アクティブになると発生します。|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|この図形の親グループを指定します。|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|この図形のテキスト フレーム オブジェクトを返します。|
||[type](/javascript/api/excel/excel.shape#type)|この図形の種類を返します。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|図形の回転角度を度で指定します。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|指定した係数分だけ図形の高さを変更します。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|指定した係数分だけ図形の幅を変更します。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[top](/javascript/api/excel/excel.shape#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[visible](/javascript/api/excel/excel.shape#visible)|図形が表示される場合に指定します。|
||[width](/javascript/api/excel/excel.shape#width)|図形の幅をポイント単位で指定します。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|図形をアクティブ化するワークシートの ID を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|幾何学的図形をワークシートに追加します。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|このコレクションのワークシート内の図形のサブセットをグループ化します。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|base64 エンコード文字列から画像を作成し、それをワークシートに追加します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|ワークシートに行を追加します。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|指定されたテキストを含むテキスト ボックスをワークシートに追加します。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|ワークシートの図形数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|コレクション内の位置を使用して図形を取得します。|
||[items](/javascript/api/excel/excel.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|非アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|図形が非アクティブ化されているワークシートの ID を取得します。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|図形塗りつぶしの前景色を #RRGGBB HTML の色形式で表します ("FFA500"など) 形式で、または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[type](/javascript/api/excel/excel.shapefill#type)|図形の塗りつぶしの種類を返します。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|図形の塗りつぶしの書式設定を均一な色に設定します。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|塗りつぶしの透明度の割合を 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.shapefont#color)|テキストの色の HTML カラー コード表現 (例: "#FF0000" は赤を表します)。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.shapefont#name)|フォント名 ("Calibri" など) を表します。|
||[size](/javascript/api/excel/excel.shapefont#size)|フォント サイズをポイント (11 など) で表します。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|フォントに適用する下線の種類。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|図形識別子を指定します。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|グループに関連 `Shape` 付けられているオブジェクトを返します。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|オブジェクトのコレクションを返 `Shape` します。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|指定した図形グループに含まれるグループ化された図形のグループを解除します。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|線の色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500" など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|図形の線スタイルを表します。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|図形の線スタイルを表します。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|図形要素の線の書式設定が表示される場合に指定します。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|線の太さ (ポイント数) を表します。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|並べ替えるリッチ値のターゲット プロパティ名であるサブフィールドを指定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|コレクション内のスタイルの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|コレクション内の位置に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|テーブルの `AutoFilter` オブジェクトを表します。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|追加されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|テーブルを追加するワークシートの ID を取得します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|変更の詳細に関する情報を取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|ブックに新しいテーブルが追加された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|指定されたテーブルがブックで削除されたときに発生します。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|削除されるテーブルの ID を取得します。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|削除されるテーブルの名前を取得します。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|テーブルが削除されるワークシートの ID を取得します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|コレクションに含まれるテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|コレクション内の最初のテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|名前または ID でテーブルを取得します。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|テキスト フレームの自動サイズ設定。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|テキスト フレーム内のテキストをすべて削除します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|テキスト フレームの水平方向の配置を表します。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|テキスト フレームの方向を指定する角度を表します。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|テキスト フレームの読む方向を表します (左から右または右から左)。|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|テキスト フレームにテキストが含まれている場合に指定します。|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|テキスト フレームの垂直方向の配置を表します。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|指定された範囲の部分文字列に対する TextRange オブジェクトを返します。|
||[font](/javascript/api/excel/excel.textrange#font)|テキスト範囲の `ShapeFont` フォント属性を表すオブジェクトを返します。|
||[text](/javascript/api/excel/excel.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|ブックで現在アクティブになっているグラフを取得します。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|ブックで現在アクティブになっているグラフを取得します。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|ブックが `true` 複数のユーザーによって編集されている場合 (共同編集による) 場合に返します。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|ブックから現在選択されている 1 つまたは複数の範囲を取得します。|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|ブックが最後に保存された後に変更が行われた場合に指定します。|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|ブックが自動保存モードの場合に指定します。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|Excel 計算エンジンのバージョンとして数字を返します。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|ブックで AutoSave 設定が変更された場合に発生します。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|ブックがローカルまたはオンラインで保存された場合に指定します。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|イベントの種類を取得します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|必要に応Excelワークシートを再計算する必要があるかどうかを決定します。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|指定された条件に基づいて、指定された文字列のすべての出現を検索し、1 つ以上の四角形の範囲で構成されるオブジェクト `RangeAreas` として返します。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|指定された条件に基づいて、指定された文字列のすべての出現を検索し、1 つ以上の四角形の範囲で構成されるオブジェクト `RangeAreas` として返します。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|アドレスまたは名前で指定された、四角形の範囲の 1 つ以上のブロックを表す `RangeAreas` オブジェクトを取得します。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|ワークシートの `AutoFilter` オブジェクトを表します。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|ワークシートの水平改ページをまとめて取得します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|フォーマットが特定のワークシートで変更されたときに発生します。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|ワークシートの `PageLayout` オブジェクトを取得します。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|ワークシート上のすべての Shape オブジェクトをまとめて返します。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|ワークシートの垂直改ページをまとめて取得します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|現在のワークシート内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|変更の詳細に関する情報を表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|ブックのワークシートが変更されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|ブック内のワークシートの形式が変更された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|ワークシートで選択範囲を変更したときに発生します。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|データが変更されたワークシートの ID を取得します。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|一致で大文字と小文字が区別される場合を指定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
