---
title: Excel JavaScript API 要件セット1.9
description: ExcelApi 1.9 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1d7e16a6e0aca202798016c136dfc7e2188c44f0
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940852"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Excel JavaScript API 1.9 の新機能

1.9 の要件セットにより、500 件を超える新しい Excel API が 導入されました。 最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | 画像、幾何学な図形、テキスト ボックスを挿入、位置変更、書式設定します。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape) [Image](/javascript/api/excel/excel.image) |
| [オート フィルター](../../excel/excel-add-ins-worksheets.md#filter-data) | 範囲にフィルターを追加します。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [エリア](../../excel/excel-add-ins-multiple-ranges.md) | 連続していない範囲をサポートします。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特別なセル](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | ある範囲内に日付、コメント、数式を含むセルを取得します。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [検索](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | ある範囲またはワークシート内で値や数式を見つけます。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [コピーと貼り付け](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | 範囲間で値、書式、数式をコピーします。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [計算](../../excel/performance.md#suspend-calculation-temporarily) | Excel 計算エンジンを細かく操作できます。 | [アプリケーション](/javascript/api/excel/excel.application) |
| 新しいグラフ | 新しくサポートされたグラフである、マップ、箱ひげ図、ウォーターフォール、サンバースト、パレート、 じょうごをお試しください。 | [Chart](/javascript/api/excel/excel.charttype) |
| 範囲の形式 | 範囲の形式の新しい機能です。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。 読み取り専用です。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|アプリケーションの計算の状態を返します。 詳細については、Excel.CalculationState をご覧ください。 読み取り専用です。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|反復計算の設定を返します。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|次の "context.sync()" が呼び出されるまで画面の更新を一時停止します。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|範囲にオートフィルターを適用します。 列インデックスやフィルター条件が指定されている場合、列にフィルターを適用します。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|オートフィルターのフィルター条件がクリアされます。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|AutoFilter が適用される範囲を表す Range オブジェクトを返します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|オートフィルターが適用される範囲を表す Range オブジェクトを返します。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。 読み取り専用です。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|AutoFilter が有効かどうかを示します。 読み取り専用です。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|AutoFilter にフィルター条件が与えられているかどうかを示します。 読み取り専用です。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|その範囲で現在指定されている Autofilter オブジェクトを適用します。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|範囲の AutoFilter を削除します。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|1 つの境界線の `color` プロパティを表します。|
||[style](/javascript/api/excel/excel.cellborder#style)|1 つの境界線の `style` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|1 つの境界線の `tintAndShade` プロパティを表します。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|1 つの境界線の `weight` プロパティを表します。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|`format.borders.bottom` プロパティを表します。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|`format.borders.diagonalDown` プロパティを表します。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|`format.borders.diagonalUp` プロパティを表します。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|`format.borders.horizontal` プロパティを表します。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|`format.borders.left` プロパティを表します。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|`format.borders.right` プロパティを表します。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|`format.borders.top` プロパティを表します。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|`format.borders.vertical` プロパティを表します。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|`hidden` プロパティを表します。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|`format.fill.color` プロパティを表します。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|`format.fill.pattern` プロパティを表します。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|`format.fill.patternColor` プロパティを表します。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|`format.fill.patternTintAndShade` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|`format.fill.tintAndShade` プロパティを表します。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|`format.font.bold` プロパティを表します。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|`format.font.color` プロパティを表します。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|`format.font.italic` プロパティを表します。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|`format.font.name` プロパティを表します。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|`format.font.size` プロパティを表します。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|`format.font.strikethrough` プロパティを表します。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|`format.font.subscript` プロパティを表します。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|`format.font.superscript` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|`format.font.tintAndShade` プロパティを表します。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|`format.font.underline` プロパティを表します。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|`autoIndent` プロパティを表します。|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|`borders` プロパティを表します。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|`fill` プロパティを表します。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|`font` プロパティを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|`horizontalAlignment` プロパティを表します。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|`indentLevel` プロパティを表します。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|`protection` プロパティを表します。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|`readingOrder` プロパティを表します。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|`shrinkToFit` プロパティを表します。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|`textOrientation` プロパティを表します。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|`useStandardHeight` プロパティを表します。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|`useStandardWidth` プロパティを表します。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|`verticalAlignment` プロパティを表します。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|`wrapText` プロパティを表します。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|`format.protection.formulaHidden` プロパティを表します。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|`format.protection.locked` プロパティを表します。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|変更後の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|変更前の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|変更後の値の型を表します。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|変更前の値の型を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Excel UI でグラフをアクティブにします。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|ピボット グラフのオプションをカプセル化します。 読み取り専用です。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|グラフの配色を返すか、設定します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|グラフのグラフ エリアの角が丸いかどうかを指定します。 読み取り/書き込み可能。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|箱ひげ図で、内部ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|箱ひげ図で、平均線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|箱ひげ図で、平均マーカーが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|箱ひげ図で、特異ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|誤差範囲に終点のスタイル線端があるかどうかを指定します。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|誤差範囲のどの部分を含めるかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|誤差範囲の書式の種類を指定します。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|誤差範囲が表示されるかどうかを指定します。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|グラフの線の書式設定を表します。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|リージョン マップ グラフの系列マッピング レベルを設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを指定します。 ShowAxisFieldButtons プロパティは、ピボットグラフを選択したときに有効になる [分析] タブの [フィールド ボタン] ボックスの一覧にある [軸フィールド ボタンを表示する] コマンドに対応します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを指定します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを指定します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|ピボットグラフに値フィールド ボタンを表示するかどうかを指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。 読み取り専用です。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。 読み取り専用です。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|リージョン マップ グラフのオプションをカプセル化します。 読み取り専用です。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|ウォーターフォール図で、コネクタが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|系列の各データ ラベルの引き出し線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|`columnIndex` プロパティを表します。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|1 つまたは複数の長方形範囲で構成され、条件付き書式が適用された RangeAreas を返します。 読み取り専用です。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは ItemNotFound エラーがスローされます。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは null が返されます。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|リッチな値にリッチなフィルターを適用する場合、フィルターによって使用されるプロパティです。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|図形 ID を返します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|幾何学的図形の Shape オブジェクトを返します。 読み取り専用です。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|図形グループの図形の数を返します。 読み取り専用です。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|コレクション内の位置に基づいて図形を取得します。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|画像の形式を返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.image#id)|画像オブジェクトの図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.image#shape)|画像に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|指定されたコネクタの始点を指定された図形に接続します。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|指定されたコネクタの終点を指定された図形に接続します。|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|線のコネクタの種類を表します。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|指定されたコネクタの始点を図形から切り離します。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|指定されたコネクタの終点を図形から切り離します。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|指定された線の始点が接続されている図形を表します。 読み取り専用です。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|コネクタの始点が接続されている結合点を表します。 読み取り専用です。 線の始点がどの図形にも接続されていない場合は、null を返します。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|指定された線の終点が接続されている図形を表します。 読み取り専用です。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|コネクタの終点が接続されている結合点を表します。 読み取り専用です。 線の終点がどの図形にも接続されていない場合は、null を返します。|
||[id](/javascript/api/excel/excel.line#id)|図形 ID を表します。 読み取り専用です。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定された線の始点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定された線の終点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.line#shape)|線に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|改ページ オブジェクトを削除します。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|改ページの後の最初のセルを取得します。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|改ページの行インデックスを表します。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|指定された範囲の左上セルの前に改ページを追加します。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|コレクション内の改ページの数を取得します。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|インデックス経由で改ページ オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|コレクション内の手動改ページをすべてリセットします。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、ItemNotFound エラーがスローされます。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、null オブジェクトが返されます。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|タイトル列を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|タイトル行を表す範囲オブジェクトを取得します。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|タイトル行を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|ワークシートの印刷範囲を設定します。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|ワークシートのページ余白を単位で設定します。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|セルを含む列を、印刷時、ワークシートの各ページの左で繰り返すように設定します。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|セルを含む行を、印刷時、ワークシートの各ページの上で繰り返すように設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|印刷時に使用するように指定された単位でページ レイアウトの下余白を表します。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|印刷時に使用するように指定された単位でページ レイアウトのフッター余白を表します。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|印刷時に使用するように指定された単位でページ レイアウトのヘッダー余白を表します。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|印刷時に使用するように指定された単位でページ レイアウトの左余白を表します。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|印刷時に使用するように指定された単位でページ レイアウトの右余白を表します。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|印刷時に使用するように指定された単位でページ レイアウトの上余白を表します。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|横方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|印刷ページのスケール値は 10 から 400 までです。 縦または横方向にページを合わせるように指定されている場合、この値には null を指定できます。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|縦方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。 この範囲によって、並べ替えに使用する特定の値が定義されます。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定されるかどうかを指定します。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|PivotTable 内で指定された範囲の値を計算するために使用される DataHierarchy を取得します。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存されるかどうかを指定します。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|必要なすべての条件とコンテキストを自動的に選択するため、指定したセルを使用して自動的に並べ替えを実行するようピボットテーブルを設定します。 これは、UI から自動並べ替えを適用するのと同じ動作です。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|ピボットテーブルでユーザーがデータの値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|ピボット テーブルで並べ替えを実行する際にユーザー設定リストを使用するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|現在の範囲から対象の範囲までの範囲に値を設定します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|データ型を含む範囲セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|ワークシート内で範囲セルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の範囲にセル データまたは書式設定をコピーします。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|現在の範囲に対してフラッシュ フィルを実行します。フラッシュ フィルでは、パターンを感知して自動的にデータが設定されるので、範囲は単一列範囲で、かつパターンを検出できるように周囲にデータが存在する必要があります。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|2D 配列を返します。各セルのフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|一次元配列を返します。各列のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された列内の列間で一貫性のないプロパティについては、null が返されます。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|一次元配列を返します。各行のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された行内の列間で一貫性のないプロパティについては、null が返されます。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の範囲を構成する RangeAreas オブジェクトを取得します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|範囲と重なるテーブルの集まりを範囲限定で取得します。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|各セルのデータ型の状態を表します。 読み取り専用です。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|列によって指定される範囲から重複する値を削除します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|現在の範囲内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|セル プロパティの 2D 配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|列プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|次の再計算が発生したときに再計算する範囲を設定します。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|行プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|RangeAreas のすべてのセルを計算します。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|この RangeAreas オブジェクトを構成する各領域で値、フォーマット、塗りつぶし、罫線などを消去します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|RangeAreas 内でデータ型を含むすべてのセルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|RangeAreas 内のすべてのセルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の RangeAreas にセル データまたは書式設定をコピーします。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|RangeAreas の列全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11, H2" を表す場合、列 "B:E, H:H" を表す RangeAreas が返されます)。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|RangeAreas の行全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11" を表す場合、行 "4:11" を表す RangeAreas が返されます)。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、ItemNotFound エラーがスローされます。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、null オブジェクトが返されます。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|特定の行と列のオフセットによってシフトされる RangeAreas オブジェクトを返します。 返される RangeAreas のディメンションは元のオブジェクトと一致します。 結果の RangeAreas がワークシート グリッドの境界線の外にはみ出る場合、エラーがスローされます。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、エラーがスローされます。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、null オブジェクトを返します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|この RangeAreas オブジェクトの範囲と重なるテーブルの集まりを範囲限定で返します。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[address](/javascript/api/excel/excel.rangeareas#address)|A1 スタイルで RageAreas 参照を返します。 address 値にはセルの各長方形ブロックのワークシート名が含まれます ("Sheet1!A1:B4, Sheet1!D1:D4" など)。 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|ユーザー ロケールで RageAreas 参照を返します。 読み取り専用です。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|この RangeAreas オブジェクトを構成する長方形範囲の数を返します。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|この RangeAreas オブジェクトを構成する長方形範囲の集まりを返します。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|RangeAreas オブジェクトのセル数を返します。すべての個別長方形範囲のセル数が合計されます。 セル数が 2^31-1 (2,147,483,647) を超える場合、-1 を返します。 読み取り専用です。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|この RangeAreas オブジェクトのセルと交差する ConditionalFormats の集まりを返します。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareas#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。 読み取り専用です。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|この RangeAreas オブジェクトの全範囲が列全体を表すかどうかを示します ("A:C, Q:Z" など)。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|この RangeAreas オブジェクトの全範囲が行全体を表すかどうかを示します ("1:3, 5:7" など)。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|現在の RangeAreas のワークシートを返します。 読み取り専用です。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|次の再計算が発生したときに再計算する RangeAreas を設定します。|
||[style](/javascript/api/excel/excel.rangeareas#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|RangeCollection 内の範囲数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|RangeCollection 内のその位置に基づいて範囲オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|一致方法として完全一致か部分一致を指定します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|`rowIndex` プロパティを表します。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 完全一致は、セルの内容全体に一致します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|検索の方向を指定します。 既定値は前方向です。 Excel.SearchDirection をご覧ください。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|`format` プロパティを表します。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|`hyperlink` プロパティを表します。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|`style` プロパティを表します。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|`columnHidden` プロパティを表します。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel. CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties # 形式)|`format` プロパティを表します。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel. CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties # 形式)|`format` プロパティを表します。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|`rowHidden` プロパティを表します。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|ワークシートから図形を削除します。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。 DPI は 96 です。 サポートされている形式は、`Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG`、`Excel.PictureFormat.GIF` だけです。|
||[height](/javascript/api/excel/excel.shape#height)|図形の高さをポイント数で表します。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|指定したポイント数だけ水平方向に図形を移動します。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|z 軸を中心に、指定された度数だけ、図形を時計回りに回転します。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|指定したポイント数だけ垂直方向に図形を移動します。|
||[left](/javascript/api/excel/excel.shape#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shape#name)|図形の名前を表します。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|この図形の結合点の数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shape#fill)|この図形の塗りつぶしの書式設定を返します。 読み取り専用です。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|図形に関連付けられた幾何学的図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[group](/javascript/api/excel/excel.shape#group)|図形に関連付けられた図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[id](/javascript/api/excel/excel.shape#id)|図形 ID を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.shape#image)|図形に関連付けられた画像を返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[level](/javascript/api/excel/excel.shape#level)|指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.shape#line)|図形に関連付けられた線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|この図形の線の書式設定を返します。 読み取り専用です。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|図形がアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|図形が非アクティブになると発生します。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|この図形の親グループを表します。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|この図形のテキスト フレーム オブジェクトを返します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.shape#type)|この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。 読み取り専用です。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|図形の回転を角度で表します。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の高さを変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の高さに対して拡大または縮小されます。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の幅を変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の幅に対して拡大または縮小されます。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[top](/javascript/api/excel/excel.shape#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[visible](/javascript/api/excel/excel.shape#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shape#width)|図形の幅 (ポイント数) を表します。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|図形がアクティブにされたワークシートの ID を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|幾何学的図形をワークシートに追加します。 新しい図形を表す Shape オブジェクトを返します。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|このコレクションのワークシート内の図形のサブセットをグループ化します。 図形の新しいグループを表す Shape オブジェクトを返します。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|base64 エンコード文字列から画像を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|ワークシートに行を追加します。 新しい行を表す Shape オブジェクトを返します。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|指定されたテキストを含むテキスト ボックスをワークシートに追加します。 新しいテキスト ボックスを表す Shape オブジェクトを返します。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|ワークシートの図形数を返します。 読み取り専用です。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|コレクション内の位置を使用して図形を取得します。|
||[items](/javascript/api/excel/excel.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|非アクティブにされた図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|図形が非アクティブにされたワークシートの ID を取得します。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[type](/javascript/api/excel/excel.shapefill#type)|図形の塗りつぶしの種類を返します。 読み取り専用です。 詳細については、Excel.ShapeFillType をご覧ください。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|図形の塗りつぶしの書式設定を均一な色に設定します。 これにより、塗りつぶしの種類が "Solid" に変更されます。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefont#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefont#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[size](/javascript/api/excel/excel.shapefont#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|グループに関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Shape オブジェクトのコレクションを返します。 読み取り専用です。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|指定した図形グループに含まれるグループ化された図形のグループを解除します。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|並べ替え基準となるリッチな値のターゲット プロパティ名である下位フィールドを表します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|コレクション内のスタイルの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|コレクション内の位置に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|テーブルの AutoFilter オブジェクトを表します。 読み取り専用です。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|追加されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|テーブルが追加されたワークシートの ID を取得します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|変更の詳細に関する情報を表します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|新しいテーブルがブックに追加されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|指定されたテーブルがブックで削除されたときに発生します。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|イベントのソースを指定します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|削除されたテーブルの ID を指定します。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|削除されたテーブルの名前を指定します。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|イベントの種類を指定します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|テーブルが削除されたワークシートの ID を指定します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|コレクションに含まれるテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|コレクション内の最初のテーブルを取得します。 一番上の左のテーブルがコレクション内で最初のテーブルになるように、コレクションのテーブルが上から下へ、左から右への順で並べ替えられます。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|テキスト フレーム内のテキストをすべて削除します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|テキスト フレームにテキストが含まれるかどうかを指定します。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。 詳細については、Excel.TextRange を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|指定された範囲の部分文字列に対する TextRange オブジェクトを返します。|
||[font](/javascript/api/excel/excel.textrange#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、このステートメントを呼び出すと、例外がスローされます。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、null オブジェクトを返します。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|ブックが複数のユーザーによって編集されている場合 (共同編集)、true となります。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|ブックから現在選択されている 1 つまたは複数の範囲を取得します。 getSelectedRange() の場合と同様に、このメソッドは、選択されているすべての範囲を表す RangeAreas オブジェクトを返します。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|ブックが自動保存モードかどうかを指定します。 読み取り専用です。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Excel 計算エンジンのバージョンとして数字を返します。 読み取り専用です。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|ブックで autoSave の設定が変更されると発生します。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|ブックがローカル環境またはオンライン環境に保存されたかどうかを指定します。 読み取り専用です。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|アドレスまたは名前で指定され、1 つまたは複数の長方形範囲ブロックを表す RangeAreas オブジェクトを取得します。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|ワークシートの AutoFilter オブジェクトを表します。 読み取り専用です。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|ワークシートの水平改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|フォーマットが特定のワークシートで変更されたときに発生します。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|ワークシート上のすべての Shape オブジェクトをまとめて返します。 読み取り専用です。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|ワークシートの垂直改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|現在のワークシート内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|変更の詳細に関する情報を表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|ブックのワークシートが変更されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|ブック内のワークシートの書式が変更されたときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|ワークシートで選択範囲を変更したときに発生します。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 完全一致は、セルの内容全体に一致します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
