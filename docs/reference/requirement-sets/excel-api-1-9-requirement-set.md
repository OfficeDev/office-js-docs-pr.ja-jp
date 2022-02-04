---
title: Excel JavaScript API 要件セット 1.9
description: ExcelApi 1.9 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
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

次の表に、JavaScript API 要件セット 1.9 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.9 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|アプリケーションの計算の状態を返します。|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|反復計算の設定を返します。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|次の呼び出しが呼び出されるまで、画面の更新 `context.sync()` を中断します。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|範囲にオートフィルターを適用します。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|オートフィルターのフィルター条件と並べ替え状態をクリアします。|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|オートフィルターが有効になっている場合に指定します。|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|オートフィルターを `Range` 適用する範囲を表すオブジェクトを返します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|オートフィルターを `Range` 適用する範囲を表すオブジェクトを返します。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|オートフィルターにフィルター条件がある場合に指定します。|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|その範囲で現在指定されている Autofilter オブジェクトを適用します。|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|範囲の AutoFilter を削除します。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|1 つの境界線の `color` プロパティを表します。|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|1 つの境界線の `style` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|1 つの境界線の `tintAndShade` プロパティを表します。|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|1 つの境界線の `weight` プロパティを表します。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|`format.borders.bottom` プロパティを表します。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|`format.borders.diagonalDown` プロパティを表します。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|`format.borders.diagonalUp` プロパティを表します。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|`format.borders.horizontal` プロパティを表します。|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|`format.borders.left` プロパティを表します。|
||[right](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|`format.borders.right` プロパティを表します。|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|`format.borders.top` プロパティを表します。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|`format.borders.vertical` プロパティを表します。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|`addressLocal` プロパティを表します。|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|`hidden` プロパティを表します。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|`format.fill.color` プロパティを表します。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|`format.fill.pattern` プロパティを表します。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|`format.fill.patternColor` プロパティを表します。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|`format.fill.patternTintAndShade` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|`format.fill.tintAndShade` プロパティを表します。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|`format.font.bold` プロパティを表します。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|`format.font.color` プロパティを表します。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|`format.font.italic` プロパティを表します。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|`format.font.name` プロパティを表します。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|`format.font.size` プロパティを表します。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|`format.font.strikethrough` プロパティを表します。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|`format.font.subscript` プロパティを表します。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|`format.font.superscript` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|`format.font.tintAndShade` プロパティを表します。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|`format.font.underline` プロパティを表します。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|`autoIndent` プロパティを表します。|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|`borders` プロパティを表します。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|`fill` プロパティを表します。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|`font` プロパティを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|`horizontalAlignment` プロパティを表します。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|`indentLevel` プロパティを表します。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|`protection` プロパティを表します。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|`readingOrder` プロパティを表します。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|`shrinkToFit` プロパティを表します。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|`textOrientation` プロパティを表します。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|`useStandardHeight` プロパティを表します。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|`useStandardWidth` プロパティを表します。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|`verticalAlignment` プロパティを表します。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|`wrapText` プロパティを表します。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|`format.protection.formulaHidden` プロパティを表します。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|`format.protection.locked` プロパティを表します。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|変更後の値を表します。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|変更前の値を表します。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|変更後の値の種類を表します。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|変更前の値の種類を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|Excel UI でグラフをアクティブにします。|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|ピボット グラフのオプションをカプセル化します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|グラフの配色を指定します。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|グラフのグラフ領域の角が丸い場合に指定します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|数値の形式がセルにリンクされている場合に指定します。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|ヒストグラム グラフまたはパレート グラフでビン オーバーフローが有効になっている場合に指定します。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|ヒストグラム グラフまたはパレート グラフでビンアンダーフローが有効になっている場合に指定します。|
||[count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|ヒストグラム グラフまたはパレート グラフのビン数を指定します。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|ヒストグラム グラフまたはパレート グラフのビン オーバーフロー値を指定します。|
||[type](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|ヒストグラム グラフまたはパレート グラフのビンの種類を指定します。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|ヒストグラム グラフまたはパレート グラフのビンアンダーフロー値を指定します。|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|ヒストグラム グラフまたはパレート グラフのビン幅の値を指定します。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|ボックスグラフとひげグラフの四分位計算の種類を指定します。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|ボックスとひげグラフに内側の点を表示する場合に指定します。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|平均線をボックスとひげグラフに表示する場合に指定します。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|平均マーカーをボックスとひげグラフに表示する場合に指定します。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|ボックスとひげグラフに外れ値ポイントを表示する場合に指定します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|セルに番号の書式をリンクする (セル内でラベルが変更された場合に数値の書式が変更される) 場合に指定します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|数値の形式がセルにリンクされている場合に指定します。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|エラー バーに終了スタイル の上限が設定されている場合に指定します。|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|誤差範囲の書式の種類を指定します。|
||[include](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|誤差範囲のどの部分を含めるかを指定します。|
||[type](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|エラー バーを表示するかどうかを指定します。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|グラフの線の書式設定を表します。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|地域マップ グラフの系列マップ ラベル戦略を指定します。|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|地域マップ グラフの系列マッピング レベルを指定します。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|地域マップ グラフの系列投影の種類を指定します。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|軸フィールド ボタンをウィンドウに表示するかどうかを指定ピボットグラフ。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|凡例フィールド ボタンを凡例フィールド ボタンで表示するかどうかを指定ピボットグラフ。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|レポート にレポート フィルター フィールド ボタンを表示するかどうかを指定ピボットグラフ。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|フィールドの [値の表示] ボタンを表示するかどうかを指定ピボットグラフ。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|箱ひげ図グラフのオプションをカプセル化します。|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|地域マップ グラフ系列の最大値の色を指定します。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|地域マップ グラフ系列の最大値の種類を指定します。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|地域マップ グラフ系列の最大値を指定します。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|地域マップ グラフ系列の中点の値の色を指定します。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|地域マップ グラフ系列の中点値の種類を指定します。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|地域マップ グラフ系列の中点の値を指定します。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|地域マップ グラフ系列の最小値の色を指定します。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|地域マップ グラフ系列の最小値の種類を指定します。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|地域マップ グラフ系列の最小値を指定します。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|地域マップ グラフの系列グラデーション スタイルを指定します。|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|系列内の負のデータ ポイントの塗りつぶしの色を指定します。|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|リージョン マップ グラフのオプションをカプセル化します。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|ツリーマップ グラフの系列の親ラベル戦略領域を指定します。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|ウォーターフォール グラフにコネクタ線を表示するかどうかを指定します。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|系列内のデータ ラベルごとに引き出し線を表示するかどうかを指定します。|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|円グラフまたは棒グラフの 2 つのセクションを分割するしきい値を指定します。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|グラフ系列の誤差範囲オブジェクトを表します。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|セルに番号の書式をリンクする (セル内でラベルが変更された場合に数値の書式が変更される) 場合に指定します。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|`addressLocal` プロパティを表します。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|`columnIndex` プロパティを表します。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|1 つ以上 `RangeAreas`の四角形の範囲で構成され、その範囲に対して同じ形式が適用される値を返します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|無効なセル値 `RangeAreas` を持つ 1 つ以上の四角形の範囲を含むオブジェクトを返します。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|無効なセル値 `RangeAreas` を持つ 1 つ以上の四角形の範囲を含むオブジェクトを返します。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|豊富な値に対してリッチ フィルターを実行するためにフィルターで使用されるプロパティ。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|図形 ID を返します。|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|幾何学的図形の `Shape` オブジェクトを返します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|図形グループの図形の数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|コレクション内の位置に基づいて図形を取得します。|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|ワークシートの中央フッター。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|ワークシートの中央ヘッダー。|
||[leftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|ワークシートの左側のフッター。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|ワークシートの左側のヘッダー。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|ワークシートの右側のフッター。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|ワークシートの右側のヘッダー。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|ヘッダー/フッターが設定されている状態。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|画像の形式を返します。|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|イメージ オブジェクトの図形識別子を指定します。|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|イメージに関連付 `Shape` けられているオブジェクトを返します。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|循環参照を解決するために、各反復間のExcelを指定します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|循環参照の解決に使用Excel繰り返しの最大数を指定します。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|指定された線の始点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|指定された線の始点が接続されている図形を表します。|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|コネクタの始点が接続されている結合点を表します。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|指定されたコネクタの始点を指定された図形に接続します。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|指定されたコネクタの終点を指定された図形に接続します。|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|線のコネクタの種類を表します。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|指定されたコネクタの始点を図形から切り離します。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|指定されたコネクタの終点を図形から切り離します。|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|指定された線の終点の矢印の幅を表します。|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|指定された線の終点が接続されている図形を表します。|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|コネクタの終点が接続されている結合点を表します。|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|図形識別子を指定します。|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|指定した線の先頭が図形に接続される場合に指定します。|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|指定した線の端が図形に接続される場合に指定します。|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|行に関連 `Shape` 付けられたオブジェクトを返します。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|ページブレークの列インデックスを指定します。|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|改ページ オブジェクトを削除します。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|改ページの後の最初のセルを取得します。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|ページブレークの行インデックスを指定します。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|指定された範囲の左上セルの前に改ページを追加します。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|コレクション内の改ページの数を取得します。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|インデックス経由で改ページ オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|コレクション内の手動改ページをすべてリセットします。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|ワークシートの白黒印刷オプション。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|ポイントでの印刷に使用するワークシートの下部ページ余白。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|ワークシートの中央に水平方向にフラグを設定します。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|ワークシートの中央に垂直フラグを設定します。|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|ワークシートの下書きモード オプション。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|印刷するワークシートの最初のページ番号。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|印刷時に使用するワークシートのフッター余白をポイントで指定します。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|ワークシートの `RangeAreas` 印刷領域を表す 1 つ以上の四角形の範囲を含むオブジェクトを取得します。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|ワークシートの `RangeAreas` 印刷領域を表す 1 つ以上の四角形の範囲を含むオブジェクトを取得します。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|タイトル行を表す範囲オブジェクトを取得します。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|タイトル行を表す範囲オブジェクトを取得します。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|印刷時に使用するワークシートのヘッダー余白をポイントで指定します。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|ワークシートのヘッダーとフッターの構成。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|印刷時に使用するワークシートの左余白をポイントで指定します。|
||[orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|ワークシートのページの向き。|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|ワークシートのページの用紙サイズ。|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|印刷時にワークシートのコメントを表示する必要がある場合に指定します。|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|ワークシートの印刷エラー オプション。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|ワークシートの枠線を印刷する場合に指定します。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|ワークシートの見出しを印刷する場合に指定します。|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|ワークシートのページ印刷順序オプション。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|印刷時に使用するワークシートの右余白をポイントで指定します。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|ワークシートの印刷範囲を設定します。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|ワークシートのページ余白を単位で設定します。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|セルを含む列を、印刷時、ワークシートの各ページの左で繰り返すように設定します。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|セルを含む行を、印刷時、ワークシートの各ページの上で繰り返すように設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|印刷時に使用するワークシートの上余白をポイントで指定します。|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|ワークシートの印刷ズーム オプション。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|印刷に使用する単位でページ レイアウトの下部余白を指定します。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|印刷に使用する単位のページ レイアウト フッター余白を指定します。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|印刷に使用する単位のページ レイアウト ヘッダー余白を指定します。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|印刷に使用する単位のページ レイアウト左余白を指定します。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|印刷に使用する単位のページ レイアウト右余白を指定します。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|印刷に使用する単位でページ レイアウトの上余白を指定します。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|横方向に合わせるページ数。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|印刷ページのスケール値は 10 から 400 までです。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|縦方向に合わせるページ数。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|書式設定が更新時またはフィールドの移動時に自動的に書式設定される場合を指定します。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|PivotTable 内で指定された範囲の値を計算するために使用される DataHierarchy を取得します。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算される場合に書式設定を保持する場合に指定します。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|必要なすべての条件とコンテキストを自動的に選択するため、指定したセルを使用して自動的に並べ替えを実行するようピボットテーブルを設定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|ピボットテーブルでデータ本文の値をユーザーが編集できる場合を指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|ピボットテーブルが並べ替え時にカスタム リストを使用する場合に指定します。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|指定した AutoFill ロジックを使用して、現在の範囲から移動先の範囲の範囲を塗りつぶしします。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|データ型を持つ範囲セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|ワークシート内の範囲セルをリンクされたデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|セル データまたは書式をソース範囲または現在の範囲 `RangeAreas` にコピーします。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|指定された条件に基づいて指定された文字列を見つけます。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|指定された条件に基づいて指定された文字列を見つけます。|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|現在の範囲にフラッシュ塗りつぶしを実行します。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|2D 配列を返します。各セルのフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|一次元配列を返します。各列のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|一次元配列を返します。各行のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|指定した `RangeAreas` 種類と値に一致するセルを表す、1 つ以上の四角形の範囲を含むオブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|指定した `RangeAreas` 種類と値に一致するセルを表す 1 つ以上の範囲を含むオブジェクトを取得します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|範囲と重なるテーブルの集まりを範囲限定で取得します。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|各セルのデータ型の状態を表します。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|列によって指定される範囲から重複する値を削除します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|現在の範囲内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|セル プロパティの 2D 配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|列プロパティの 1 次元配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|次の再計算が発生したときに再計算する範囲を設定します。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|行プロパティの 1 次元配列に基づいて範囲を更新し、フォント、塗りつぶし、罫線、配置をカプセル化します。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[address](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|A1 スタイル `RangeAreas` の参照を返します。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|ユーザー ロケール内 `RangeAreas` の参照を返します。|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|このオブジェクトを構成する四角形の範囲の数を返 `RangeAreas` します。|
||[areas](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|このオブジェクトを構成する四角形の範囲のコレクションを返 `RangeAreas` します。|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|内のすべてのセルを計算します `RangeAreas`。|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|オブジェクト内のセルの数を `RangeAreas` 返し、個々の四角形のすべての範囲のセル数を合計します。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|このオブジェクトを構成する各領域の値、書式、塗りつぶし、罫線、その他のプロパティをクリア `RangeAreas` します。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|このオブジェクト内のセルと交差する条件付き書式のコレクションを返 `RangeAreas` します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|データ型を含むすべてのセル `RangeAreas` をテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|リンクされたデータ型に、すべてのセル `RangeAreas` を変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|セル データまたは書式をソース範囲または現在の範囲から `RangeAreas` コピーします `RangeAreas`。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|内のすべての範囲のデータ検証オブジェクトを返します `RangeAreas`。|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|オブジェクト内のすべての範囲 `RangeFormat` のフォント、塗りつぶし、罫線、配置、その他のプロパティをカプセル化するオブジェクトを返 `RangeAreas` します。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|`RangeAreas` `RangeAreas` 列全体を表すオブジェクトを返します (`RangeAreas`たとえば、カレントがセル "B4:E11, H2" を表す場合は、列 "B:E, H:H" `RangeAreas` を表す a を返します)。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|`RangeAreas` `RangeAreas` 行全体を表すオブジェクトを返します (`RangeAreas`たとえば、カレントがセル "B4:E11" を表す場合は、行 "4:11" `RangeAreas` を表す a を返します)。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|指定した範囲 `RangeAreas` の交差部分を表すオブジェクトを返します `RangeAreas`。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|指定した範囲 `RangeAreas` の交差部分を表すオブジェクトを返します `RangeAreas`。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|特定の行 `RangeAreas` と列のオフセットによってシフトされるオブジェクトを返します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|指定した型と `RangeAreas` 値に一致するすべてのセルを表すオブジェクトを返します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|指定した型と `RangeAreas` 値に一致するすべてのセルを表すオブジェクトを返します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|このオブジェクト内の任意の範囲と重なるテーブルのスコープ付きコレクションを返 `RangeAreas` します。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|オブジェクト内の個々の `RangeAreas` 四角形範囲のすべての使用領域を含む使用される領域を返 `RangeAreas` します。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|オブジェクト内の個々の `RangeAreas` 四角形範囲のすべての使用領域を含む使用される領域を返 `RangeAreas` します。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|このオブジェクトのすべての範囲が `RangeAreas` 列全体を表す ("A:C、Q:Z"など) を指定します。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|このオブジェクトのすべての範囲が `RangeAreas` 行全体を表す (例: "1:3, 5:7") を指定します。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|次の `RangeAreas` 再計算が行われるときに再計算されるように設定します。|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|このオブジェクトのすべての範囲のスタイルを表 `RangeAreas` します。|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|現在のワークシートを返します `RangeAreas`。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|範囲の境界線の色を明るくまたは暗くする倍数を指定します。値は -1 (最も暗い) ~ 1 (最も明るい) で、元の色の場合は 0 です。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|範囲の境界線の色を明るくまたは暗くする倍数を指定します。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|内の範囲の数を返します `RangeCollection`。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|内の位置に基づいて range オブジェクトを返します `RangeCollection`。|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|範囲のパターン。|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|範囲パターンの色を表す HTML カラー コードは、#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("オレンジ色" など) として表されます。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|範囲塗りつぶしのパターンの色を明るくまたは暗くする倍数を指定します。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|範囲塗りつぶしの色を明るくまたは暗くする倍数を指定します。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|フォントの取り消し線の状態を指定します。|
||[subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|フォントの下付き文字の状態を指定します。|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|フォントの上付き文字の状態を指定します。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|範囲フォントの色を明るくまたは暗くする倍数を指定します。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|テキストの配置が等しい分布に設定されている場合に、テキストが自動的にインデントされる場合に指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|使用可能な列の幅に収まるテキストを自動的に縮小する場合に指定します。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|結果として生じた範囲に存在する残りの一意の行の数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|一致で大文字と小文字が区別される場合を指定します。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|`addressLocal` プロパティを表します。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|`rowIndex` プロパティを表します。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|一致で大文字と小文字が区別される場合を指定します。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|検索の方向を指定します。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|`format` プロパティを表します。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|`hyperlink` プロパティを表します。|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|`style` プロパティを表します。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|`columnHidden` プロパティを表します。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[format: Excel。CellPropertiesFormat & { columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|`format` プロパティを表します。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel。CellPropertiesFormat & { rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|`format` プロパティを表します。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|`rowHidden` プロパティを表します。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|オブジェクトの代替説明テキストを指定 `Shape` します。|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|オブジェクトの代替タイトル テキストを指定 `Shape` します。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|この図形の結合点の数を返します。|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|ワークシートから図形を削除します。|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|この図形の塗りつぶしの書式設定を返します。|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|図形に関連付けられた幾何学的図形を返します。|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|この幾何学的な図形の幾何学的な図形の種類を指定します。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。|
||[group](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|図形に関連付けられた図形グループを返します。|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|図形の高さをポイントで指定します。|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|図形識別子を指定します。|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|図形に関連付けられた画像を返します。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|指定したポイント数だけ水平方向に図形を移動します。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|z 軸を中心に、指定された度数だけ、図形を時計回りに回転します。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|指定したポイント数だけ垂直方向に図形を移動します。|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|指定した図形のレベルを指定します。|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|図形に関連付けられた線を返します。|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|この図形の線の書式設定を返します。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|この図形の縦横比をロックする場合に指定します。|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|図形の名前を指定します。|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|図形がアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|図形が非アクティブになると発生します。|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|この図形の親グループを指定します。|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|図形の回転角度を度で指定します。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|指定した係数分だけ図形の高さを変更します。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|指定した係数分だけ図形の幅を変更します。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|この図形のテキスト フレーム オブジェクトを返します。|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|この図形の種類を返します。|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|図形が表示される場合に指定します。|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|図形の幅をポイント単位で指定します。|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|図形をアクティブ化するワークシートの ID を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|幾何学的図形をワークシートに追加します。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|このコレクションのワークシート内の図形のサブセットをグループ化します。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|base64 エンコード文字列から画像を作成し、それをワークシートに追加します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|ワークシートに行を追加します。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|指定されたテキストを含むテキスト ボックスをワークシートに追加します。|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|ワークシートの図形数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|コレクション内の位置を使用して図形を取得します。|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|非アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|図形が非アクティブ化されているワークシートの ID を取得します。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|図形塗りつぶしの前景色を #RRGGBB HTML の色形式で表します ("FFA500"など) 形式で、または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|図形の塗りつぶしの書式設定を均一な色に設定します。|
||[transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|塗りつぶしの透明度の割合を 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
||[type](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|図形の塗りつぶしの種類を返します。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|テキストの色の HTML カラー コード表現 (例: "#FF0000" は赤を表します)。|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|フォント名 ("Calibri" など) を表します。|
||[size](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|フォント サイズをポイント (11 など) で表します。|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|フォントに適用する下線の種類。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|図形識別子を指定します。|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|グループに関連付 `Shape` けられているオブジェクトを返します。|
||[shapes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|オブジェクトのコレクションを返 `Shape` します。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|指定した図形グループに含まれるグループ化された図形のグループを解除します。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|線の色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500" など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|図形の線スタイルを表します。|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|図形の線スタイルを表します。|
||[transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|図形要素の線の書式設定が表示される場合に指定します。|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|線の太さ (ポイント数) を表します。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|並べ替えるリッチ値のターゲット プロパティ名であるサブフィールドを指定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|コレクション内のスタイルの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|コレクション内の位置に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|テーブルのオブジェクト `AutoFilter` を表します。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|追加されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|テーブルを追加するワークシートの ID を取得します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|変更の詳細に関する情報を取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|ブックに新しいテーブルが追加された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|指定されたテーブルがブックで削除されたときに発生します。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|削除されるテーブルの ID を取得します。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|削除されるテーブルの名前を取得します。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|テーブルが削除されるワークシートの ID を取得します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|コレクションに含まれるテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|コレクション内の最初のテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|名前または ID でテーブルを取得します。|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|テキスト フレームの自動サイズ設定。|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|テキスト フレーム内のテキストをすべて削除します。|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|テキスト フレームにテキストが含まれている場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|テキスト フレームの水平方向の配置を表します。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|テキスト フレームの水平方向のオーバーフローの動作を表します。|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|テキスト フレームの方向を指定する角度を表します。|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|テキスト フレームの読む方向を表します (左から右または右から左)。|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|テキスト フレームの右余白を表します (ポイント数)。|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|テキスト フレームの垂直方向の配置を表します。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|テキスト フレームの垂直方向のオーバーフローの動作を表します。|
|[TextRange](/javascript/api/excel/excel.textrange)|[font](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|テキスト範囲の `ShapeFont` フォント属性を表すオブジェクトを返します。|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|指定された範囲の部分文字列に対する TextRange オブジェクトを返します。|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|ブックが自動保存モードの場合に指定します。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Excel 計算エンジンのバージョンとして数字を返します。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|ブックで現在アクティブになっているグラフを取得します。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|ブックで現在アクティブになっているグラフを取得します。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|ブックが `true` 複数のユーザーによって編集されている場合 (共同編集による) 場合に返します。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|ブックから現在選択されている 1 つまたは複数の範囲を取得します。|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|ブックが最後に保存された後に変更が行われた場合に指定します。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|ブックで AutoSave 設定が変更された場合に発生します。|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|ブックがローカルまたはオンラインで保存された場合に指定します。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|イベントの種類を取得します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|ワークシートのオブジェクト `AutoFilter` を表します。|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|必要に応Excelワークシートを再計算する必要があるかどうかを決定します。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|指定された条件に基 `RangeAreas` づいて、指定された文字列のすべての出現を検索し、1 つ以上の四角形の範囲で構成されるオブジェクトとして返します。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|指定された条件に基 `RangeAreas` づいて、指定された文字列のすべての出現を検索し、1 つ以上の四角形の範囲で構成されるオブジェクトとして返します。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|アドレスまたは名前 `RangeAreas` で指定された、四角形の範囲の 1 つ以上のブロックを表すオブジェクトを取得します。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|ワークシートの水平改ページをまとめて取得します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|フォーマットが特定のワークシートで変更されたときに発生します。|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|ワークシートの `PageLayout` オブジェクトを取得します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|現在のワークシート内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|ワークシート上のすべての Shape オブジェクトをまとめて返します。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|ワークシートの垂直改ページをまとめて取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|変更の詳細に関する情報を表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|ブックのワークシートが変更されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|ブック内のワークシートの形式が変更された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|ワークシートで選択範囲を変更したときに発生します。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|データが変更されたワークシートの ID を取得します。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|一致が完了する必要がある場合と部分的に行う必要がある場合に指定します。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|一致で大文字と小文字が区別される場合を指定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
