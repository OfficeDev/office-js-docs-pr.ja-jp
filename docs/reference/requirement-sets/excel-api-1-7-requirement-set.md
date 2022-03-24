---
title: Excel JavaScript API 要件セット 1.7
description: ExcelApi 1.7 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: cd8f0f333b76306a6feecff95b9ba8831428606a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744528"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 の新機能

Excel JavaScript API 要件セット 1.7 の機能には、グラフ、イベント、ワークシート、範囲、ドキュメント プロパティ、名前付きアイテム、保護のオプションとスタイルに対応する API が含まれます。

## <a name="customize-charts"></a>グラフのカスタマイズ

新しいグラフ API を使用して実行できる操作には、他の種類のグラフの作成、グラフへのデータ系列の追加、グラフ タイトルの設定、軸タイトルの追加、表示単位の追加、移動平均を使用した近似曲線の追加、線形近似曲線への変更などがあります。 以下はその一例です。

- グラフの軸: グラフ内の軸の単位、ラベル、タイトルを取得、設定、書式設定する。
- グラフのデータ系列: グラフ内のデータ系列を追加、設定、削除する。  データ系列マーカー、プロット順序、サイジングを変更する。
- グラフの近似曲線: グラフ内の近似曲線を追加、取得、書式設定する。
- グラフの凡例: グラフ内の凡例のフォントを書式設定する。
- グラフ データ ポイント要素: データ要素の色を設定する。
- グラフ タイトルのサブ文字列: グラフ タイトルのサブ文字列を取得、設定する。
- グラフの種類: 他の種類のグラフを作成するオプションを使用する。

## <a name="events"></a>イベント

Excel イベント API には各種のイベント ハンドラーが用意されています。これらのハンドラーを使用することで、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できます。 実行する関数は、目的のシナリオに必要な処理を行うように設計できます。 現在利用可能なイベントのリストについては、[Excel JavaScript API を使用してイベントを操作する](../../excel/excel-add-ins-events.md)を参照してください。

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>ワークシートと範囲の外観のカスタマイズ

新しい API を使用して、ワークシートの外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

- ワークシート内でスクロールするときに特定の行または列が常に表示されるよう、ウィンドウ枠を固定する。 たとえば、ワークシート内の最初の行にヘッダーが示される場合、その行にウィンドウ枠を固定すると、ワークシートをスクロールダウンしても列の見出しは表示されたままになります。
- ワークシートのタブの色を変更する。
- ワークシートの見出しを追加する。

範囲の外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

- 特定の範囲に対してセルのスタイルを設定し、その範囲内のすべてのセルに一貫した書式設定が適用されるようにする。 セルのスタイルとは、フォント、フォントのサイズ、数値形式、セルの罫線、セルの網掛けなど、文字に定義された書式設定一式を指します。 Excel の組み込みのセル スタイルのいずれかを使用することも、独自のカスタム セル スタイルを作成することもできます。
- 範囲に適用するテキストの向きを設定する。
- 特定の範囲からブック内の別の場所または外部の場所にリンクするハイパーリンクを追加または変更する。

## <a name="manage-document-properties"></a>ドキュメント プロパティの管理

ドキュメント プロパティ API を使用して、組み込みのドキュメント プロパティにアクセスできます。また、ブックの状態を格納してワークフローやビジネス ロジックを操作するためのカスタム ドキュメント プロパティを作成、管理することもできます。

## <a name="copy-worksheets"></a>ワークシートのコピー

ワークシート コピー API を使用して、ワークシートのデータと書式設定を同じブック内の新しいワークシートにコピーできます。これにより、必要となるデータの転送量を削減することができます。

## <a name="handle-ranges-with-ease"></a>範囲の操作の簡易化

各種の範囲 API を使用して、周りの領域の取得や範囲のサイズ変更など、さまざまな操作を行うことができます。 これらの API により、範囲の操作やアドレス指定などのタスクが効率化されます。

さらに、次の機能も使用できます。

- ブックとワークシートの保護オプション: これらの API を使用して、ワークシートおよびブック構造内のデータを保護する。
- 名前付きアイテムの更新: この API を使用して、名前付きアイテムを更新する。
- アクティブ セルの取得: この API を使用して、ブックのアクティブ セルを取得する。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.7 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.7 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|グラフの種類を指定します。|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|グラフの一意の ID。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|すべてのフィールド ボタンを 1 つのウィンドウに表示するかどうかをピボットグラフ。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[罫線](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|色、線のスタイル、太さなど、グラフ領域の罫線の形式を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType、 group?: Excel。ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|種類とグループで識別された特定の軸を返します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|指定した軸のグループを指定します。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|指定したカテゴリ軸の基本単位を指定します。|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|カテゴリ軸の種類を指定します。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|ユーザー設定の軸表示単位の値を指定します。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|軸の表示単位を表します。|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|グラフ軸の高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|軸の左端からグラフ領域の左側までの距離をポイントで指定します。|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|対数スケールを使用する場合の対数の基数を指定します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|指定した軸の目盛の種類を指定します。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|プロパティがに設定されている場合、カテゴリ軸の `categoryType` メジャー単位スケール値を指定します `dateAxis`。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|指定した軸の目盛りの種類を指定します。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|プロパティがに設定されている場合、カテゴリ軸の `categoryType` マイナー単位スケール値を指定します `dateAxis`。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|最後から最初Excelデータ ポイントをプロットする方法を指定します。|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|値軸のスケールの種類を指定します。|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|指定した軸のすべてのカテゴリ名を設定します。|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|軸の表示単位をカスタム値に設定します。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|軸表示単位ラベルが表示される場合に指定します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|指定された軸の目盛ラベルの位置を指定します。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|目盛ラベル間のカテゴリまたは系列の数を指定します。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|目盛の間のカテゴリまたは系列の数を指定します。|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|軸の上端からグラフ領域の上端までの距離をポイントで指定します。|
||[type](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|軸の種類を指定します。|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|軸が表示される場合に指定します。|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|グラフ軸の幅をポイント単位で指定します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|罫線のスタイルを表します。|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|罫線の太さ (ポイント数) を表します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|データ ラベルの位置を表す値。|
||[区切り記号](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|データ ラベルのバブル サイズが表示される場合に指定します。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|データ ラベル のカテゴリ名が表示される場合に指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|データ ラベルの凡例キーが表示される場合に指定します。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|データ ラベルの割合を表示する場合に指定します。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|データ ラベルの系列名が表示される場合に指定します。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|データ ラベルの値が表示される場合に指定します。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|グラフ文字オブジェクトのフォント名、フォント サイズ、色などのフォント属性を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|グラフ上の凡例の高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|グラフ上の凡例の左の値をポイントで指定します。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|凡例に含まれる凡例エントリのコレクションを表します。|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|凡例にグラフに影が付く場合を指定します。|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|グラフの凡例の上部を指定します。|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|グラフ上の凡例の幅をポイント単位で指定します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|グラフの凡例エントリの表示を表します。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|コレクション内の凡例エントリの数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|指定したインデックスの凡例エントリを返します。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|線のスタイルを表します。|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|線の太さ (ポイント数) を表します。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|グラフ データ ポイントのデータ ラベルを返します。|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|データ ポイントにデータ ラベルが含されているかどうかを表します。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|データ ポイントのマーカー背景色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|データ ポイントのマーカーの前景色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|データ ポイントのマーカー スタイルを表します。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[罫線](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|色、スタイル、および重み情報を含むグラフ データ ポイントの罫線の形式を表します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|グラフ系列の種類を表します。|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|グラフ系列を削除します。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|グラフ系列のドーナツの穴の大きさを表します。|
||[フィルター処理](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|系列をフィルター処理する場合に指定します。|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|グラフ系列間に設けられる間隔を表します。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|系列にデータ ラベルが含む場合を指定します。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|グラフ系列のマーカーの背景色を指定します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|グラフ系列のマーカーの前景色を指定します。|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|グラフ系列のマーカー サイズを指定します。|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|グラフ系列のマーカー スタイルを指定します。|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|グラフ グループ内のグラフ系列のプロット順序を指定します。|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|グラフ系列のバブル サイズを設定します。|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|グラフ系列の値を設定します。|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|グラフ系列の x 軸の値を設定します。|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|系列に影が付く場合を指定します。|
||[スムーズ](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|系列が滑らかな場合に指定します。|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|系列内の傾向線のコレクション。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|コレクションに新しい系列を追加します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|グラフタイトルの部分文字列を取得します。|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|グラフ タイトルの高さ (ポイント数) を返します。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|グラフタイトルの水平方向の配置を指定します。|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|グラフ タイトルの左端からグラフ領域の左端までの距離をポイントで指定します。|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|グラフ タイトルの位置を表します。|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|A1 スタイルの表記法を使用するグラフ タイトルの数式を表す文字列値を設定します。|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|グラフ タイトルのテキストの向きを指定します。|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|グラフ タイトルの上端からグラフ領域の上端までの距離をポイントで指定します。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|グラフ タイトルの垂直方向の配置を指定します。|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|グラフ タイトルの幅をポイント単位で指定します。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[罫線](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|色、線のスタイル、太さなど、グラフタイトルの罫線の形式を表します。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|trendline オブジェクトを削除します。|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|グラフの近似曲線の書式設定を表します。|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|近似曲線の切片の値を表します。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|グラフの傾向線の期間を表します。|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|近似曲線の名前を表します。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|グラフの傾向線の順序を表します。|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|近似曲線のコレクションに新しい近似曲線を追加します。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|コレクションに含まれる近似曲線の数を返します。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|items 配列の挿入順序である、インデックス別の trendline オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|グラフの線の書式設定を表します。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|カスタム プロパティのキー。|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|カスタム プロパティに使用される値の種類。|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|カスタム プロパティの値を指定します。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|コレクション内のすべてのデータ接続を更新します。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|ブックの作成者。|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|ブックのカテゴリ。|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|ブックのコメント。|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|ブックの会社。|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|ブックの作成日を取得します。|
||[カスタム](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|ブックのカスタム プロパティのコレクションを取得します。|
||[キーワード](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|ブックのキーワード。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|ブックの最後の作成者を取得します。|
||[上司](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|ブックのマネージャー。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|ブックのリビジョン番号を取得します。|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|ブックの件名。|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|ブックのタイトル。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|名前付きアイテムの値と型を含むオブジェクトを返します。|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|名前付きアイテムの数式。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|名前付きアイテム配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|現在の `Range` オブジェクトと同じ左上 `Range` のセルを持つオブジェクトを取得しますが、指定した行数と列数を持つオブジェクトを取得します。|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|範囲を base64 エンコードされた png イメージとしてレンダリングします。|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|この範囲の `Range` 左上のセルの周囲の領域を表すオブジェクトを返します。|
||[hyperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|現在の範囲のハイパーリンクを表します。|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|現在の範囲が列全体であるかどうかを表します。|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|現在の範囲が行全体であるかどうかを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|ユーザーのExcelに基づいて、指定した範囲の数値書式コードを表します。|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|アクティブ セルに多数の値が含まれる場合、そのセルのカードを表示します。|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|現在の範囲のスタイルを表します。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|範囲内のすべてのセルのテキストの向き。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|オブジェクトの行の高さが `Range` シートの標準の高さと等しいかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|オブジェクトの列幅が `Range` シートの標準幅と等しい場合に指定します。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|ハイパーリンクの URL ターゲットを表します。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|ハイパーリンクのドキュメント参照ターゲットを表します。|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|ハイパーリンクの上にカーソルを合わせると表示される文字列を表します。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|該当する範囲内の左上端のセルに表示される文字列を表します。|
|[スタイル](/javascript/api/excel/excel.style)|[borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|4 つの罫線のスタイルを表す 4 つの罫線オブジェクトのコレクション。|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|スタイルが組み込みのスタイルである場合に指定します。|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|このスタイルを削除します。|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|スタイル `Font` のフォントを表すオブジェクト。|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|ワークシートを保護するときに数式を非表示に設定する場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|スタイルでの水平方向の配置を表します。|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|スタイルに自動インデント、水平方向の配置、垂直方向の配置、折り返しテキスト、インデント レベル、およびテキストの向きのプロパティが含まれる場合を指定します。|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|スタイルに色、色インデックス、線のスタイル、太さ罫線のプロパティが含まれる場合に指定します。|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|スタイルに背景、太字、色、色インデックス、フォント スタイル、斜体、名前、サイズ、取り消し線、下付き文字、下線のフォント プロパティが含まれる場合に指定します。|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|スタイルに number format プロパティが含まれる場合に指定します。|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|スタイルに色、色インデックス、負の場合は反転、パターン、パターンの色、パターンの色インデックスの内部プロパティを含む場合を指定します。|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|スタイルに非表示およびロックされた保護プロパティの数式が含まれる場合に指定します。|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|ワークシートが保護されているときにオブジェクトがロックされる場合に指定します。|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|スタイルの名前。|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|スタイルで適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|使用可能な列の幅に収まるテキストを自動的に縮小する場合に指定します。|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|スタイルの垂直方向の配置を指定します。|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|オブジェクト内のExcelを折り返す値を指定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|コレクションに新しいスタイルを追加します。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|名前で取得 `Style` します。|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|セル内のデータが特定のテーブルで変更された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|特定のテーブルで選択範囲が変更された場合に発生します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|特定のワークシート上のテーブル内で変更されたエリアを表すアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|変更されたイベントのトリガー方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|データが変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|データが変更されたワークシートの ID を取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|ブックまたはワークシート内の任意のテーブルでデータが変更された場合に発生します。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|特定のワークシート上のテーブル内で選択されたエリアを表す範囲のアドレスを取得します。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|選択範囲がテーブル内にある場合に指定します。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|選択範囲が変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|選択範囲が変更されたワークシートの ID を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|ブック内のすべてのデータ接続を表します。|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|ブックで現在アクティブなセルを取得します。|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|ブックの名前を取得します。|
||[プロパティ](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|ブックのプロパティを取得します。|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|ブックの保護オブジェクトを返します。|
||[スタイル](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|ブックに関連付けられているスタイルのコレクションを表します。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|ブックを保護します。|
||[保護](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|ブックが保護される場合に指定します。|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|ブックの保護を解除します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel.WorksheetPositionType、 relativeTo?: Excel。ワークシート)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|ワークシートをコピーし、指定した位置に配置します。|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|ワークシートの固定されたウィンドウを操作するために使用できるオブジェクトを取得します。|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|特定の行 `Range` インデックスと列インデックスから始まり、特定の数の行と列にまたがるオブジェクトを取得します。|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|ワークシートがアクティブ化されると発生します。|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|特定のワークシートでデータが変更された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|ワークシートが非アクティブ化された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|特定のワークシートで選択範囲が変更された場合に発生します。|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|ワークシート内のすべての列の標準 (既定) の幅を指定します。|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|ワークシートのタブの色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|アクティブ化されたワークシートの ID を取得します。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|ブックに追加されるワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|変更されたイベントのトリガー方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|データが変更されたワークシートの ID を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|ブック内のワークシートがアクティブ化されると発生します。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|新しいワークシートがブックに追加された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|ブック内のワークシートが非アクティブ化された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|ワークシートがブックから削除された場合に発生します。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|非アクティブ化されたワークシートの ID を取得します。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|ブックから削除されたワークシートの ID を取得します。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|アクティブなワークシート ビューに固定セルを設定します。|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|ワークシートの最初の列または列を固定します。|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|ワークシートの一番上の行または行を固定します。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|ワークシートからすべての固定ウィンドウを削除します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|ワークシートの保護を解除します。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|オブジェクトの編集を許可するワークシート保護オプションを表します。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|シナリオの編集を許可するワークシート保護オプションを表します。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|選択モードのワークシート保護オプションを表します。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|特定のワークシートで選択されたエリアを表す範囲のアドレスを取得します。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|選択範囲が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
