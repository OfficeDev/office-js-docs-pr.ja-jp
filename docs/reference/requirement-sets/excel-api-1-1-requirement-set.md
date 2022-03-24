---
title: Excel JavaScript API 要件セット 1.1
description: ExcelApi 1.1 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45061afc7e401e18a67377bf88fa1670bb7a8ece
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745956"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要件セット 1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。 この要件は、Excelによってサポートされる唯一の要件セットExcel 2016。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.1 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.1 でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true) の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Excel で現在開いているすべてのブックを再計算します。|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|内の定数で定義されているブックで使用される計算モードを返します `Excel.CalculationMode`。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|バインディングによって表される範囲を返します。|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|バインドによって表されるテーブルを返します。|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|バインドによって表されるテキストを返します。|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|バインド識別子を表します。|
||[type](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|バインドの種類を返します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|コレクション内にあるバインドの数を取得します。|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|ID によってバインド オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|グラフの軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|グラフのデータ ラベルを表します。|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|グラフ オブジェクトを削除します。|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|グラフ領域の書式設定プロパティをカプセル化します。|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|グラフ オブジェクトの高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[凡例](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|グラフの凡例を表します。|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|グラフ オブジェクトの名前を指定します。|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|グラフの元データをリセットします。|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|ワークシート上のセルを基準にしてグラフを配置します。|
||[title](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|オブジェクトの上端から行 1 の上端までの距離 (ワークシート上) またはグラフ領域の上端 (グラフ上) をポイントで指定します。|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|グラフ オブジェクトの幅をポイント単位で指定します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|グラフの項目軸を表します。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|3-D グラフの系列軸を表します。|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|軸の数値軸を表します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|線とフォントの書式設定を含むグラフ オブジェクトの書式設定を表します。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|指定した軸の主グリッド線を表すオブジェクトを返します。|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|2 つの大きい目盛の間隔を表します。|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|数値軸の最大値を表します。|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|数値軸の最小値を表します。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|指定した軸の小さい枠線を表すオブジェクトを返します。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|2 つの小さい目盛の間隔を表します。|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|軸タイトルを表します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|グラフの線の書式設定を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|グラフ軸のタイトルの書式を指定します。|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|軸のタイトルを指定します。|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|軸のタイトルが表示される場合に指定します。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|グラフ軸タイトル オブジェクトのグラフ軸タイトルのフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType、sourceData: Range、seriesBy?: Excel。ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|新しいグラフを作成します。|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|ワークシート上のグラフの数を返します。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|グラフ名を使用してグラフを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|コレクション内での位置を基にグラフを取得します。|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|現在のグラフのデータ ラベルの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|グラフ データ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|塗りつぶしとフォントの書式設定を含むグラフ データ ラベルの形式を指定します。|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|データ ラベルの位置を表す値。|
||[区切り記号](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|データ ラベルのバブル サイズが表示される場合に指定します。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|データ ラベル のカテゴリ名が表示される場合に指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|データ ラベルの凡例キーが表示される場合に指定します。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|データ ラベルの割合を表示する場合に指定します。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|データ ラベルの系列名が表示される場合に指定します。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|データ ラベルの値が表示される場合に指定します。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|グラフ要素の塗りつぶしの色をクリアします。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|テキストの色の HTML カラー コード表現 (赤を表#FF0000など)。|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|フォント名 ("Calibri"など)|
||[size](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|フォントのサイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|フォントに適用する下線の種類。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|グラフの目盛線の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|軸のグリッド線が表示される場合に指定します。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|グラフの線の書式設定を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|塗りつぶしとフォントの書式設定を含むグラフの凡例の書式設定を表します。|
||[overlay](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|グラフの凡例がグラフの本体と重なっている必要がある場合に指定します。|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|グラフ上の凡例の位置を指定します。|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|グラフの凡例が表示される場合に指定します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|グラフの凡例のフォント名、フォント サイズ、色などのフォント属性を表します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|グラフ要素の線の形式をクリアします。|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|グラフの線の色を表す HTML カラー コード。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|グラフのポイントの書式設定プロパティをカプセル化します。|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|グラフのポイントの値を返します。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|背景の書式設定情報を含むグラフの塗りつぶしの形式を表します。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|系列に含まれるグラフのポイントの数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|データ系列内の位置に基づくポイントを取得します。|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|塗りつぶしと線の書式設定を含むグラフ系列の書式設定を表します。|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|グラフ内の系列の名前を指定します。|
||[ポイント](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|系列内のすべてのポイントのコレクションを返します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|コレクションに含まれるデータ系列の数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|コレクション内の位置に基づいてデータ系列を取得します。|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|背景書式情報を含むグラフ系列の塗りつぶし形式を表します。|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|線の書式設定を表します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|塗りつぶしとフォントの書式設定を含むグラフ タイトルの書式設定を表します。|
||[overlay](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|グラフのタイトルがグラフをオーバーレイする場合に指定します。|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|グラフのタイトル テキストを指定します。|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|グラフのタイトルが目に見えて表示される場合に指定します。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|名前に関連付けられている範囲オブジェクトを返します。|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|オブジェクトの名前。|
||[type](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|名前の数式によって返される値の種類を指定します。|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|名前の数式で計算された値を表します。|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|オブジェクトが表示される場合に指定します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|その名前を `NamedItem` 使用してオブジェクトを取得します。|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|A1 スタイルの範囲参照を指定します。|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|ユーザーの言語で指定した範囲の範囲参照を表します。|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|範囲内のセルの数を指定します。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|範囲内の列の総数を指定します。|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|範囲内の最初のセルの列番号を指定します。|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|範囲に関連付けられているセルを削除します。|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|指定した範囲を包含する、最小の Range オブジェクトを取得します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|範囲に含まれる列を 1 つ取得します。|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、そのオブジェクトは列 "B:E" `getEntireColumn` を表す範囲です)。|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、そのオブジェクトは行 "4:11" `GetEntireRow` を表す範囲です)。|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|指定した範囲の長方形の交差を表す Range オブジェクトを取得します。|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|範囲内の最後のセルを取得します。|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|範囲内の最後の列を取得します。|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|範囲内の最後の行を取得します。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|範囲に含まれている行を 1 つ取得します。|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|指定したExcelの数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|範囲に含まれる行の合計数を返します。|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|範囲に含まれる最初のセルの行番号を返します。|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Excel UI で指定した範囲を選択します。|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|各セルのデータの種類を指定します。|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|指定した範囲の Raw 値を表します。|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|現在の範囲を含んでいるワークシート。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|境界線の色を表す HTML カラー コード(#RRGGBB 形式 ("FFA500"など)、または名前の付いた HTML 色 ("オレンジ色" など) です。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|罫線の特定の辺を表す定数値。|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。|
||[weight](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|範囲周辺の罫線の太さを指定します。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|コレクションに含まれる境界線オブジェクトの数。|
||[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|範囲の背景をリセットします。|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|背景の色を表す HTML カラー コード (#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("orange"など)|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|テキストの色の HTML カラー コード表現 (赤を表#FF0000など)。|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|フォントの italic 状態を指定します。|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|フォント名 ("Calibri"など)。|
||[size](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|フォントに適用する下線の種類。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|選択した範囲全体に適用する境界線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|範囲全体に定義された塗りつぶしオブジェクトを返します。|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|範囲全体に定義されたフォント オブジェクトを返します。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|指定したオブジェクトの水平方向の配置を表します。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|指定したオブジェクトの垂直方向の配置を表します。|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|オブジェクト内のExcelを折り返す値を指定します。|
|[Table](/javascript/api/excel/excel.table)|[列](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|テーブルに含まれるすべての列のコレクションを表します。|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|テーブルを削除します。|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|表のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|表の集計行に関連付けられた範囲オブジェクトを取得します。|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|指定されたブックのテーブルを一意に識別する値を返します。|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|テーブルの名前。|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|テーブルに含まれるすべての行のコレクションを表します。|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|ヘッダー行が表示される場合に指定します。|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|合計行が表示される場合に指定します。|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|表のスタイルを表す定数値。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|新しいテーブルを作成します。|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|ブックに含まれるテーブルの数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|名前または ID でテーブルを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|コレクション内の位置に基づいてテーブルを取得します。|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|テーブルから列を削除します。|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|列全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|列の集計行に関連付けられた範囲オブジェクトを取得します。|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|テーブル内の列を識別する一意のキーを返します。|
||[index](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|テーブルの列コレクション内の列のインデックス番号を返します。|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|テーブル列の名前を指定します。|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|指定した範囲の Raw 値を表します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<配列 \| \| \|<ブール>> 文字列\|\|番号、 name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|テーブルに新しい列を追加します。|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|テーブルの列数を返します。|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|名前または ID によって、列オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|コレクション内の位置に基づいて列を取得します。|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|テーブルから行を削除します。|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|行全体に関連付けられた範囲オブジェクトを返します。|
||[index](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|テーブルの行コレクション内の行のインデックス番号を返します。|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|指定した範囲の Raw 値を表します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<配列 \| \| \| \| \|<ブール>> 文字列番号, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|テーブルに 1 つ以上の行を追加します。|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|テーブルの行数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|コレクション内の位置を基に行を取得します。|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|このブックをExcelアプリケーション インスタンスを表します。|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|ブックの一部であるバインドのコレクションを表します。|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|ブックから現在選択されている 1 つの範囲を取得します。|
||[名前](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|ブックスコープの名前付きアイテム (名前付き範囲と定数) のコレクションを表します。|
||[テーブル](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|ブックに関連付けられているテーブルのコレクションを表します。|
||[ワークシート](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|ブックに関連付けられているワークシートのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Excel UI でワークシートをアクティブにします。|
||[グラフ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|ワークシートの一部であるグラフのコレクションを返します。|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|ブックからワークシートを削除します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|行番号と `Range` 列番号に基づいて 1 つのセルを含むオブジェクトを取得します。|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|アドレスまたは名前 `Range` で指定された 1 つの四角形のセル ブロックを表すオブジェクトを取得します。|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|指定されたブックのワークシートを一意に識別する値を返します。|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|0 を起点とした、ブック内のワークシートの位置。|
||[テーブル](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|ワークシートの一部になっているグラフのコレクション。|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|ワークシートの可視性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|新しいワークシートをブックに追加します。|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|ブックの、現在作業中のワークシートを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|名前または ID を使用して、ワークシート オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
