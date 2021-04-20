---
title: Excel JavaScript API 要件セット1.1
description: ExcelApi 1.1 の要件セットに関する詳細。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 648013738729961a2d36897534f500dd025cab75
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996249"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要件セット1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。 Excel 2016 でサポートされている唯一の Excel 固有の要件セットです。

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.1 の Api を示します。 Excel JavaScript API 要件セット1.1 でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.1 の Excel api](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[計算 (電卓 Ationtype: Excel. 電卓 Ationtype)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Excel で現在開いているすべてのブックを再計算します。|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|バインディングによって表される範囲を返します。|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|バインドによって表されるテーブルを返します。|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|バインドによって表されるテキストを返します。|
||[id](/javascript/api/excel/excel.binding#id)|バインド識別子を表します。|
||[type](/javascript/api/excel/excel.binding#type)|バインドの種類を返します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|ID によってバインド オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|コレクション内にあるバインドの数を取得します。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|グラフ オブジェクトを削除します。|
||[height](/javascript/api/excel/excel.chart#height)|グラフオブジェクトの高さをポイント単位で指定します。|
||[left](/javascript/api/excel/excel.chart#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[name](/javascript/api/excel/excel.chart#name)|グラフオブジェクトの名前を指定します。|
||[直交](/javascript/api/excel/excel.chart#axes)|グラフの軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|グラフのデータラベルを表します。|
||[format](/javascript/api/excel/excel.chart#format)|グラフ領域の書式設定プロパティをカプセル化します。|
||[まつわる](/javascript/api/excel/excel.chart#legend)|グラフの凡例を表します。|
||[series](/javascript/api/excel/excel.chart#series)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。|
||[title](/javascript/api/excel/excel.chart#title)|指定したグラフのタイトルのテキスト、可視性、位置、書式などを指定します。|
||[setData (sourceData: Range、系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|グラフの元データをリセットします。|
||[setPosition (startCell: Range \| string, endcell?: 範囲 \| 文字列)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|ワークシート上のセルを基準にしてグラフを配置します。|
||[top](/javascript/api/excel/excel.chart#top)|オブジェクトの上端から (ワークシートの) 行1の上端またはグラフエリアの上端までの距離をポイント単位で指定します (グラフの場合)。|
||[width](/javascript/api/excel/excel.chart#width)|グラフオブジェクトの幅をポイント単位で指定します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|グラフの項目軸を表します。|
||[系列軸](/javascript/api/excel/excel.chartaxes#seriesaxis)|3 次元グラフの系列軸を表します。|
||[数値軸](/javascript/api/excel/excel.chartaxes#valueaxis)|軸の数値軸を表します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|2 つの大きい目盛の間隔を表します。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|数値軸の最大値を表します。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|数値軸の最小値を表します。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|2 つの小さい目盛の間隔を表します。|
||[format](/javascript/api/excel/excel.chartaxis#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。|
||[title](/javascript/api/excel/excel.chartaxis#title)|軸タイトルを表します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|グラフ軸要素のフォント属性 (フォント名、フォントサイズ、色など) を指定します。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|グラフの線の書式を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|グラフ軸のタイトルの書式を指定します。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|軸のタイトルを指定します。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|軸のタイトルが shape.visibile かどうかを指定します。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|グラフ軸のタイトルのフォントの属性 (フォント名、フォントサイズ、色など) を指定します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add (type: ChartType, sourceData: Range, 系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|新しいグラフを作成します。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|グラフ名を使用してグラフを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|コレクション内での位置を基にグラフを取得します。|
||[count](/javascript/api/excel/excel.chartcollection#count)|ワークシート上のグラフの数を返します。|
||[items](/javascript/api/excel/excel.chartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|現在のグラフのデータ ラベルの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|データ ラベルの位置を表す DataLabelPosition 値。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|グラフのデータラベルの書式を指定します。これには、塗りつぶしとフォントの書式設定が含まれます。|
||[記号](/javascript/api/excel/excel.chartdatalabels#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|データラベルのバブルサイズを表示するかどうかを指定します。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|データラベルのカテゴリ名を表示するかどうかを指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|データラベルの凡例マーカーを表示するかどうかを指定します。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|データラベルのパーセンテージを表示するかどうかを指定します。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|データラベルの系列名を表示するかどうかを指定します。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|データラベルの値を表示するかどうかを指定します。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|グラフ要素の塗りつぶしの色をクリアします。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfont#color)|テキストの色の HTML カラーコード表現 (#FF0000、赤を表すなど)。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfont#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|フォントのサイズ (たとえば、11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|フォントに適用する下線の種類。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|グラフの目盛線の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|軸目盛線を表示するかどうかを指定します。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|グラフの線の書式設定を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|グラフの凡例をグラフのメイン本文と重ねて表示するかどうかを指定します。|
||[position](/javascript/api/excel/excel.chartlegend#position)|グラフの凡例の位置を指定します。|
||[format](/javascript/api/excel/excel.chartlegend#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|ChartLegend を表示するかどうかを指定します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|フォント名、フォントサイズ、色などのフォント属性を表します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|グラフ要素の線の書式をクリアします。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|グラフのポイントの書式設定プロパティをカプセル化します。|
||[value](/javascript/api/excel/excel.chartpoint#value)|グラフのポイントの値を返します。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|背景の書式設定情報を含むグラフの塗りつぶしの書式を表します。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|データ系列内の位置に基づくポイントを取得します。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|系列に含まれるグラフのポイントの数を返します。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|グラフの系列の名前を指定します。|
||[format](/javascript/api/excel/excel.chartseries#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。|
||[数](/javascript/api/excel/excel.chartseries#points)|系列内のすべてのポイントのコレクションを返します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|コレクション内の位置に基づいてデータ系列を取得します。|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|コレクションに含まれるデータ系列の数を返します。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|背景の書式設定情報を含むグラフ系列の塗りつぶしの書式を表します。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|線の書式設定を表します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|グラフタイトルがグラフを重ねて表示するかどうかを指定します。|
||[format](/javascript/api/excel/excel.charttitle#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。|
||[text](/javascript/api/excel/excel.charttitle#text)|グラフのタイトルテキストを指定します。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|グラフのタイトルが shape.visibile かどうかを指定します。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|名前に関連付けられている範囲オブジェクトを返します。|
||[name](/javascript/api/excel/excel.nameditem#name)|オブジェクトの名前。|
||[type](/javascript/api/excel/excel.nameditem#type)|名前の数式によって返される値の型を指定します。|
||[value](/javascript/api/excel/excel.nameditem#value)|名前の数式で計算された値を表します。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|オブジェクトを表示するかどうかを指定します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|名前を使用して、NamedItem オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[delete (shift: DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|範囲に関連付けられているセルを削除します。|
||[formulas](/javascript/api/excel/excel.range#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[getBoundingRect (anotherRange: Range \| 文字列)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|指定した範囲を包含する、最小の Range オブジェクトを取得します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|範囲に含まれる列を 1 つ取得します。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4: E11" を表している場合 `getEntireColumn` は、"B: E" という列を表す範囲)。|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4: E11" を表している場合 `GetEntireRow` は、行 "4:11" を表す範囲になります)。|
||[getIntersection セクション (anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|指定した範囲の長方形の交差を表す Range オブジェクトを取得します。|
||[getLastCell ()](/javascript/api/excel/excel.range#getlastcell--)|範囲内の最後のセルを取得します。|
||[getLastColumn ()](/javascript/api/excel/excel.range#getlastcolumn--)|範囲内の最後の列を取得します。|
||[getLastRow ()](/javascript/api/excel/excel.range#getlastrow--)|範囲内の最後の行を取得します。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|範囲に含まれている行を 1 つ取得します。|
||[insert (shift: InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|指定された範囲の Excel の数値書式コードを表します。|
||[address](/javascript/api/excel/excel.range#address)|セル範囲の参照を A1 形式で指定します。|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|ユーザーの言語で指定された範囲の範囲参照を指定します。|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|範囲内のセルの数を指定します。|
||[columnCount](/javascript/api/excel/excel.range#columncount)|範囲内の列の合計数を指定します。|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|範囲内の最初のセルの列番号を指定します。|
||[format](/javascript/api/excel/excel.range#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|範囲に含まれる行の合計数を返します。|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|範囲に含まれる最初のセルの行番号を返します。|
||[text](/javascript/api/excel/excel.range#text)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|各セルのデータの種類を指定します。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|現在の範囲を含んでいるワークシート。|
||[select()](/javascript/api/excel/excel.range#select--)|Excel UI で指定した範囲を選択します。|
||[values](/javascript/api/excel/excel.range#values)|指定した範囲の Raw 値を表します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|境界線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|罫線の特定の辺を表す定数値。|
||[style](/javascript/api/excel/excel.rangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|範囲周辺の罫線の太さを指定します。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|範囲の背景をリセットします。|
||[color](/javascript/api/excel/excel.rangefill#color)|背景の色を表す HTML カラーコード ("FFA500" など)、または名前付き #RRGGBB の HTML の色 (たとえば、"オレンジ")。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefont#color)|テキストの色の HTML カラーコード表現 (#FF0000、赤を表すなど)。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|フォントの斜体の状態を指定します。|
||[name](/javascript/api/excel/excel.rangefont#name)|フォント名 ("Calibri" など)。|
||[size](/javascript/api/excel/excel.rangefont#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefont#underline)|フォントに適用する下線の種類。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。|
||[font](/javascript/api/excel/excel.rangeformat#font)|範囲全体に定義されたフォント オブジェクトを返します。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Excel でオブジェクトのテキストを折り返すかどうかを指定します。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|テーブルを削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.table#getdatabodyrange--)|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|テーブルのヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.table#gettotalrowrange--)|テーブルの集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.table#name)|テーブルの名前。|
||[列](/javascript/api/excel/excel.table#columns)|テーブルに含まれるすべての列のコレクションを表します。|
||[id](/javascript/api/excel/excel.table#id)|指定されたブックのテーブルを一意に識別する値を返します。|
||[rows](/javascript/api/excel/excel.table#rows)|テーブルに含まれるすべての行のコレクションを表します。|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|ヘッダー行を表示するかどうかを指定します。|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|合計行を表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.table#style)|テーブル スタイルを表す定数値。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|新しいテーブルを作成します。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|コレクション内の位置に基づいてテーブルを取得します。|
||[count](/javascript/api/excel/excel.tablecollection#count)|ブックに含まれるテーブルの数を返します。|
||[items](/javascript/api/excel/excel.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|テーブルから列を削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|列全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|列の集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|テーブルの列の名前を指定します。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|テーブル内の列を識別する一意のキーを返します。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|テーブルの列コレクション内の列のインデックス番号を返します。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|指定した範囲の Raw 値を表します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index?: number, values?: Array<Array<boolean 文字列番号 \| \|>> \| boolean \| 文字列 \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|テーブルに新しい列を追加します。|
||[getItem (key: number \| 文字列)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|名前または ID を使用して列オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|コレクション内の位置に基づいて列を取得します。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|テーブルの列数を返します。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|テーブルから行を削除します。|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|行全体に関連付けられた範囲オブジェクトを返します。|
||[index](/javascript/api/excel/excel.tablerow#index)|テーブルの行コレクション内の行のインデックス番号を返します。|
||[values](/javascript/api/excel/excel.tablerow#values)|指定した範囲の Raw 値を表します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index?: number, values?: Array<Array<boolean \| 文字列 \| 番号>> \| boolean \| 文字列 \| 番号)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|テーブルに 1 つ以上の行を追加します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|コレクション内の位置を基に行を取得します。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|テーブルの行数を返します。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|ブックから現在選択されている1つのセル範囲を取得します。|
||[application](/javascript/api/excel/excel.workbook#application)|このブックを含む Excel アプリケーションインスタンスを表します。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|ブックの一部であるバインドのコレクションを表します。|
||[姓名](/javascript/api/excel/excel.workbook#names)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。|
||[テーブル](/javascript/api/excel/excel.workbook#tables)|ブックに関連付けられているテーブルのコレクションを表します。|
||[what-if](/javascript/api/excel/excel.workbook#worksheets)|ブックに関連付けられているワークシートのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Excel UI でワークシートをアクティブにします。|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|ブックからワークシートを削除します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。|
||[getRange (address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|アドレスまたは名前で指定された、セルの単一の四角形のブロックを表す range オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.worksheet#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheet#position)|0 を起点とした、ブック内のワークシートの位置。|
||[管理図](/javascript/api/excel/excel.worksheet#charts)|ワークシートの一部であるグラフのコレクションを返します。|
||[id](/javascript/api/excel/excel.worksheet#id)|指定されたブックのワークシートを一意に識別する値を返します。|
||[テーブル](/javascript/api/excel/excel.worksheet#tables)|ワークシートの一部になっているグラフのコレクション。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|ワークシートの可視性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add (name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|新しいワークシートをブックに追加します。|
||[getActiveWorksheet ()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|ブックの、現在作業中のワークシートを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|名前または ID を使用して、ワークシート オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
