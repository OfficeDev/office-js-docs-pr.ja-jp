---
title: ExcelJavaScript API 要件セット 1.1
description: ExcelApi 1.1 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ef764de37c8f0fea49755ba69d1beda932e17bd9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154066"
---
# <a name="excel-javascript-api-requirement-set-11"></a>ExcelJavaScript API 要件セット 1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。 この要件は、Excelでサポートされている唯一の要件Excel 2016。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.1 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.1 でサポートされるすべての API の API リファレンス ドキュメントを表示するには、「Excel API in requirement [set 1.1」](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel.CalculationType)](/javascript/api/excel/excel.application#calculate_calculationType_)|Excel で現在開いているすべてのブックを再計算します。|
||[calculationMode](/javascript/api/excel/excel.application#calculationMode)|内の定数で定義されているブックで使用される計算モードを返します `Excel.CalculationMode` 。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getRange__)|バインディングによって表される範囲を返します。|
||[getTable()](/javascript/api/excel/excel.binding#getTable__)|バインドによって表されるテーブルを返します。|
||[getText()](/javascript/api/excel/excel.binding#getText__)|バインドによって表されるテキストを返します。|
||[id](/javascript/api/excel/excel.binding#id)|バインド識別子を表します。|
||[type](/javascript/api/excel/excel.binding#type)|バインドの種類を返します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getItem_id_)|ID によってバインド オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getItemAt_index_)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|コレクション内にあるバインドの数を取得します。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete__)|グラフ オブジェクトを削除します。|
||[height](/javascript/api/excel/excel.chart#height)|グラフ オブジェクトの高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chart#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[name](/javascript/api/excel/excel.chart#name)|グラフ オブジェクトの名前を指定します。|
||[軸](/javascript/api/excel/excel.chart#axes)|グラフの軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chart#dataLabels)|グラフのデータ ラベルを表します。|
||[format](/javascript/api/excel/excel.chart#format)|グラフ領域の書式設定プロパティをカプセル化します。|
||[凡例](/javascript/api/excel/excel.chart#legend)|グラフの凡例を表します。|
||[series](/javascript/api/excel/excel.chart#series)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。|
||[title](/javascript/api/excel/excel.chart#title)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。|
||[setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy)](/javascript/api/excel/excel.chart#setData_sourceData__seriesBy_)|グラフの元データをリセットします。|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setPosition_startCell__endCell_)|ワークシート上のセルを基準にしてグラフを配置します。|
||[top](/javascript/api/excel/excel.chart#top)|オブジェクトの上端から行 1 の上端までの距離 (ワークシート上) またはグラフ領域の上端 (グラフ上) をポイントで指定します。|
||[width](/javascript/api/excel/excel.chart#width)|グラフ オブジェクトの幅をポイント単位で指定します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryAxis)|グラフの項目軸を表します。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesAxis)|3-D グラフの系列軸を表します。|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueAxis)|軸の数値軸を表します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorUnit)|2 つの大きい目盛の間隔を表します。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|数値軸の最大値を表します。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|数値軸の最小値を表します。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorUnit)|2 つの小さい目盛の間隔を表します。|
||[format](/javascript/api/excel/excel.chartaxis#format)|線とフォントの書式設定を含むグラフ オブジェクトの書式設定を表します。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorGridlines)|指定した軸の主グリッド線を表すオブジェクトを返します。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorGridlines)|指定した軸の小さい枠線を表すオブジェクトを返します。|
||[title](/javascript/api/excel/excel.chartaxis#title)|軸タイトルを表します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|グラフの線の書式設定を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|グラフ軸のタイトルの書式を指定します。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|軸のタイトルを指定します。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|軸のタイトルが表示される場合に指定します。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|グラフ軸タイトル オブジェクトのグラフ軸タイトルのフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel.ChartType、sourceData: Range、seriesBy?: Excel。ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add_type__sourceData__seriesBy_)|新しいグラフを作成します。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getItem_name_)|グラフ名を使用してグラフを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getItemAt_index_)|コレクション内での位置を基にグラフを取得します。|
||[count](/javascript/api/excel/excel.chartcollection#count)|ワークシート上のグラフの数を返します。|
||[items](/javascript/api/excel/excel.chartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|現在のグラフのデータ ラベルの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|グラフ データ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|データ ラベルの位置を表す値。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|塗りつぶしとフォントの書式設定を含むグラフ データ ラベルの形式を指定します。|
||[区切り記号](/javascript/api/excel/excel.chartdatalabels#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showBubbleSize)|データ ラベルのバブル サイズが表示される場合に指定します。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showCategoryName)|データ ラベル のカテゴリ名が表示される場合に指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showLegendKey)|データ ラベルの凡例キーが表示される場合に指定します。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showPercentage)|データ ラベルの割合を表示する場合に指定します。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showSeriesName)|データ ラベルの系列名が表示される場合に指定します。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showValue)|データ ラベルの値が表示される場合に指定します。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear__)|グラフ要素の塗りつぶしの色をクリアします。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setSolidColor_color_)|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfont#color)|テキストの色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfont#name)|フォント名 ("Calibri"など)|
||[size](/javascript/api/excel/excel.chartfont#size)|フォントのサイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|フォントに適用する下線の種類。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|グラフの目盛線の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|軸のグリッド線が表示される場合に指定します。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|グラフの線の書式設定を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|グラフの凡例がグラフの本体と重なっている必要がある場合に指定します。|
||[position](/javascript/api/excel/excel.chartlegend#position)|グラフ上の凡例の位置を指定します。|
||[format](/javascript/api/excel/excel.chartlegend#format)|塗りつぶしとフォントの書式設定を含むグラフの凡例の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|グラフの凡例が表示される場合に指定します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|グラフの凡例のフォント名、フォント サイズ、色などのフォント属性を表します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear__)|グラフ要素の線の形式をクリアします。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|グラフのポイントの書式設定プロパティをカプセル化します。|
||[value](/javascript/api/excel/excel.chartpoint#value)|グラフのポイントの値を返します。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|背景の書式設定情報を含むグラフの塗りつぶしの形式を表します。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getItemAt_index_)|データ系列内の位置に基づくポイントを取得します。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|系列に含まれるグラフのポイントの数を返します。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|グラフ内の系列の名前を指定します。|
||[format](/javascript/api/excel/excel.chartseries#format)|塗りつぶしと線の書式設定を含むグラフ系列の書式設定を表します。|
||[ポイント](/javascript/api/excel/excel.chartseries#points)|系列内のすべてのポイントのコレクションを返します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getItemAt_index_)|コレクション内の位置に基づいてデータ系列を取得します。|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|コレクションに含まれるデータ系列の数を返します。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|背景書式情報を含むグラフ系列の塗りつぶし形式を表します。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|線の書式設定を表します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|グラフのタイトルがグラフをオーバーレイする場合に指定します。|
||[format](/javascript/api/excel/excel.charttitle#format)|塗りつぶしとフォントの書式設定を含むグラフ タイトルの書式設定を表します。|
||[text](/javascript/api/excel/excel.charttitle#text)|グラフのタイトル テキストを指定します。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|グラフのタイトルが目に見えて表示される場合に指定します。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getRange__)|名前に関連付けられている範囲オブジェクトを返します。|
||[name](/javascript/api/excel/excel.nameditem#name)|オブジェクトの名前。|
||[type](/javascript/api/excel/excel.nameditem#type)|名前の数式によって返される値の種類を指定します。|
||[value](/javascript/api/excel/excel.nameditem#value)|名前の数式で計算された値を表します。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|オブジェクトが表示される場合に指定します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getItem_name_)|その名前を `NamedItem` 使用してオブジェクトを取得します。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear_applyTo_)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[delete(shift: Excel.DeleteShiftDirection)](/javascript/api/excel/excel.range#delete_shift_)|範囲に関連付けられているセルを削除します。|
||[formulas](/javascript/api/excel/excel.range#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.range#formulasLocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getBoundingRect_anotherRange_)|指定した範囲を包含する、最小の Range オブジェクトを取得します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getCell_row__column_)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getColumn_column_)|範囲に含まれる列を 1 つ取得します。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getEntireColumn__)|範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、そのオブジェクトは列 `getEntireColumn` "B:E" を表す範囲です)。|
||[getEntireRow()](/javascript/api/excel/excel.range#getEntireRow__)|範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4:E11" を表す場合、そのオブジェクトは行 `GetEntireRow` "4:11" を表す範囲です)。|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersection_anotherRange_)|指定した範囲の長方形の交差を表す Range オブジェクトを取得します。|
||[getLastCell()](/javascript/api/excel/excel.range#getLastCell__)|範囲内の最後のセルを取得します。|
||[getLastColumn()](/javascript/api/excel/excel.range#getLastColumn__)|範囲内の最後の列を取得します。|
||[getLastRow()](/javascript/api/excel/excel.range#getLastRow__)|範囲内の最後の行を取得します。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_)|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getRow_row_)|範囲に含まれている行を 1 つ取得します。|
||[insert(shift: Excel.InsertShiftDirection)](/javascript/api/excel/excel.range#insert_shift_)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。|
||[numberFormat](/javascript/api/excel/excel.range#numberFormat)|指定したExcelの数値書式コードを表します。|
||[address](/javascript/api/excel/excel.range#address)|A1 スタイルの範囲参照を指定します。|
||[addressLocal](/javascript/api/excel/excel.range#addressLocal)|ユーザーの言語で指定した範囲の範囲参照を表します。|
||[cellCount](/javascript/api/excel/excel.range#cellCount)|範囲内のセルの数を指定します。|
||[columnCount](/javascript/api/excel/excel.range#columnCount)|範囲内の列の総数を指定します。|
||[columnIndex](/javascript/api/excel/excel.range#columnIndex)|範囲内の最初のセルの列番号を指定します。|
||[format](/javascript/api/excel/excel.range#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。|
||[rowCount](/javascript/api/excel/excel.range#rowCount)|範囲に含まれる行の合計数を返します。|
||[rowIndex](/javascript/api/excel/excel.range#rowIndex)|範囲に含まれる最初のセルの行番号を返します。|
||[text](/javascript/api/excel/excel.range#text)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.range#valueTypes)|各セルのデータの種類を指定します。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|現在の範囲を含んでいるワークシート。|
||[select()](/javascript/api/excel/excel.range#select__)|Excel UI で指定した範囲を選択します。|
||[values](/javascript/api/excel/excel.range#values)|指定した範囲の Raw 値を表します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|境界線の色を表す HTML カラー コード(#RRGGBB 形式 ("FFA500"など)、または名前の付いた HTML 色 ("オレンジ色" など) です。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideIndex)|罫線の特定の辺を表す定数値。|
||[style](/javascript/api/excel/excel.rangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|範囲周辺の罫線の太さを指定します。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem(index: Excel.BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getItem_index_)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getItemAt_index_)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear__)|範囲の背景をリセットします。|
||[color](/javascript/api/excel/excel.rangefill#color)|背景の色を表す HTML カラー コード (#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("orange"など)|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefont#color)|テキストの色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|フォントの italic 状態を指定します。|
||[name](/javascript/api/excel/excel.rangefont#name)|フォント名 ("Calibri"など)。|
||[size](/javascript/api/excel/excel.rangefont#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefont#underline)|フォントに適用する下線の種類。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalAlignment)|指定したオブジェクトの水平方向の配置を表します。|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。|
||[font](/javascript/api/excel/excel.rangeformat#font)|範囲全体に定義されたフォント オブジェクトを返します。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalAlignment)|指定したオブジェクトの垂直方向の配置を表します。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wrapText)|オブジェクト内のExcelを折り返す値を指定します。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete__)|テーブルを削除します。|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getDataBodyRange__)|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getHeaderRowRange__)|表のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.table#getRange__)|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange()](/javascript/api/excel/excel.table#getTotalRowRange__)|表の集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.table#name)|テーブルの名前。|
||[列](/javascript/api/excel/excel.table#columns)|テーブルに含まれるすべての列のコレクションを表します。|
||[id](/javascript/api/excel/excel.table#id)|指定されたブックのテーブルを一意に識別する値を返します。|
||[rows](/javascript/api/excel/excel.table#rows)|テーブルに含まれるすべての行のコレクションを表します。|
||[showHeaders](/javascript/api/excel/excel.table#showHeaders)|ヘッダー行が表示される場合に指定します。|
||[showTotals](/javascript/api/excel/excel.table#showTotals)|合計行が表示される場合に指定します。|
||[style](/javascript/api/excel/excel.table#style)|表のスタイルを表す定数値。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add_address__hasHeaders_)|新しいテーブルを作成します。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getItem_key_)|名前または ID でテーブルを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getItemAt_index_)|コレクション内の位置に基づいてテーブルを取得します。|
||[count](/javascript/api/excel/excel.tablecollection#count)|ブックに含まれるテーブルの数を返します。|
||[items](/javascript/api/excel/excel.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete__)|テーブルから列を削除します。|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getDataBodyRange__)|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getHeaderRowRange__)|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getRange__)|列全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#getTotalRowRange__)|列の集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|テーブル列の名前を指定します。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|テーブル内の列を識別する一意のキーを返します。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|テーブルの列コレクション内の列のインデックス番号を返します。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|指定した範囲の Raw 値を表します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<配列<ブール>> 文字列番号 \| \| 、 \| \| \| name?: string)](/javascript/api/excel/excel.tablecolumncollection#add_index__values__name_)|テーブルに新しい列を追加します。|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItem_key_)|名前または ID によって、列オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getItemAt_index_)|コレクション内の位置に基づいて列を取得します。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|テーブルの列数を返します。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete__)|テーブルから行を削除します。|
||[getRange()](/javascript/api/excel/excel.tablerow#getRange__)|行全体に関連付けられた範囲オブジェクトを返します。|
||[index](/javascript/api/excel/excel.tablerow#index)|テーブルの行コレクション内の行のインデックス番号を返します。|
||[values](/javascript/api/excel/excel.tablerow#values)|指定した範囲の Raw 値を表します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<配列<ブール>> \| \| \| 文字列 \| \| 番号)](/javascript/api/excel/excel.tablerowcollection#add_index__values_)|テーブルに 1 つ以上の行を追加します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getItemAt_index_)|コレクション内の位置を基に行を取得します。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|テーブルの行数を返します。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getSelectedRange__)|ブックから現在選択されている 1 つの範囲を取得します。|
||[application](/javascript/api/excel/excel.workbook#application)|このブックをExcelアプリケーション インスタンスを表します。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|ブックの一部であるバインドのコレクションを表します。|
||[名前](/javascript/api/excel/excel.workbook#names)|ブックスコープの名前付きアイテム (名前付き範囲と定数) のコレクションを表します。|
||[テーブル](/javascript/api/excel/excel.workbook#tables)|ブックに関連付けられているテーブルのコレクションを表します。|
||[ワークシート](/javascript/api/excel/excel.workbook#worksheets)|ブックに関連付けられているワークシートのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate__)|Excel UI でワークシートをアクティブにします。|
||[delete()](/javascript/api/excel/excel.worksheet#delete__)|ブックからワークシートを削除します。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getCell_row__column_)|行番号 `Range` と列番号に基づいて 1 つのセルを含むオブジェクトを取得します。|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getRange_address_)|アドレスまたは名前で指定された 1 つの四角形のセル ブロックを表す `Range` オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.worksheet#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheet#position)|0 を起点とした、ブック内のワークシートの位置。|
||[グラフ](/javascript/api/excel/excel.worksheet#charts)|ワークシートの一部であるグラフのコレクションを返します。|
||[id](/javascript/api/excel/excel.worksheet#id)|指定されたブックのワークシートを一意に識別する値を返します。|
||[テーブル](/javascript/api/excel/excel.worksheet#tables)|ワークシートの一部になっているグラフのコレクション。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|ワークシートの可視性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add_name_)|新しいワークシートをブックに追加します。|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getActiveWorksheet__)|ブックの、現在作業中のワークシートを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getItem_key_)|名前または ID を使用して、ワークシート オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
