---
title: Excel JavaScript API 要件セット1.1
description: ExcelApi 1.1 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 90d7ee7cef2e8c48e458b2e14893ba9c13c68a30
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940788"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要件セット1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。 Excel 2016 でサポートされている唯一の Excel 固有の要件セットです。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[計算 (電卓 Ationtype: Excel. 電卓 Ationtype)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Excel で現在開いているすべてのブックを再計算します。|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。 使用可能な値`Automatic`は次のとおりです。 Excel では、再計算が制御されます。`AutomaticExceptTables`、Excel は再計算を制御しますが、テーブル内の変更は無視します。`Manual`、ユーザーが要求すると、計算が行われます。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|バインドによって表されるテキストを返します。 バインドが正しい型ではない場合、エラーがスローされます。|
||[id](/javascript/api/excel/excel.binding#id)|バインド識別子を表します。 読み取り専用。|
||[type](/javascript/api/excel/excel.binding#type)|バインドの種類を返します。 詳細については、「Excel. BindingType」を参照してください。 読み取り専用です。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|ID によってバインド オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|コレクション内にあるバインドの数を取得します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|グラフ オブジェクトを削除します。|
||[height](/javascript/api/excel/excel.chart#height)|グラフ オブジェクトの高さをポイント単位で表します。|
||[left](/javascript/api/excel/excel.chart#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[name](/javascript/api/excel/excel.chart#name)|グラフ オブジェクトの名前を表します。|
||[直交](/javascript/api/excel/excel.chart#axes)|グラフの軸を表します。 読み取り専用です。|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|グラフのデータラベルを表します。 読み取り専用です。|
||[format](/javascript/api/excel/excel.chart#format)|グラフ領域の書式設定プロパティをカプセル化します。 読み取り専用です。|
||[まつわる](/javascript/api/excel/excel.chart#legend)|グラフの凡例を表します。 読み取り専用です。|
||[series](/javascript/api/excel/excel.chart#series)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。 読み取り専用です。|
||[title](/javascript/api/excel/excel.chart#title)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。 読み取り専用です。|
||[setData (sourceData: Range、系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|グラフの元データをリセットします。|
||[setPosition (startCell: Range \| String, endcell?: 範囲\|文字列)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|ワークシート上のセルを基準にしてグラフを配置します。|
||[top](/javascript/api/excel/excel.chart#top)|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.chart#width)|グラフ オブジェクトの幅をポイント単位で表します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|グラフの項目軸を表します。 読み取り専用です。|
||[系列軸](/javascript/api/excel/excel.chartaxes#seriesaxis)|3 次元グラフの系列軸を表します。 読み取り専用です。|
||[数値軸](/javascript/api/excel/excel.chartaxes#valueaxis)|軸の数値軸を表します。 読み取り専用です。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|数値軸の最小値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|2 つの小さい目盛の間隔を表します。 数値の値または空の文字列を設定できます (軸の値が自動の場合)。 戻り値は常に数値です。|
||[format](/javascript/api/excel/excel.chartaxis#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。 読み取り専用です。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。 値の取得のみ可能です。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。 読み取り専用です。|
||[title](/javascript/api/excel/excel.chartaxis#title)|軸タイトルを表します。 読み取り専用です。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|グラフの線の書式設定を表します。 読み取り専用です。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|グラフ軸のタイトルの書式設定を表します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|軸タイトルを表します。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|軸のタイトルの表示/非表示を指定するブール型の値です。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|グラフの軸タイトルのオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add (type: ChartType, sourceData: Range, 系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|新しいグラフを作成します。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|グラフ名を使用してグラフを取得します。 同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|コレクション内での位置を基にグラフを取得します。|
||[count](/javascript/api/excel/excel.chartcollection#count)|ワークシート上のグラフの数を返します。 値の取得のみ可能です。|
||[items](/javascript/api/excel/excel.chartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|現在のグラフのデータ ラベルの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。 読み取り専用です。|
||[記号](/javascript/api/excel/excel.chartdatalabels#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|グラフ要素の塗りつぶしの色をクリアします。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfont#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|フォント サイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|フォントに適用する下線の種類。 詳細については、「Excel のグラフ」を参照してください。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|グラフの目盛線の書式設定を表します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|軸の目盛線を表示するか非表示にするかを表すブール型の値。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|グラフの線の書式設定を表します。 読み取り専用です。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|
||[position](/javascript/api/excel/excel.chartlegend#position)|グラフの凡例の位置を表します。 詳細については、「ChartLegendPosition」を参照してください。|
||[format](/javascript/api/excel/excel.chartlegend#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|グラフの凡例のフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|グラフ要素の線の書式をクリアします。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|グラフのポイントの書式設定プロパティをカプセル化します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.chartpoint#value)|グラフのポイントの値を返します。 読み取り専用です。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|背景の書式設定情報を含むグラフの塗りつぶしの書式を表します。 読み取り専用です。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|データ系列内の位置に基づくポイントを取得します。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|系列内にあるグラフのポイントの数を取得します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|グラフのデータ系列の名前を表します。|
||[format](/javascript/api/excel/excel.chartseries#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。 読み取り専用です。|
||[数](/javascript/api/excel/excel.chartseries#points)|データ系列にあるすべてのポイントのコレクションを返します。 読み取り専用です。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|コレクション内の位置に基づいてデータ系列を取得します。|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|コレクション内にあるデータ系列の数を取得します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|背景の書式設定情報を含むグラフ系列の塗りつぶしの書式を表します。 読み取り専用です。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|線の書式設定を表します。 読み取り専用です。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
||[format](/javascript/api/excel/excel.charttitle#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.charttitle#text)|グラフのタイトルのテキストを表します。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 値の取得のみ可能です。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|名前に関連付けられている範囲オブジェクトを返します。 名前付きアイテムの型が範囲でない場合、エラーをスローします。|
||[name](/javascript/api/excel/excel.nameditem#name)|オブジェクトの名前。 読み取り専用です。|
||[type](/javascript/api/excel/excel.nameditem#type)|名前の数式によって返される値の型を示します。 詳細については、「Excel. NamedItemType」を参照してください。 読み取り専用です。|
||[value](/javascript/api/excel/excel.nameditem#value)|名前の数式で計算された値を表します。 名前付き範囲の場合は範囲のアドレスを返します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|オブジェクトを表示するかどうかを指定します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|名前を使用して、NamedItem オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[delete (shift: DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|範囲に関連付けられているセルを削除します。|
||[formulas](/javascript/api/excel/excel.range#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[getBoundingRect (anotherRange: Range \|文字列)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|指定した範囲を包含する、最小の Range オブジェクトを取得します。 たとえば、"B2:C5" と "D10:E15" の GetBoundingRect は、"B2:E15" になります。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。 ワークシートのグリッド内に収まるセルは、親の範囲の境界の外側にある場合があります。 返されるセルは、範囲の左上のセルを基準に配置されます。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|範囲に含まれる列を 1 つ取得します。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|範囲の列全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4: E11" を表している`getEntireColumn`場合は、"B: E" という列を表す範囲)。|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|範囲の行全体を表すオブジェクトを取得します (たとえば、現在の範囲がセル "B4: E11" を表している`GetEntireRow`場合は、行 "4:11" を表す範囲になります)。|
||[getIntersection セクション (anotherRange: \| Range string)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|指定した範囲の長方形の交差を表す Range オブジェクトを取得します。|
||[getLastCell ()](/javascript/api/excel/excel.range#getlastcell--)|範囲内の最後のセルを取得します。たとえば、"B2:D5" の最後のセルは "D5" になります。|
||[getLastColumn ()](/javascript/api/excel/excel.range#getlastcolumn--)|範囲内の最後の列を取得します。たとえば、"B2:D5" の最後の列は "D2:D5" になります。|
||[getLastRow ()](/javascript/api/excel/excel.range#getlastrow--)|範囲内の最後の行を取得します。たとえば、"B2:D5" の最後の行は "B5:D5" になります。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|指定した範囲からのオフセットで範囲を表すオブジェクトを取得します。返される範囲のディメンションは、この範囲と一致します。結果の範囲がワークシートのグリッドの境界線の外にはみ出る場合は、エラーがスローされます。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|範囲に含まれている行を 1 つ取得します。|
||[insert (shift: InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|指定された範囲の Excel の数値書式コードを表します。|
||[address](/javascript/api/excel/excel.range#address)|A1 スタイルの範囲参照を表します。 Address 値にはシート参照が含まれます (例: "Sheet1!A1: B4 ") 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|ユーザーの言語で指定された範囲の範囲参照を表します。 読み取り専用です。|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|範囲に含まれるセルの数。 セルの数が 2^31-1 (2,147,483,647) を超えると、この API は -1 を返します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.range#columncount)|範囲に含まれる列の合計数を表します。 読み取り専用です。|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|範囲に含まれる最初のセルの列番号を表します。 0 を起点とする番号になります。 読み取り専用。|
||[format](/javascript/api/excel/excel.range#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。 読み取り専用です。|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|範囲に含まれる行の合計数を返します。 読み取り専用です。|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|範囲に含まれる最初のセルの行番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[text](/javascript/api/excel/excel.range#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|現在の範囲を含んでいるワークシート。 読み取り専用です。|
||[select()](/javascript/api/excel/excel.range#select--)|Excel UI で指定した範囲を選択します。|
||[values](/javascript/api/excel/excel.range#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel BorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|範囲周辺の罫線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。 読み取り専用です。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|範囲の背景をリセットします。|
||[color](/javascript/api/excel/excel.rangefill#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.rangefont#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.rangefont#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefont#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Range過小な Linestyle」を参照してください。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.rangeformat#font)|範囲全体に定義されたフォント オブジェクトを返します。 読み取り専用です。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|オブジェクト内のテキストを Excel でラップするかどうかを表します。 null 値は、範囲全体に一様なラップ設定がないことを表します。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|テーブルを削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.table#getdatabodyrange--)|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|テーブルのヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.table#gettotalrowrange--)|テーブルの集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.table#name)|テーブルの名前。|
||[列](/javascript/api/excel/excel.table#columns)|テーブルに含まれるすべての列のコレクションを表します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.table#id)|指定されたブックのテーブルを一意に識別する値を返します。識別子の値は、テーブルの名前が変更された場合も変わりません。読み取り専用です。|
||[rows](/javascript/api/excel/excel.table#rows)|テーブルに含まれるすべての行のコレクションを表します。 読み取り専用です。|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|ヘッダー行を表示するかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|集計行を表示するかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.table#style)|テーブル スタイルを表す定数値。使用可能な値は次のとおりです。TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address: Range \| String, hasheaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|コレクション内の位置に基づいてテーブルを取得します。|
||[count](/javascript/api/excel/excel.tablecollection#count)|ブックに含まれるテーブルの数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|テーブルから列を削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|列全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|列の集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|テーブル列の名前を表します。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|テーブル内の列を識別する一意のキーを返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|テーブルの列コレクション内の列のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index?: number, values?: Array<Array<boolean \| \|文字列番号>> \| boolean \|文字列\| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|テーブルに新しい列を追加します。|
||[getItem (key: number \|文字列)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|名前または ID を使用して列オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|コレクション内の位置に基づいて列を取得します。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|テーブルの列数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|テーブルから行を削除します。|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|行全体に関連付けられた範囲オブジェクトを返します。|
||[index](/javascript/api/excel/excel.tablerow#index)|テーブルの行コレクション内の行のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[values](/javascript/api/excel/excel.tablerow#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index?: number, values?: Array<Array<boolean \|文字列\|番号>> \| boolean \|文字列\|番号)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|テーブルに 1 つ以上の行を追加します。 戻りオブジェクトは新しく追加された行の先頭になります。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|コレクション内の位置を基に行を取得します。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|テーブルの行数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|ブックから現在選択されている1つのセル範囲を取得します。 複数の範囲が選択されている場合、このメソッドはエラーをスローします。|
||[application](/javascript/api/excel/excel.workbook#application)|このブックを含む Excel アプリケーションインスタンスを表します。 読み取り専用です。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|ブックの一部であるバインドのコレクションを表します。 読み取り専用。|
||[names](/javascript/api/excel/excel.workbook#names)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。 読み取り専用です。|
||[テーブル](/javascript/api/excel/excel.workbook#tables)|ブックに関連付けられているテーブルのコレクションを表します。 読み取り専用です。|
||[what-if](/javascript/api/excel/excel.workbook#worksheets)|ブックに関連付けられているワークシートのコレクションを表します。 読み取り専用です。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Excel UI でワークシートをアクティブにします。|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|ブックからワークシートを削除します。 ワークシートの可視性が "非常に非表示" に設定されている場合は、削除操作が一般の例外によって失敗することに注意してください。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。 ワークシートのグリッド内に収まるセルは、親の範囲の境界の外側にある場合があります。|
||[getRange (address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|アドレスまたは名前で指定された、セルの単一の四角形のブロックを表す range オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.worksheet#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheet#position)|0 を起点とした、ブック内のワークシートの位置。|
||[管理図](/javascript/api/excel/excel.worksheet#charts)|ワークシートの一部になっているグラフのコレクションを返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.worksheet#id)|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|
||[テーブル](/javascript/api/excel/excel.worksheet#tables)|ワークシートの一部になっているグラフのコレクション。 読み取り専用です。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|ワークシートの可視性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add (name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。|
||[getActiveWorksheet ()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|ブックの、現在作業中のワークシートを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|名前または ID を使用して、ワークシート オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
