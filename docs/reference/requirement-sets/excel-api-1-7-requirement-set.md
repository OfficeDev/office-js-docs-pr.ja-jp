---
title: ExcelJavaScript API 要件セット 1.7
description: ExcelApi 1.7 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 67f30fd61e3065f8d7d193668c6f79fd09debf2f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350212"
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

次の表に、JavaScript API 要件セット 1.7 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.7 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|グラフの種類を指定します。|
||[id](/javascript/api/excel/excel.chart#id)|グラフの一意の ID。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|すべてのフィールド ボタンを 1 つのウィンドウに表示するかどうかをピボットグラフ。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|色、線のスタイル、太さなど、グラフ領域の罫線の形式を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel.ChartAxisType、group?: Excel。ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|種類とグループで識別された特定の軸を返します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|指定したカテゴリ軸の基本単位を指定します。|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|カテゴリ軸の種類を指定します。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|軸の表示単位を表します。|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|対数スケールを使用する場合の対数の基数を指定します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|指定した軸の目盛の種類を指定します。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|CategoryType プロパティが TimeScale に設定されている場合に、カテゴリ軸のメジャー単位スケール値を指定します。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|指定した軸の目盛りの種類を指定します。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|CategoryType プロパティが TimeScale に設定されている場合に、カテゴリ軸のマイナー単位スケール値を指定します。|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|指定した軸のグループを指定します。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|ユーザー設定の軸表示単位の値を指定します。|
||[height](/javascript/api/excel/excel.chartaxis#height)|グラフ軸の高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chartaxis#left)|軸の左端からグラフ領域の左側までの距離をポイントで指定します。|
||[top](/javascript/api/excel/excel.chartaxis#top)|軸の上端からグラフ領域の上端までの距離をポイントで指定します。|
||[type](/javascript/api/excel/excel.chartaxis#type)|軸の種類を指定します。|
||[width](/javascript/api/excel/excel.chartaxis#width)|グラフ軸の幅をポイント単位で指定します。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|最後から最初Excelデータ ポイントをプロットする方法を指定します。|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|値軸のスケールの種類を指定します。|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|指定した軸のすべてのカテゴリ名を設定します。|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|軸の表示単位をカスタム値に設定します。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|軸表示単位ラベルが表示される場合に指定します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|指定された軸の目盛ラベルの位置を指定します。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|目盛ラベル間のカテゴリまたは系列の数を指定します。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|目盛の間のカテゴリまたは系列の数を指定します。|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|軸が表示される場合に指定します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|罫線のスタイルを表します。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|データ ラベルの位置を表す DataLabelPosition 値。|
||[区切り記号](/javascript/api/excel/excel.chartdatalabel#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|データ ラベルのバブル サイズが表示される場合に指定します。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|データ ラベル のカテゴリ名が表示される場合に指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|データ ラベルの凡例キーが表示される場合に指定します。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|データ ラベルの割合を表示する場合に指定します。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|データ ラベルの系列名が表示される場合に指定します。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|データ ラベルの値が表示される場合に指定します。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|フォント名、フォント サイズ、色などのフォント属性を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|グラフ上の凡例の高さをポイントで指定します。|
||[left](/javascript/api/excel/excel.chartlegend#left)|グラフ上の凡例の左をポイントで指定します。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|凡例に含まれる凡例エントリのコレクションを表します。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|凡例にグラフに影が付く場合を指定します。|
||[top](/javascript/api/excel/excel.chartlegend#top)|グラフの凡例の上部を指定します。|
||[width](/javascript/api/excel/excel.chartlegend#width)|グラフ上の凡例の幅をポイント単位で指定します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|コレクションに含まれる凡例エントリの数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|指定されたインデックスに位置する凡例エントリを返します。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|線のスタイルを表します。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|線の太さ (ポイント数) を表します。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|データ ポイントにデータ ラベルが含されているかどうかを表します。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|データ ポイントのマーカー背景色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|データ ポイントのマーカーの前景色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|データ ポイントのマーカー スタイルを表します。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|グラフ データ ポイントのデータ ラベルを返します。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|色、スタイル、および重み情報を含むグラフ データ ポイントの罫線の形式を表します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|グラフ系列の種類を表します。|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|グラフ系列を削除します。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|グラフ系列のドーナツの穴の大きさを表します。|
||[フィルター処理](/javascript/api/excel/excel.chartseries#filtered)|系列をフィルター処理する場合に指定します。|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|グラフ系列間に設けられる間隔を表します。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|系列にデータ ラベルが含む場合を指定します。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|グラフ系列のマーカーの背景色を指定します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|グラフ系列のマーカーの前景色を指定します。|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|グラフ系列のマーカー サイズを指定します。|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|グラフ系列のマーカー スタイルを指定します。|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|グラフ グループ内のグラフ系列のプロット順序を指定します。|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|系列内の傾向線のコレクション。|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|グラフ系列のバブル サイズを設定します。|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|グラフ系列の値を設定します。|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|グラフ系列の X 軸の値を設定します。|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|系列に影が付く場合を指定します。|
||[スムーズ](/javascript/api/excel/excel.chartseries#smooth)|系列が滑らかな場合に指定します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|コレクションに新しい系列を追加します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|グラフタイトルの部分文字列を取得します。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|グラフタイトルの水平方向の配置を指定します。|
||[left](/javascript/api/excel/excel.charttitle#left)|グラフ タイトルの左端からグラフ領域の左端までの距離をポイントで指定します。|
||[position](/javascript/api/excel/excel.charttitle#position)|グラフ タイトルの位置を表します。|
||[height](/javascript/api/excel/excel.charttitle#height)|グラフ タイトルの高さ (ポイント数) を返します。|
||[width](/javascript/api/excel/excel.charttitle#width)|グラフ タイトルの幅をポイント単位で指定します。|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|A1 スタイルの表記法を使用するグラフ タイトルの数式を表す文字列値を設定します。|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|グラフ タイトルのテキストの向きを指定します。|
||[top](/javascript/api/excel/excel.charttitle#top)|グラフ タイトルの上端からグラフ領域の上端までの距離をポイントで指定します。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|グラフ タイトルの垂直方向の配置を指定します。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|色、線のスタイル、太さなど、グラフタイトルの罫線の形式を表します。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|trendline オブジェクトを削除します。|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|近似曲線の切片の値を表します。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|グラフの傾向線の期間を表します。|
||[name](/javascript/api/excel/excel.charttrendline#name)|近似曲線の名前を表します。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|グラフの傾向線の順序を表します。|
||[format](/javascript/api/excel/excel.charttrendline#format)|グラフの近似曲線の書式設定を表します。|
||[type](/javascript/api/excel/excel.charttrendline#type)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel.ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|近似曲線のコレクションに新しい近似曲線を追加します。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|コレクションに含まれる近似曲線の数を返します。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|インデックス (項目配列内の挿入順序) に基づいて trendline オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|グラフの線の書式設定を表します。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.customproperty#key)|カスタム プロパティのキー。|
||[type](/javascript/api/excel/excel.customproperty#type)|カスタム プロパティに使用される値の種類。|
||[value](/javascript/api/excel/excel.customproperty#value)|カスタム プロパティの値を指定します。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|コレクションに含まれるすべてのデータ接続を更新します。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|ブックの作成者。|
||[category](/javascript/api/excel/excel.documentproperties#category)|ブックのカテゴリ。|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|ブックのコメント。|
||[company](/javascript/api/excel/excel.documentproperties#company)|ブックの会社。|
||[キーワード](/javascript/api/excel/excel.documentproperties#keywords)|ブックのキーワード。|
||[上司](/javascript/api/excel/excel.documentproperties#manager)|ブックのマネージャー。|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|ブックの作成日を取得します。|
||[カスタム](/javascript/api/excel/excel.documentproperties#custom)|ブックのカスタム プロパティのコレクションを取得します。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|ブックの最後の作成者を取得します。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|ブックのリビジョン番号を取得します。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|ブックの件名。|
||[title](/javascript/api/excel/excel.documentproperties#title)|ブックのタイトル。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|名前付きアイテムの数式。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|名前付きアイテムの値と型を含むオブジェクトを返します。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|名前付きアイテム配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|現在の Range オブジェクトと左上のセルが同じで、指定した数の行と列を含む Range オブジェクトを取得します。|
||[getImage()](/javascript/api/excel/excel.range#getimage--)|範囲を base64 エンコードされた png イメージとしてレンダリングします。|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|指定された範囲の左上のセルを囲む領域を表す Range オブジェクトを返します。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|ユーザー Excelの言語設定に基づいて、指定した範囲の数値書式コードを表します。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|現在の範囲が列全体であるかどうかを表します。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|現在の範囲が行全体であるかどうかを表します。|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|アクティブ セルに多数の値が含まれる場合、そのセルのカードを表示します。|
||[style](/javascript/api/excel/excel.range#style)|現在の範囲のスタイルを表します。|
|[範囲の形式](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|範囲内のすべてのセルのテキストの向き。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Range オブジェクトの列幅がシートの標準幅と等しい場合に指定します。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|ハイパーリンクの URL ターゲットを表します。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|ハイパーリンクのドキュメント参照ターゲットを表します。|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screentip)|ハイパーリンクの上にカーソルを合わせると表示される文字列を表します。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|該当する範囲内の左上端のセルに表示される文字列を表します。|
|[スタイル](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|このスタイルを削除します。|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|ワークシートを保護するときに数式を非表示に設定する場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|スタイルでの水平方向の配置を表します。|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|スタイルに AutoIndent プロパティ、HorizontalAlignment プロパティ、VerticalAlignment プロパティ、WrapText プロパティ、IndentLevel プロパティ、TextOrientation プロパティが含まれる場合に指定します。|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|スタイルに Color、ColorIndex、LineStyle、および Weight の罫線プロパティが含まれる場合に指定します。|
||[includeFont](/javascript/api/excel/excel.style#includefont)|スタイルに、背景、太字、色、ColorIndex、FontStyle、斜体、名前、サイズ、取り消し線、下付き文字、下線のフォント プロパティが含まれる場合に指定します。|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|スタイルに NumberFormat プロパティが含まれる場合に指定します。|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|スタイルに Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、および PatternColorIndex の内部プロパティが含まれる場合に指定します。|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|スタイルに FormulaHidden プロパティと Locked protection プロパティが含まれる場合に指定します。|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.style#locked)|ワークシートが保護されているときにオブジェクトがロックされる場合に指定します。|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|スタイルで適用される読み上げ順序。|
||[borders](/javascript/api/excel/excel.style#borders)|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。|
||[builtIn](/javascript/api/excel/excel.style#builtin)|スタイルが組み込みのスタイルである場合に指定します。|
||[fill](/javascript/api/excel/excel.style#fill)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.style#font)|スタイルのフォントを表す Font オブジェクト。|
||[name](/javascript/api/excel/excel.style#name)|スタイルの名前。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|使用可能な列の幅に収まるテキストを自動的に縮小する場合に指定します。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|スタイルの垂直方向の配置を指定します。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|オブジェクト内のExcelを折り返す値を指定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|コレクションに新しいスタイルを追加します。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|名前に基づいてスタイルを取得します。|
||[items](/javascript/api/excel/excel.stylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|セル内のデータが特定のテーブルで変更された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|特定のテーブルで選択範囲が変更された場合に発生します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|特定のワークシート上のテーブル内で変更されたエリアを表すアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|イベントのソースを取得します。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|データが変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|ブックまたはワークシート内の任意のテーブルでデータが変更された場合に発生します。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|特定のワークシート上のテーブル内で選択されたエリアを表す範囲のアドレスを取得します。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|IsInsideTable が false の場合、選択範囲がテーブル内にある場合、アドレスは使用できません。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|選択範囲が変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|選択範囲が変更されたワークシートの ID を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|ブックで現在アクティブなセルを取得します。|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|ブック内のすべてのデータ接続を表します。|
||[name](/javascript/api/excel/excel.workbook#name)|ブックの名前を取得します。|
||[プロパティ](/javascript/api/excel/excel.workbook#properties)|ブックのプロパティを取得します。|
||[protection](/javascript/api/excel/excel.workbook#protection)|ブックの保護オブジェクトを返します。|
||[スタイル](/javascript/api/excel/excel.workbook#styles)|ブックに関連付けられているスタイルのコレクションを表します。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|ブックを保護します。|
||[保護](/javascript/api/excel/excel.workbookprotection#protected)|ブックが保護される場合に指定します。|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|ブックの保護を解除します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel.WorksheetPositionType、 relativeTo?: Excel。ワークシート)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|ワークシートをコピーし、指定した位置に配置します。|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、Range オブジェクトを取得します。|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|ワークシートの固定されたウィンドウを操作するために使用できるオブジェクトを取得します。|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|ワークシートがアクティブ化されると発生します。|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|特定のワークシートでデータが変更された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|ワークシートが非アクティブ化された場合に発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|特定のワークシートで選択範囲が変更された場合に発生します。|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|ワークシート内のすべての列の標準 (既定) の幅を指定します。|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|ワークシートのタブの色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|アクティブにされたワークシートの ID を取得します。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|ブックに追加されたワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|ブック内のワークシートがアクティブ化されると発生します。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|新しいワークシートがブックに追加された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|ブック内のワークシートが非アクティブ化された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|ワークシートがブックから削除された場合に発生します。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|非アクティブにされたワークシートの ID を取得します。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|ブックから削除されたワークシートの ID を取得します。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|アクティブなワークシート ビューに固定セルを設定します。|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|ワークシートの最初の列 (複数可) を所定の場所に固定します。|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|ワークシートの最初の行 (複数可) を所定の場所に固定します。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|ワークシートからすべての固定ウィンドウを削除します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|ワークシートの保護を解除します。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|オブジェクトの編集を可能にするワークシート保護オプションを表します。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|シナリオの編集を可能にするワークシート保護オプションを表します。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|選択モードのワークシート保護オプションを表します。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|特定のワークシートで選択されたエリアを表す範囲のアドレスを取得します。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|選択範囲が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
