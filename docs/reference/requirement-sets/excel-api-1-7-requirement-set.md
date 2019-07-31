---
title: Excel JavaScript API 要件セット1.7
description: ExcelApi 1.7 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7a4d0dde78a290de61fb6edc1966ea66dc2ba3b0
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940697"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 の新機能

Excel JavaScript API 要件セット 1.7 の機能には、グラフ、イベント、ワークシート、範囲、ドキュメント プロパティ、名前付きアイテム、保護のオプションとスタイルに対応する API が含まれます。

## <a name="customize-charts"></a>グラフのカスタマイズ

新しいグラフ API を使用して実行できる操作には、他の種類のグラフの作成、グラフへのデータ系列の追加、グラフ タイトルの設定、軸タイトルの追加、表示単位の追加、移動平均を使用した近似曲線の追加、線形近似曲線への変更などがあります。 以下はその例です。

* グラフの軸: グラフ内の軸の単位、ラベル、タイトルを取得、設定、書式設定する。
* グラフのデータ系列: グラフ内のデータ系列を追加、設定、削除する。  データ系列マーカー、プロット順序、サイジングを変更する。
* グラフの近似曲線: グラフ内の近似曲線を追加、取得、書式設定する。
* グラフの凡例: グラフ内の凡例のフォントを書式設定する。
* グラフ データ ポイント要素: データ要素の色を設定する。
* グラフ タイトルのサブ文字列: グラフ タイトルのサブ文字列を取得、設定する。
* グラフの種類: 他の種類のグラフを作成するオプションを使用する。

## <a name="events"></a>イベント

Excel イベント API には各種のイベント ハンドラーが用意されています。これらのハンドラーを使用することで、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できます。 実行する関数は、目的のシナリオに必要な処理を行うように設計できます。 現在利用可能なイベントのリストについては、[Excel JavaScript API を使用してイベントを操作する](/office/dev/add-ins/excel/excel-add-ins-events)を参照してください。

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>ワークシートと範囲の外観のカスタマイズ

新しい API を使用して、ワークシートの外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

* ワークシート内でスクロールするときに特定の行または列が常に表示されるよう、ウィンドウ枠を固定する。 たとえば、ワークシート内の最初の行にヘッダーが示される場合、その行にウィンドウ枠を固定すると、ワークシートをスクロールダウンしても列の見出しは表示されたままになります。
* ワークシートのタブの色を変更する。
* ワークシートの見出しを追加する。

範囲の外観をさまざまな方法でカスタマイズできます。たとえば、次のようなカスタマイズが可能です。

* 特定の範囲に対してセルのスタイルを設定し、その範囲内のすべてのセルに一貫した書式設定が適用されるようにする。 セルのスタイルとは、フォント、フォントのサイズ、数値形式、セルの罫線、セルの網掛けなど、文字に定義された書式設定一式を指します。 Excel の組み込みのセル スタイルのいずれかを使用することも、独自のカスタム セル スタイルを作成することもできます。
* 範囲に適用するテキストの向きを設定する。
* 特定の範囲からブック内の別の場所または外部の場所にリンクするハイパーリンクを追加または変更する。

## <a name="manage-document-properties"></a>ドキュメント プロパティの管理

ドキュメント プロパティ API を使用して、組み込みのドキュメント プロパティにアクセスできます。また、ブックの状態を格納してワークフローやビジネス ロジックを操作するためのカスタム ドキュメント プロパティを作成、管理することもできます。

## <a name="copy-worksheets"></a>ワークシートのコピー

ワークシート コピー API を使用して、ワークシートのデータと書式設定を同じブック内の新しいワークシートにコピーできます。これにより、必要となるデータの転送量を削減することができます。

## <a name="handle-ranges-with-ease"></a>範囲の操作の簡易化

各種の範囲 API を使用して、周りの領域の取得や範囲のサイズ変更など、さまざまな操作を行うことができます。 これらの API により、範囲の操作やアドレス指定などのタスクが効率化されます。

さらに、次の機能も使用できます。

* ブックとワークシートの保護オプション: これらの API を使用して、ワークシートおよびブック構造内のデータを保護する。
* 名前付きアイテムの更新: この API を使用して、名前付きアイテムを更新する。
* アクティブ セルの取得: この API を使用して、ブックのアクティブ セルを取得する。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|グラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[id](/javascript/api/excel/excel.chart#id)|グラフの一意の ID。 読み取り専用です。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|ピボットグラフにすべてのフィールド ボタンを表示するかどうかを示します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[罫線](/javascript/api/excel/excel.chartareaformat#border)|グラフエリアの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。 読み取り専用です。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (type: ChartAxisType, group?: ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|種類とグループで識別された特定の軸を返します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|指定された項目軸の基本単位を返すか設定します。|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|項目軸の種類を返すか設定します。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|軸の表示単位を表します。 詳細については、「ChartAxisDisplayUnit」を参照してください。|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|対数目盛りを使用する場合の対数の底を表します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|指定した軸の目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の目盛のスケール値を返すか設定します。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|指定した軸の補助目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の補助目盛のスケール値を返すか設定します。|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|指定した軸に対応するグループを表します。 詳細については、「ChartAxisGroup」を参照してください。 読み取り専用です。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|カスタム軸の表示単位の値を表します。 読み取り専用です。 このプロパティを設定するには、SetCustomDisplayUnit(double) メソッドを使用してください。|
||[height](/javascript/api/excel/excel.chartaxis#height)|グラフ軸の高さ (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[left](/javascript/api/excel/excel.chartaxis#left)|軸の左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[top](/javascript/api/excel/excel.chartaxis#top)|軸の上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[type](/javascript/api/excel/excel.chartaxis#type)|軸の種類を表します。 詳細については、「ChartAxisType」を参照してください。|
||[width](/javascript/api/excel/excel.chartaxis#width)|グラフ軸の幅 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Microsoft Excel でデータ ポイントを最後から最初への順でプロットするかどうかを表します。|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|数値軸のスケールの種類を表します。 詳細については、「ChartAxisScaleType」を参照してください。|
||[Setカテゴリ名 (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|指定した軸のすべてのカテゴリ名を設定します。|
||[setCustomDisplayUnit (値: number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|軸の表示単位をカスタム値に設定します。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|軸の表示単位のラベルを表示するかどうかを表します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|指定した軸での目盛ラベルの位置を示します。 詳細については、「ChartAxisTickLabelPosition」を参照してください。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|目盛ラベル間の項目または系列の数を表します。 1 から 31999 の範囲内で値を設定できます。自動的に設定する場合は、空の文字列にします。 戻り値は常に数値です。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|目盛間の項目または系列の数を表します。|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|軸を表示するかどうかを表すブール値。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|グラフ要素の罫線の書式設定をクリアします。|
||[color](/javascript/api/excel/excel.chartborder#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|罫線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[スパイク](/javascript/api/excel/excel.chartdatalabel#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[position](/javascript/api/excel/excel.chartdatalabel#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|グラフのデータ ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[記号](/javascript/api/excel/excel.chartdatalabel#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|フォント名、フォントサイズ、色など、グラフの文字オブジェクトのフォントの属性を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|グラフの凡例の高さをポイント単位で表します。 凡例が表示されていない場合は Null。|
||[left](/javascript/api/excel/excel.chartlegend#left)|グラフの凡例の左のポイントを表します。 凡例が表示されていない場合は Null。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|凡例に含まれる凡例エントリのコレクションを表します。 読み取り専用です。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|凡例がグラフに影付きかどうかを表します。|
||[top](/javascript/api/excel/excel.chartlegend#top)|グラフ凡例の上を表します。|
||[width](/javascript/api/excel/excel.chartlegend#width)|グラフの凡例の幅をポイント単位で表します。 凡例が表示されていない場合は Null。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|グラフの凡例に表示される凡例エントリの高さを表します。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|グラフの凡例に含まれる凡例エントリのインデックスを表します。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|グラフの凡例エントリの左を表します。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|グラフの凡例エントリの上を表します。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|グラフの凡例に表示される凡例エントリの幅を表します。|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|コレクションに含まれる凡例エントリの数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|指定されたインデックスに位置する凡例エントリを返します。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|線の太さ (ポイント数) を表します。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|データ ポイントのマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|グラフ データ ポイントのデータ ラベルを返します。 読み取り専用です。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[罫線](/javascript/api/excel/excel.chartpointformat#border)|色、スタイル、および重みの情報を含むグラフデータポイントの罫線の書式を表します。 読み取り専用です。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|グラフ系列の種類を表します。 詳細については、「ChartType」を参照してください。|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|グラフ系列を削除します。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|グラフ系列のドーナツの穴の大きさを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|
||[対象](/javascript/api/excel/excel.chartseries#filtered)|データ系列がフィルター処理されるかどうかを表すブール値。 等高線グラフには適用されません。|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|グラフ系列間に設けられる間隔を表します。  横棒グラフと縦棒グラフでのみ有効です。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|系列のデータ ラベルの有無を表すブール値。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|グラフ系列のマーカー背景色を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|グラフ系列のマーカー前景色を表します。|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|グラフ系列のマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|グラフ系列のマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|グラフ グループ内でのグラフ系列のプロット順序を表します。|
||[曲線](/javascript/api/excel/excel.chartseries#trendlines)|データ系列に含まれる近似曲線のコレクションを表します。 読み取り専用です。|
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|グラフ系列のバブル サイズを設定します。 バブル チャートにのみ適用されます。|
||[setValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|グラフ系列の値を設定します。 散布図の場合、Y 軸の値を意味します。|
||[Setx軸値 (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|グラフ系列の X 軸の値を設定します。 散布図にのみ適用されます。|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|系列に影があるかどうかを表すブール型 (Boolean) の値を指定します。|
||[smooth](/javascript/api/excel/excel.chartseries#smooth)|系列が平滑化されるかどうかを表すブール値。 折れ線グラフおよび散布図にのみ適用されます。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add (name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|コレクションに新しい系列を追加します。 新たに追加された系列は、値を設定するか x 軸の値/バブルサイズを (グラフの種類に応じて) 表示されるまで表示されません。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (start: number, length: number)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|グラフタイトルのサブ文字列を取得します。 改行 ' \n ' は1文字もカウントします。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|グラフ タイトルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttitle#left)|グラフ タイトルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[position](/javascript/api/excel/excel.charttitle#position)|グラフ タイトルの位置を表します。 詳細については、「Excel Charttitle Position」を参照してください。|
||[height](/javascript/api/excel/excel.charttitle#height)|グラフ タイトルの高さ (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
||[width](/javascript/api/excel/excel.charttitle#width)|グラフ タイトルの幅 (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
||[setFormula (formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|A1 スタイルの表記法を使用するグラフ タイトルの数式を表す文字列値を設定します。|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|グラフ タイトルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttitle#top)|グラフ タイトルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|グラフ タイトルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[罫線](/javascript/api/excel/excel.charttitleformat#border)|グラフタイトルの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。 読み取り専用です。|
|[ChartTrendline 曲線](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[delete()](/javascript/api/excel/excel.charttrendline#delete--)|trendline オブジェクトを削除します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[y](/javascript/api/excel/excel.charttrendline#intercept)|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendline#name)|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[format](/javascript/api/excel/excel.charttrendline#format)|グラフの近似曲線の書式設定を表します。|
||[付け](/javascript/api/excel/excel.charttrendline#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendline#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
||[type](/javascript/api/excel/excel.charttrendline#type)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add (type?: ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|近似曲線のコレクションに新しい近似曲線を追加します。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|コレクションに含まれる近似曲線の数を返します。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|インデックス (項目配列内の挿入順序) に基づいて trendline オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|グラフの線の書式設定を表します。 値の取得のみ可能です。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.customproperty#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.customproperty#type)|カスタム プロパティの値の型を取得します。 読み取り専用。|
||[value](/javascript/api/excel/excel.customproperty#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|コレクションに含まれるすべてのデータ接続を更新します。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[判別](/javascript/api/excel/excel.documentproperties#author)|ブックの作成者を取得または設定します。|
||[項目](/javascript/api/excel/excel.documentproperties#category)|ブックのカテゴリを取得または設定します。|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|ブックのコメントを取得または設定します。|
||[company](/javascript/api/excel/excel.documentproperties#company)|ブックの会社を取得または設定します。|
||[キーワード](/javascript/api/excel/excel.documentproperties#keywords)|ブックのキーワードを取得または設定します。|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|ブックのマネージャーを取得または設定します。|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|ブックの作成日を取得します。 読み取り専用です。|
||[配色](/javascript/api/excel/excel.documentproperties#custom)|ブックのカスタム プロパティのコレクションを取得します。 読み取り専用です。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|ブックの最後の作成者を取得します。 読み取り専用です。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|ブックのリビジョン番号を取得します。 読み取り専用です。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|ブックの件名を取得または設定します。|
||[title](/javascript/api/excel/excel.documentproperties#title)|ブックのタイトルを取得または設定します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|名前付きアイテムの値と型を含むオブジェクトを返します。 読み取り専用です。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|名前付きアイテムの配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: number, Numrows: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|現在の Range オブジェクトと左上のセルが同じで、指定した数の行と列を含む Range オブジェクトを取得します。|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|範囲を base64 でエンコードされた png 画像としてレンダリングします。|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|指定された範囲の左上のセルを囲む領域を表す Range オブジェクトを返します。 周囲の領域は、この範囲に相対の空白の行と空白の列の任意の組み合わせで囲まれた範囲です。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|アクティブ セルに多数の値が含まれる場合、そのセルのカードを表示します。|
||[style](/javascript/api/excel/excel.range#style)|現在の範囲のスタイルを表します。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|該当する範囲内のすべてのセルのテキストの向きを設定します。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Range オブジェクトの列の幅が、シートの標準の幅と等しいかどうかを示します。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|ハイパーリンクの URL ターゲットを表します。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|ハイパーリンクのドキュメント参照先を表します。|
||[ポップヒント](/javascript/api/excel/excel.rangehyperlink#screentip)|ハイパーリンクの上にカーソルを合わせると表示される文字列を表します。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|該当する範囲内の左上端のセルに表示される文字列を表します。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[delete()](/javascript/api/excel/excel.style#delete--)|このスタイルを削除します。|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|スタイルでの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|スタイルに配置のプロパティ (AddIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、および TextOrientation) が含まれるかどうかを示します。|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|スタイルに罫線のプロパティ (Color、ColorIndex、LineStyle、Weight) が含まれているかどうかを示します。|
||[includeFont](/javascript/api/excel/excel.style#includefont)|スタイルにフォントのプロパティ (Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript、Underline) が含まれているかどうかを示します。|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|スタイルに NumberFormat プロパティが含まれているかどうかを示します。|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|スタイルに塗りつぶしのプロパティ (Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、PatternColorIndex) が含まれているかどうかを示します。|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|スタイルに保護のプロパティ (FormulaHidden および Locked) が含まれているかどうかを示します。|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.style#locked)|ワークシートが保護されている場合、オブジェクトがロックされるかどうかを示します。|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|スタイルで適用される読み上げ順序。|
||[borders](/javascript/api/excel/excel.style#borders)|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。|
||[Unset](/javascript/api/excel/excel.style#builtin)|スタイルが組み込みのスタイルであるかどうかを示します。|
||[fill](/javascript/api/excel/excel.style#fill)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.style#font)|スタイルのフォントを表す Font オブジェクト。|
||[name](/javascript/api/excel/excel.style#name)|スタイルの名前。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|スタイルで適用されるテキストの向き。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|スタイルで適用される垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|コレクションに新しいスタイルを追加します。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|名前に基づいてスタイルを取得します。|
||[items](/javascript/api/excel/excel.stylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|特定の表で、セル内のデータが変更されたときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|特定の表で選択範囲が変更されたときに発生します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|特定のワークシート上のテーブル内で変更されたエリアを表すアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 詳細については、「DataChangeType」を参照してください。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|特定のワークシート上のテーブルの変更された領域を表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|特定のワークシート上のテーブルの変更された領域を表す範囲を取得します。 null オブジェクトを返すこともあります。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|データが変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|ブックまたはワークシート内のテーブルでデータが変更されたときに発生します。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|特定のワークシート上のテーブル内で選択されたエリアを表す範囲のアドレスを取得します。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|選択範囲がテーブル内に収まっているかどうかを示します。IsInsideTable が false の場合、アドレスは無効です。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|選択範囲が変更されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。 読み取り専用です。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|選択範囲が変更されたワークシートの ID を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|ブックで現在アクティブなセルを取得します。|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|ブック内のすべてのデータ接続を表します。 読み取り専用です。|
||[name](/javascript/api/excel/excel.workbook#name)|ブックの名前を取得します。 読み取り専用です。|
||[プロパティ](/javascript/api/excel/excel.workbook#properties)|ブックのプロパティを取得します。 読み取り専用です。|
||[protection](/javascript/api/excel/excel.workbook#protection)|ブックの workbookProtection オブジェクトを返します。 読み取り専用です。|
||[線](/javascript/api/excel/excel.workbook#styles)|ブックに関連付けられているスタイルのコレクションを表します。 読み取り専用です。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect (password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|ブックを保護します。 ブックが保護されている場合は失敗します。|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|ブックが保護されているかどうかを示します。 読み取り専用です。|
||[保護の解除 (password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|ブックの保護を解除します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy (positionType?: Excel. ワークシートの種類, relativeTo?: Excel)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|ワークシートをコピーして、指定した位置に配置します。 コピーしたワークシートを返します。|
||[getRangeByIndexes (startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|特定の行インデックスと列インデックスから開始し、一定数の行と列にわたる、Range オブジェクトを取得します。|
||[Freezepanes プロパティが](/javascript/api/excel/excel.worksheet#freezepanes)|ワークシート上の固定されたウィンドウを操作するために使用できるオブジェクトを取得します。 読み取り専用です。|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|ワークシートがアクティブになるときに発生します。|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|特定のワークシートでデータが変更されたときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|ワークシートが非アクティブになるときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|特定のワークシートで選択範囲が変更されたときに発生します。|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。 読み取り専用です。|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|ワークシートのすべての列の標準 (既定) の幅を返すか設定します。|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|ワークシートのタブの色を取得または設定します。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|アクティブにされたワークシートの ID を取得します。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|ブックに追加されたワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 詳細については、「DataChangeType」を参照してください。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|ブック内のすべてのワークシートがアクティブになったときに発生します。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|新しいワークシートがブックに追加されるときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|ブック内のすべてのワークシートが非アクティブ化されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|ブックからワークシートが削除されるときに発生します。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|非アクティブにされたワークシートの ID を取得します。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|ブックから削除されたワークシートの ID を取得します。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: Range \|文字列)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|アクティブなワークシート ビューに固定セルを設定します。|
||[freezeColumns (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|ワークシートの最初の列 (複数可) を所定の場所に固定します。|
||[Freeゼロ Ws (count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|ワークシートの最初の行 (複数可) を所定の場所に固定します。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|アクティブなワークシート ビュー内の固定セルを記述する範囲を取得します。|
||[保持解除 ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|ワークシートからすべての固定ウィンドウを削除します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[保護の解除 (password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|ワークシートの保護を解除します。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|オブジェクトの編集を可能にするワークシート保護オプションを表します。|
||[allowEditScenarios シナリオ](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|シナリオの編集を可能にするワークシート保護オプションを表します。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|選択モードのワークシート保護オプションを表します。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|特定のワークシートで選択されたエリアを表す範囲のアドレスを取得します。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|選択範囲が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
