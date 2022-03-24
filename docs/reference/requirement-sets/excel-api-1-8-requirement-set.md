---
title: Excel JavaScript API 要件セット 1.8
description: ExcelApi 1.8 要件セットの詳細。
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 39f3a5daf89849d3f8517794ab8cd4214309a667
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746856"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>JavaScript API 1.8 Excel新機能

Excel JavaScript API 要件セット 1.8 の機能には、ピボットテーブル、データの入力規則、グラフ、グラフのイベント、パフォーマンス オプション、ブック作成に対応する API が含まれます。

## <a name="pivottable"></a>ピボットテーブル

ピボットテーブル API の Wave 2 では、アドインでピボットテーブルの階層を設定できます。 データとデータの集計方法を制御できるようになりました。 新しいピボットテーブルの機能について詳しくは、[ピボットテーブルの記事](../../excel/excel-add-ins-pivottables.md)を参照してください。

## <a name="data-validation"></a>データの入力規則

データの入力規則により、ユーザーがワークシートに入力する内容を制御できます。 定義済みの回答セットにセルを制限したり、望ましくない入力に関する警告をポップアップ表示したりできます。 詳細については、[データの入力規則を範囲に追加する方法](../../excel/excel-add-ins-data-validation.md)を参照してください。

## <a name="charts"></a>グラフ

グラフ要素をより詳細にプログラムで制御できる、一連のグラフ API がさらに追加されました。 凡例、軸、近似曲線、プロット エリアがより使いやすくなっています。

## <a name="events"></a>イベント

グラフの[イベント](../../excel/excel-add-ins-events.md)がさらに追加されました。 グラフを操作するユーザーに対し、アドインで対応できます。 ブック全体にわたり、起動する[イベントの切り替え](../../excel/performance.md#enable-and-disable-events)もできます。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.8 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.8 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|データの検証に使用する演算子。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#excel-excel-chart-categorylabellevel-member)|ソース カテゴリ ラベルのレベルを参照して、グラフ カテゴリ ラベル レベル列挙定数を指定します。|
||[displayBlanksAs](/javascript/api/excel/excel.chart#excel-excel-chart-displayblanksas-member)|空白のセルをグラフにプロットする方法を指定します。|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|グラフがアクティブ化されると発生します。|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|グラフが非アクティブ化された場合に発生します。|
||[plotArea](/javascript/api/excel/excel.chart#excel-excel-chart-plotarea-member)|グラフのプロット領域を表します。|
||[plotBy](/javascript/api/excel/excel.chart#excel-excel-chart-plotby-member)|列や行がグラフのデータ系列として使用される方法を指定します。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#excel-excel-chart-plotvisibleonly-member)|true の場合、可視セルだけがプロットされます。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#excel-excel-chart-seriesnamelevel-member)|ソース 系列名のレベルを参照して、グラフ系列名レベルの列挙定数を指定します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#excel-excel-chart-showdatalabelsovermaximum-member)|値が値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.chart#excel-excel-chart-style-member)|グラフのグラフ スタイルを指定します。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|アクティブ化されたグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|グラフをアクティブ化するワークシートの ID を取得します。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|ワークシートに追加されるグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|グラフが追加されるワークシートの ID を取得します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[配置](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|指定した軸目盛ラベルの配置を指定します。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|値軸がカテゴリの間でカテゴリ軸と交差する場合に指定します。|
||[multiLevel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-multilevel-member)|軸がマルチレベルの場合に指定します。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-numberformat-member)|軸目盛ラベルの書式コードを指定します。|
||[offset](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-offset-member)|ラベルのレベル間の距離と、最初のレベルと軸線の間の距離を指定します。|
||[position](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-position-member)|他の軸が交差する指定した軸位置を指定します。|
||[positionAt](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-positionat-member)|他の軸が交差する軸位置を指定します。|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setpositionat-member(1))|他の軸が交差する指定した軸位置を設定します。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-textorientation-member)|グラフ軸目盛ラベルのテキストの向きを指定します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|グラフの塗りつぶしの書式設定を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[罫線](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|色、線のスタイル、太さなど、グラフ軸のタイトルの罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|グラフ軸のタイトルの塗りつぶしの書式設定を指定します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|グラフ要素の罫線の書式設定をクリアします。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|グラフがアクティブ化されると発生します。|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|ワークシートに新しいグラフが追加された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|グラフが非アクティブ化された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|グラフが削除された場合に発生します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|グラフのデータ ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|グラフのデータ ラベルの高さ (ポイント数) を返します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|グラフのデータ ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-numberformat-member)|データ ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-text-member)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-textorientation-member)|グラフ データ ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-top-member)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-verticalalignment-member)|グラフのデータ ラベルの垂直方向の配置を表します。|
||[width](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-width-member)|グラフのデータ ラベルの幅 (ポイント数) を返します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[罫線](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|グラフ データ ラベルの水平方向の配置を指定します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-numberformat-member)|データ ラベルの形式コードを指定します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-textorientation-member)|データ ラベルのテキストの向きを示す角度を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-verticalalignment-member)|グラフのデータ ラベルの垂直方向の配置を表します。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|非アクティブ化されているグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|グラフが非アクティブ化されているワークシートの ID を取得します。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|ワークシートから削除されたグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|グラフが削除されるワークシートの ID を取得します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|グラフの凡例の凡例エントリの高さを指定します。|
||[index](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|グラフ凡例の凡例エントリのインデックスを指定します。|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|グラフの凡例エントリの左の値を指定します。|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|グラフ凡例エントリの上部を指定します。|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|グラフ Legend の凡例エントリの幅を表します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[罫線](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|グラフプロット領域の書式を指定します。|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|プロット領域の高さの値を指定します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|プロット領域の内側の高さの値を指定します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|プロット領域の内側の左の値を指定します。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|プロット領域の内側の上の値を指定します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|プロット領域の内側の幅の値を指定します。|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|プロット領域の左の値を指定します。|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|プロット領域の位置を指定します。|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|プロット領域の上の値を指定します。|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|プロット領域の幅の値を指定します。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[罫線](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|グラフプロット領域の罫線属性を指定します。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|背景の書式設定情報を含むオブジェクトの塗りつぶしの形式を指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-axisgroup-member)|指定した系列のグループを指定します。|
||[dataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-datalabels-member)|系列内のすべてのデータ ラベルのコレクションを表します。|
||[爆発](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-explosion-member)|円グラフまたはドーナツ グラフのスライスの展開値を指定します。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-firstsliceangle-member)|最初の円グラフまたはドーナツ グラフのスライスの角度を度 (垂直方向から時計回り) で指定します。|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertifnegative-member)|True の場合Excelが負の数値に対応する場合に、アイテム内のパターンを反転します。|
||[オーバーラップ](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-overlap-member)|横棒と縦棒の配置方法を指定します。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-secondplotsize-member)|円グラフまたは円グラフの 2 番目のセクションのサイズを、プライマリ 円グラフのサイズに対する割合で指定します。|
||[splitType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splittype-member)|円グラフまたは円グラフの 2 つのセクションを分割する方法を指定します。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-varybycategories-member)|True の場合Excelデータ マーカーに異なる色またはパターンを割り当てる必要があります。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|近似曲線を前方へ拡張するときの区間数を表します。|
||[label](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|グラフの近似曲線のラベルを表します。|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|True の場合、トレンドラインの r-2 乗値がグラフに表示されます。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|傾向線ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|グラフの傾向線ラベルの形式。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|A1 スタイル表記を使用してグラフの傾向線ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|グラフの傾向線ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|グラフのトレンドライン ラベルの左端からグラフ領域の左端までの距離をポイントで表します。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-numberformat-member)|傾向線ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-text-member)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-textorientation-member)|グラフの傾向線ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-top-member)|グラフのトレンドライン ラベルの上端からグラフ領域の上端までの距離をポイントで表します。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-verticalalignment-member)|グラフの傾向線ラベルの垂直方向の配置を表します。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-width-member)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[罫線](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|色、線のスタイル、太さなど、罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|現在のグラフの傾向線ラベルの塗りつぶしの形式を指定します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|グラフの傾向線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|ユーザーの入力規則のカスタム数式。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[field](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|DataPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|DataPivotHierarchy の位置。|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|DataPivotHierarchy を既定値にリセットします。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|データを特定の集計計算として表示する必要がある場合に指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|DataPivotHierarchy のすべての項目を表示する場合に指定します。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|名前または ID で DataPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|名前に基づいて DataPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|現在の軸からピボット階層を削除します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|現在の範囲からデータの入力規則をクリアします。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|空白のセルに対してデータ検証を実行する場合に指定します。|
||[Prompt](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|ユーザーがセルを選択するときにプロンプトを表示します。|
||[ルール](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|さまざまな種類のデータ検証条件を含むデータ検証ルール。|
||[type](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|データ検証の種類については、「詳細」 `Excel.DataValidationType` を参照してください。|
||[有効](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|エラー通知メッセージを表します。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|ユーザーが無効なデータを入力した場合にエラー通知ダイアログを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|データ検証アラートの種類については、「詳細」を `Excel.DataValidationAlertStyle` 参照してください。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|エラー通知ダイアログのタイトルを表します。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|プロンプトのメッセージを指定します。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|ユーザーがデータ検証を使用してセルを選択するときにプロンプトを表示する場合に指定します。|
||[title](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|プロンプトのタイトルを指定します。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[カスタム](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|データ検証条件のカスタム数式。|
||[date](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|日付のデータ検証条件。|
||[decimal](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|10 進数のデータ検証条件。|
||[list](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|リストのデータ検証条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|テキストの長さデータの検証条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|時刻のデータ検証条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|数値データの検証条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|データの検証に使用する演算子。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|複数のフィルター項目を許可するかどうかを指定します。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|FilterPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|FilterPivotHierarchy の位置。|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|FilterPivotHierarchy を既定値にリセットします。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|名前または ID で FilterPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|名前に基づいて FilterPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|現在の軸からピボット階層を削除します。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|セル ドロップダウンにリストを表示するかどうかを指定します。|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|データの入力規則のリストのソース。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|PivotField の ID。|
||[アイテム](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|PivotField に関連付けられた PivotItems を返します。|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|PivotField の名前。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|PivotField のすべての項目を表示するかどうかを指定します。|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|PivotField を並べ替えます。|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|PivotField の小計。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|コレクション内のピボット フィールドの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|名前または ID で PivotField を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|名前でピボットフィールドを取得します。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[fields](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|PivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|PivotHierarchy の名前。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|名前または ID で PivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|名前に基づいて PivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|PivotItem の ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|PivotItem の名前。|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|PivotItem が表示される場合に指定します。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|コレクション内の PivotItems の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|名前または ID で PivotItem を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|名前で PivotItem を取得します。|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|ピボットテーブルの列ラベルが存在する範囲を返します。|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|ピボットテーブルのデータ値が存在する範囲を返します。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|ピボットテーブルのフィルター エリアの範囲を返します。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|ピボットテーブルの行ラベルが存在する範囲を返します。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|ピボットテーブル レポートに列の総計が表示される場合に指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|ピボットテーブル レポートに行の総計が表示される場合に指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|このプロパティは、ピボットテーブル `SubtotalLocationType` のすべてのフィールドを示します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|ピボットテーブルのデータ ピボット階層。|
||[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|ピボットテーブルを削除します。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|ピボットテーブルのピボット階層。|
||[レイアウト](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|ピボットテーブルの行ピボット階層。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|指定したソース データに基づいてピボットテーブルを追加し、移動先範囲の左上のセルに挿入します。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|dataValidation オブジェクトを返します。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|RowColumnPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|RowColumnPivotHierarchy の位置。|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|RowColumnPivotHierarchy を既定値にリセットします。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|名前または ID で RowColumnPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|名前に基づいて RowColumnPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|現在の軸からピボット階層を削除します。|
|[ランタイム](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#excel-excel-runtime-enableevents-member)|現在の作業ウィンドウまたはコンテンツ アドインの JavaScript イベントを切り替えます。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|PivotField を使用して、型 `ShowAs` に応じて計算を基に計算を行います ( `ShowAsCalculation` 該当する場合は、それ以外の場合 `null`)。|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|型に応じて該当`ShowAs``ShowAsCalculation`する場合は、計算の基に設定するアイテム。それ以外の場合`null`。|
||[計算](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|PivotField `ShowAs` に使用する計算。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|セル内のテキスト配置が等しい分布に設定されている場合に、テキストが自動的にインデントされる場合に指定します。|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|スタイルで適用されるテキストの向き。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|に `Automatic` 設定されている場合 `true`、他のすべての値は、 を設定するときに無視されます `Subtotals`。|
||[平均](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[製品](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[差異](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|数値 ID を返します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrange-member(1))|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrangeornullobject-member(1))|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|ブックが `true` 読み取り専用モードで開いている場合に返します。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)|ワークシートの計算時に発生します。|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|グリッド線がユーザーに表示される場合に指定します。|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|ユーザーに見出しを表示する場合に指定します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-worksheetid-member)|計算が発生したワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrange-member(1))|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrangeornullobject-member(1))|特定のワークシートで変更されたエリアを表す範囲を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|ブック内のワークシートが計算される場合に発生します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
