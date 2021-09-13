---
title: ExcelJavaScript API 要件セット 1.8
description: ExcelApi 1.8 要件セットの詳細。
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e97dd98d024b27aa58ca6f0c76fdee17b657c7c9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154957"
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

次の表に、JavaScript API 要件セット 1.8 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.8 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|データの検証に使用する演算子。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categoryLabelLevel)|ソース カテゴリ ラベルのレベルを参照して、グラフ カテゴリ ラベル レベル列挙定数を指定します。|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayBlanksAs)|空白のセルをグラフにプロットする方法を指定します。|
||[plotBy](/javascript/api/excel/excel.chart#plotBy)|列や行がグラフのデータ系列として使用される方法を指定します。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotVisibleOnly)|true の場合、可視セルだけがプロットされます。|
||[onActivated](/javascript/api/excel/excel.chart#onActivated)|グラフがアクティブ化されると発生します。|
||[onDeactivated](/javascript/api/excel/excel.chart#onDeactivated)|グラフが非アクティブ化された場合に発生します。|
||[plotArea](/javascript/api/excel/excel.chart#plotArea)|グラフのプロット領域を表します。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesNameLevel)|ソース 系列名のレベルを参照して、グラフ系列名レベルの列挙定数を指定します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showDataLabelsOverMaximum)|値が値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.chart#style)|グラフのグラフ スタイルを指定します。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartId)|アクティブ化されたグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetId)|グラフをアクティブ化するワークシートの ID を取得します。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartId)|ワークシートに追加されるグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetId)|グラフが追加されるワークシートの ID を取得します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[配置](/javascript/api/excel/excel.chartaxis#alignment)|指定した軸目盛ラベルの配置を指定します。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isBetweenCategories)|値軸がカテゴリの間でカテゴリ軸と交差する場合に指定します。|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multiLevel)|軸がマルチレベルの場合に指定します。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberFormat)|軸目盛ラベルの書式コードを指定します。|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|ラベルのレベル間の距離と、最初のレベルと軸線の間の距離を指定します。|
||[position](/javascript/api/excel/excel.chartaxis#position)|他の軸が交差する指定した軸位置を指定します。|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionAt)|他の軸が交差する軸位置を指定します。|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setPositionAt_value_)|他の軸が交差する指定した軸位置を設定します。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textOrientation)|グラフ軸目盛ラベルのテキストの向きを指定します。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|グラフの塗りつぶしの書式設定を指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setFormula_formula_)|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|色、線のスタイル、太さなど、グラフ軸のタイトルの罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|グラフ軸のタイトルの塗りつぶしの書式設定を指定します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear__)|グラフ要素の罫線の書式設定をクリアします。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onActivated)|グラフがアクティブ化されると発生します。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onAdded)|ワークシートに新しいグラフが追加された場合に発生します。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#onDeactivated)|グラフが非アクティブ化された場合に発生します。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#onDeleted)|グラフが削除された場合に発生します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autoText)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalAlignment)|グラフのデータ ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberFormat)|データ ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|グラフのデータ ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textOrientation)|グラフ データ ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalAlignment)|グラフのデータ ラベルの垂直方向の配置を表します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autoText)|データ ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalAlignment)|グラフ データ ラベルの水平方向の配置を指定します。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberFormat)|データ ラベルの形式コードを指定します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textOrientation)|データ ラベルのテキストの向きを示す角度を表します。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalAlignment)|グラフのデータ ラベルの垂直方向の配置を表します。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartId)|非アクティブ化されているグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetId)|グラフが非アクティブ化されているワークシートの ID を取得します。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartId)|ワークシートから削除されたグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetId)|グラフが削除されるワークシートの ID を取得します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|グラフの凡例の凡例エントリの高さを指定します。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|グラフ凡例の凡例エントリのインデックスを指定します。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|グラフの凡例エントリの左の値を指定します。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|グラフ凡例エントリの上部を指定します。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|グラフ Legend の凡例エントリの幅を表します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|プロット領域の高さの値を指定します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideHeight)|プロット領域の内側の高さの値を指定します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideLeft)|プロット領域の内側の左の値を指定します。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insideTop)|プロット領域の内側の上の値を指定します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insideWidth)|プロット領域の内側の幅の値を指定します。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|プロット領域の左の値を指定します。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|プロット領域の位置を指定します。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|グラフプロット領域の書式を指定します。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|プロット領域の上の値を指定します。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|プロット領域の幅の値を指定します。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|グラフプロット領域の罫線属性を指定します。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|背景の書式設定情報を含むオブジェクトの塗りつぶしの形式を指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisGroup)|指定した系列のグループを指定します。|
||[爆発](/javascript/api/excel/excel.chartseries#explosion)|円グラフまたはドーナツ グラフのスライスの展開値を指定します。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstSliceAngle)|最初の円グラフまたはドーナツ グラフのスライスの角度を度 (垂直方向から時計回り) で指定します。|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertIfNegative)|True の場合Excelに対応する場合は、アイテム内のパターンを反転します。|
||[オーバーラップ](/javascript/api/excel/excel.chartseries#overlap)|横棒と縦棒の配置方法を指定します。|
||[dataLabels](/javascript/api/excel/excel.chartseries#dataLabels)|系列内のすべてのデータ ラベルのコレクションを表します。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondPlotSize)|円グラフまたは円グラフの 2 番目のセクションのサイズを、プライマリ 円グラフのサイズに対する割合で指定します。|
||[splitType](/javascript/api/excel/excel.chartseries#splitType)|円グラフまたは円グラフの 2 つのセクションを分割する方法を指定します。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varyByCategories)|True の場合Excelデータ マーカーに異なる色またはパターンを割り当てる必要があります。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardPeriod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardPeriod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[label](/javascript/api/excel/excel.charttrendline#label)|グラフの近似曲線のラベルを表します。|
||[showEquation](/javascript/api/excel/excel.charttrendline#showEquation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showRSquared)|True の場合、トレンドラインの r-2 乗値がグラフに表示されます。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autoText)|傾向線ラベルがコンテキストに基づいて適切なテキストを自動的に生成する場合に指定します。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|A1 スタイル表記を使用してグラフの傾向線ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalAlignment)|グラフの傾向線ラベルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|グラフのトレンドライン ラベルの左端からグラフ領域の左端までの距離をポイントで表します。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberFormat)|傾向線ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|グラフの傾向線ラベルの形式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textOrientation)|グラフの傾向線ラベルのテキストの向きを示す角度を表します。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|グラフのトレンドライン ラベルの上端からグラフ領域の上端までの距離をポイントで表します。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalAlignment)|グラフの傾向線ラベルの垂直方向の配置を表します。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|色、線のスタイル、太さなど、罫線の形式を指定します。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|現在のグラフの傾向線ラベルの塗りつぶしの形式を指定します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|グラフの傾向線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を指定します。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|ユーザーの入力規則のカスタム数式。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberFormat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy の位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#setToDefault__)|DataPivotHierarchy を既定値にリセットします。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showAs)|データを特定の集計計算として表示する必要がある場合に指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeBy)|DataPivotHierarchy のすべての項目を表示する場合に指定します。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add_pivotHierarchy_)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getCount__)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItem_name_)|名前または ID で DataPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItemOrNullObject_name_)|名前に基づいて DataPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove_DataPivotHierarchy_)|現在の軸からピボット階層を削除します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear__)|現在の範囲からデータの入力規則をクリアします。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#errorAlert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreBlanks)|空白のセルに対してデータ検証を実行する場合に指定します。|
||[Prompt](/javascript/api/excel/excel.datavalidation#prompt)|ユーザーがセルを選択するときにプロンプトを表示します。|
||[type](/javascript/api/excel/excel.datavalidation#type)|データ検証の種類については、「詳細 `Excel.DataValidationType` 」を参照してください。|
||[有効](/javascript/api/excel/excel.datavalidation#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
||[ルール](/javascript/api/excel/excel.datavalidation#rule)|さまざまな種類のデータ検証条件を含むデータ検証ルール。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[メッセージ](/javascript/api/excel/excel.datavalidationerroralert#message)|エラー通知メッセージを表します。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showAlert)|ユーザーが無効なデータを入力した場合にエラー通知ダイアログを表示するかどうかを指定します。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|データ検証アラートの種類については、「詳細」 `Excel.DataValidationAlertStyle` を参照してください。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|エラー通知ダイアログのタイトルを表します。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[メッセージ](/javascript/api/excel/excel.datavalidationprompt#message)|プロンプトのメッセージを指定します。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showPrompt)|ユーザーがデータ検証を使用してセルを選択するときにプロンプトを表示する場合に指定します。|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|プロンプトのタイトルを指定します。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[カスタム](/javascript/api/excel/excel.datavalidationrule#custom)|データ検証条件のカスタム数式。|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|日付のデータ検証条件。|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|10 進数のデータ検証条件。|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|リストのデータ検証条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textLength)|テキストの長さデータの検証条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|時刻のデータ検証条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholeNumber)|数値データの検証条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|演算子プロパティが GreaterThan などのバイナリ演算子に設定されている場合に、右側のオペランドを指定します (左側のオペランドは、ユーザーがセルに入力しようとする値です)。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|3 項演算子 Between と NotBetween を使用して、上限オペランドを指定します。|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|データの検証に使用する演算子。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enableMultipleFilterItems)|複数のフィルター項目を許可するかどうかを指定します。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#setToDefault__)|FilterPivotHierarchy を既定値にリセットします。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add_pivotHierarchy_)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getCount__)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItem_name_)|名前または ID で FilterPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItemOrNullObject_name_)|名前に基づいて FilterPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove_filterPivotHierarchy_)|現在の軸からピボット階層を削除します。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#inCellDropDown)|セル ドロップダウンにリストを表示するかどうかを指定します。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|データの入力規則のリストのソース。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField の名前。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField の ID。|
||[アイテム](/javascript/api/excel/excel.pivotfield#items)|PivotField に関連付けられているピボット フィールドを返します。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showAllItems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortByLabels_sortBy_)|PivotField を並べ替えます。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField の小計。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getCount__)|コレクション内のピボット フィールドの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItem_name_)|名前または ID で PivotField を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItemOrNullObject_name_)|名前でピボットフィールドを取得します。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy の名前。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy の ID。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getCount__)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItem_name_)|名前または ID で PivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItemOrNullObject_name_)|名前に基づいて PivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isExpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem の名前。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem の ID。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|PivotItem が表示される場合に指定します。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getCount__)|コレクション内の PivotItems の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItem_name_)|名前または ID で PivotItem を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItemOrNullObject_name_)|名前で PivotItem を取得します。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getColumnLabelRange__)|ピボットテーブルの列ラベルが存在する範囲を返します。|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getDataBodyRange__)|ピボットテーブルのデータ値が存在する範囲を返します。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getFilterAxisRange__)|ピボットテーブルのフィルター エリアの範囲を返します。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getRange__)|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getRowLabelRange__)|ピボットテーブルの行ラベルが存在する範囲を返します。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layoutType)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showColumnGrandTotals)|ピボットテーブル レポートに列の総計が表示される場合に指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showRowGrandTotals)|ピボットテーブル レポートに行の総計が表示される場合に指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotalLocation)|このプロパティは、ピボット `SubtotalLocationType` テーブルのすべてのフィールドを示します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete__)|ピボットテーブルを削除します。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnHierarchies)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#dataHierarchies)|ピボットテーブルのデータ ピボット階層。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterHierarchies)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottable#hierarchies)|ピボットテーブルのピボット階層。|
||[レイアウト](/javascript/api/excel/excel.pivottable#layout)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowHierarchies)|ピボットテーブルの行ピボット階層。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: \| \| Range string Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add_name__source__destination_)|指定したソース データに基づいてピボットテーブルを追加し、移動先範囲の左上のセルに挿入します。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#dataValidation)|dataValidation オブジェクトを返します。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy の ID。|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#setToDefault__)|RowColumnPivotHierarchy を既定値にリセットします。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add_pivotHierarchy_)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getCount__)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItem_name_)|名前または ID で RowColumnPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItemOrNullObject_name_)|名前に基づいて RowColumnPivotHierarchy を取得します。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove_rowColumnPivotHierarchy_)|現在の軸からピボット階層を削除します。|
|[ランタイム](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableEvents)|現在の作業ウィンドウまたはコンテンツ アドインの JavaScript イベントを切り替えます。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#baseField)|PivotField を使用して、型に応じて計算を基に計算を行います (該当する場合 `ShowAs` `ShowAsCalculation` は、それ以外 `null` の場合)。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseItem)|型に応じて該当する場合は、計算の基に設定するアイテム `ShowAs` `ShowAsCalculation` 。それ以外の場合 `null` 。|
||[計算](/javascript/api/excel/excel.showasrule#calculation)|`ShowAs`PivotField に使用する計算。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoIndent)|セル内のテキスト配置が等しい分布に設定されている場合に、テキストが自動的にインデントされる場合に指定します。|
||[textOrientation](/javascript/api/excel/excel.style#textOrientation)|スタイルで適用されるテキストの向き。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|に `Automatic` 設定されている場合 `true` 、他のすべての値は、 を設定するときに無視 `Subtotals` されます。|
||[平均](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countNumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[製品](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standardDeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standardDeviationP)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[差異](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#varianceP)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyId)|数値 ID を返します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRange_ctx_)|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRangeOrNullObject_ctx_)|特定のワークシートのテーブルの変更された領域を表す範囲を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readOnly)|ブックが `true` 読み取り専用モードで開いている場合に返します。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#onCalculated)|ワークシートの計算時に発生します。|
||[showGridlines](/javascript/api/excel/excel.worksheet#showGridlines)|グリッド線がユーザーに表示される場合に指定します。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showHeadings)|ユーザーに見出しを表示する場合に指定します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetId)|計算が発生したワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRange_ctx_)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRangeOrNullObject_ctx_)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#onCalculated)|ブック内のワークシートが計算される場合に発生します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
