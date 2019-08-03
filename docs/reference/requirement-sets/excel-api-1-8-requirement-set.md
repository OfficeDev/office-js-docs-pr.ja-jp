---
title: Excel JavaScript API 要件セット1.8
description: ExcelApi 1.8 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6849ccb3dc83275509d26c63054a518d41cb060e
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064894"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 の新機能

Excel JavaScript API 要件セット 1.8 の機能には、ピボットテーブル、データの入力規則、グラフ、グラフのイベント、パフォーマンス オプション、ブック作成に対応する API が含まれます。

## <a name="pivottable"></a>ピボットテーブル

ピボットテーブル API の Wave 2 では、アドインでピボットテーブルの階層を設定できます。 データとデータの集計方法を制御できるようになりました。 新しいピボットテーブルの機能について詳しくは、[ピボットテーブルの記事](/office/dev/add-ins/excel/excel-add-ins-pivottables)を参照してください。

## <a name="data-validation"></a>データの入力規則

データの入力規則により、ユーザーがワークシートに入力する内容を制御できます。 定義済みの回答セットにセルを制限したり、望ましくない入力に関する警告をポップアップ表示したりできます。 詳細については、[データの入力規則を範囲に追加する方法](/office/dev/add-ins/excel/excel-add-ins-data-validation)を参照してください。

## <a name="charts"></a>グラフ

グラフ要素をより詳細にプログラムで制御できる、一連のグラフ API がさらに追加されました。 凡例、軸、近似曲線、プロット エリアがより使いやすくなっています。

## <a name="events"></a>イベント

グラフの[イベント](/office/dev/add-ins/excel/excel-add-ins-events)がさらに追加されました。 グラフを操作するユーザーに対し、アドインで対応できます。 ブック全体にわたり、起動する[イベントの切り替え](/office/dev/add-ins/excel/performance#enable-and-disable-events)もできます。

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.8 の Api を示します。 Excel JavaScript API 要件セット1.8 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「[要件セット1.8 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.8)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Operator プロパティが GreaterThan などのバイナリ演算子に設定されている場合、右のオペランドを指定します (左のオペランドは、ユーザーがセルに入力しようとした値です)。 三項演算子と NotBetween の間の and notbetween で、下限オペランドを指定します。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|三項演算子を and NotBetween で指定すると、上限オペランドが指定されます。 は、GreaterThan などの二項演算子では使用されません。|
||[演算子](/javascript/api/excel/excel.basicdatavalidation#operator)|データの検証に使用する演算子。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|を参照する Chartsets Labellevel 列挙定数を設定または返します。|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|空白のセルがグラフでプロットされる方法を返すか設定します。 読み取り/書き込み可能。|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|グラフ上で列または行がデータ系列として使用される方法を返すか設定します。 読み取り/書き込み可能。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|true の場合、可視セルだけがプロットされます。false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|グラフがアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|グラフが非アクティブになったときに発生します。|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|グラフのプロット エリアを表します。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|ChartSeriesNameLevel を参照する列挙定数を設定または返します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|
||[style](/javascript/api/excel/excel.chart#style)|グラフのグラフ スタイルを返すか設定します。 読み取り/書き込み可能。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|アクティブにされたグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|グラフがアクティブにされたワークシートの ID を取得します。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|ワークシートに追加されたグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|グラフが追加されたワークシートの ID を取得します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[策定](/javascript/api/excel/excel.chartaxis#alignment)|指定した軸の目盛ラベルの配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|項目の境界で数値軸が項目軸と交差するかどうかを表します。|
||[階層](/javascript/api/excel/excel.chartaxis#multilevel)|軸がマルチレベルかどうかを表します。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|軸の目盛ラベルの書式コードを表します。|
||[交互](/javascript/api/excel/excel.chartaxis#offset)|ラベルのレベル間の距離、および先頭レベルと軸線との距離を表します。 値は 0 から 1000 の範囲内でなければなりません。|
||[position](/javascript/api/excel/excel.chartaxis#position)|指定した軸と他の軸との交差位置を表します。 詳細については、「ChartAxisPosition」を参照してください。|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|指定した軸と他の軸との交差位置を表します。 このプロパティを設定するには、SetPositionAt(double) メソッドを使用する必要があります。|
||[setPositionAt (value: number)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|指定した軸と他の軸との交差位置を設定します。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|軸の目盛ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|グラフの塗りつぶしの書式設定を表します。 読み取り専用です。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[罫線](/javascript/api/excel/excel.chartaxistitleformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|グラフの塗りつぶしの書式設定を表します。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|グラフ要素の罫線の書式設定をクリアします。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|グラフがアクティブになったときに発生します。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|新しいグラフがワークシートに追加されるときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|グラフが非アクティブになったときに発生します。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|グラフが削除されるときに発生します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[スパイク](/javascript/api/excel/excel.chartdatalabel#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|グラフのデータ ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[罫線](/javascript/api/excel/excel.chartdatalabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[スパイク](/javascript/api/excel/excel.chartdatalabels#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|データ ラベルの書式コードを表します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|非アクティブにされたグラフの ID を取得します。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|グラフが非アクティブにされたワークシートの ID を取得します。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|ワークシートから削除されたグラフの ID を取得します。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|グラフが削除されたワークシートの ID を取得します。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|グラフの凡例に表示される凡例エントリの高さを表します。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|グラフの凡例に含まれる凡例エントリのインデックスを表します。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|グラフの凡例エントリの左を表します。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|グラフの凡例エントリの上を表します。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|グラフの凡例に表示される凡例エントリの幅を表します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[罫線](/javascript/api/excel/excel.chartlegendformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|プロット エリアの height 値を表します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|プロット エリアの insideHeight 値を表します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|プロット エリアの insideLeft 値を表します。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|プロット エリアの insideTop 値を表します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|プロット エリアの insideWidth 値を表します。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|プロット エリアの left 値を表します。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|プロット エリアの位置を表します。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|グラフ プロット エリアの書式設定を表します。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|プロット エリアの top 値を表します。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|プロット エリアの width 値を表します。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[罫線](/javascript/api/excel/excel.chartplotareaformat#border)|グラフ プロット エリアの罫線の属性を表します。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|指定した系列のグループを取得または設定します。 読み取り/書き込み|
||[切り出し](/javascript/api/excel/excel.chartseries#explosion)|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|横棒と縦棒の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|系列に含まれるすべてのデータ ラベルのコレクションを表します。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[ChartTrendline 曲線](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendline#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendline#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[スパイク](/javascript/api/excel/excel.charttrendlinelabel#autotext)|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|グラフの近似曲線ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|近似曲線ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|グラフの近似曲線ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|グラフの近似曲線ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[罫線](/javascript/api/excel/excel.charttrendlinelabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|現在のグラフの近似曲線ラベルの塗りつぶしの書式設定を表します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|ユーザーの入力規則のカスタム数式。 これにより、重複を防止したり、セル範囲内の合計を制限したりする特別な入力ルールが作成されます。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy の位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy の ID。|
||[setToDefault ()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|DataPivotHierarchy を既定値にリセットします。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|データを特定の集計計算として表示するかどうかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|名前または ID に基づいて DataPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|名前に基づいて DataPivotHierarchy を取得します。 DataPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|現在の軸からピボット階層を削除します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|現在の範囲からデータの入力規則をクリアします。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|
||[Prompt](/javascript/api/excel/excel.datavalidation#prompt)|ユーザーがセルを選択したときにメッセージを表示します。|
||[type](/javascript/api/excel/excel.datavalidation#type)|データの入力規則の種類。詳細については、Excel.DataValidationType を参照してください。|
||[有効な](/javascript/api/excel/excel.datavalidation#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
||[除外](/javascript/api/excel/excel.datavalidation#rule)|さまざまな種類のデータ検証条件を含むデータ入力規則。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|エラー警告メッセージを表します。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|無効なデータが入力されたときにエラー警告ダイアログを表示するかどうかを指定します。 既定値は true です。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|データの入力規則に対する警告の種類を表します。詳細については、Excel.DataValidationAlertStyle を参照してください。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|エラー警告ダイアログのタイトルを表します。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|プロンプトのメッセージを表します。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|ユーザーがデータの入力規則が適用されているセルを選択したときに、プロンプトを表示するかどうかを指定します。|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|プロンプトのタイトルを表します。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[配色](/javascript/api/excel/excel.datavalidationrule#custom)|データ検証条件のカスタム数式。|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|日付のデータ検証条件。|
||[十進](/javascript/api/excel/excel.datavalidationrule#decimal)|10 進数のデータ検証条件。|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|リストのデータ検証条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|テキスト長のデータ検証条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|時刻のデータ検証条件。|
||[整数](/javascript/api/excel/excel.datavalidationrule#wholenumber)|整数のデータ検証条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Operator プロパティが GreaterThan などのバイナリ演算子に設定されている場合、右のオペランドを指定します (左のオペランドは、ユーザーがセルに入力しようとした値です)。 三項演算子と NotBetween の間の and notbetween で、下限オペランドを指定します。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|三項演算子を and NotBetween で指定すると、上限オペランドが指定されます。 は、GreaterThan などの二項演算子では使用されません。|
||[演算子](/javascript/api/excel/excel.datetimedatavalidation#operator)|データの検証に使用する演算子。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy の ID。|
||[setToDefault ()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|FilterPivotHierarchy を既定値にリセットします。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|名前または ID に基づいて FilterPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|名前に基づいて FilterPivotHierarchy を取得します。 FilterPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|セルのドロップダウンにリストを表示するかどうかを指定します。既定では true に設定されます。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|データの入力規則のリストのソース。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField の名前。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField の ID。|
||[items](/javascript/api/excel/excel.pivotfield#items)|PivotField で構成される PivotItems を返します。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField を並べ替えます。 DataPivotHierarchy を指定すると、そのピボット階層に基づいて並べ替えが適用されます。指定しない場合、ピボット フィールド自体が並べ替えの基準になります。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField の小計。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|コレクション内のピボットフィールドの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|名前または id でピボットフィールドを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|名前によってピボットフィールドを取得します。 PivotField が存在しない場合は、null オブジェクトが返されます。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy の名前。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy の ID。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|名前または ID に基づいて PivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|名前に基づいて PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem の名前。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem の ID。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|PivotItem を表示するかどうかを指定します。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|コレクション内のピボットアイテムの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|名前または id で、PivotItem を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|名前によって PivotItem を取得します。 PivotItem が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange ()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|ピボットテーブルの列ラベルが存在する範囲を返します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|ピボットテーブルのデータ値が存在する範囲を返します。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|ピボットテーブルのフィルター エリアの範囲を返します。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|
||[getRowLabelRange ()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|ピボットテーブルの行ラベルが存在する範囲を返します。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|ピボットテーブルレポートに列の総計を表示するかどうかを指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|ピボットテーブルレポートで行の総計を表示するかどうかを指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|ピボットテーブルを削除します。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|ピボットテーブルのデータ ピボット階層。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottable#hierarchies)|ピボットテーブルのピボット階層。|
||[配列](/javascript/api/excel/excel.pivottable#layout)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|ピボットテーブルの行ピボット階層。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name: string, source: Range \| string \| Table, destination: range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|指定したソース データに基づくピボットテーブルを追加し、コピー先範囲の左上のセルに挿入します。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|dataValidation オブジェクトを返します。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy の ID。|
||[setToDefault ()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|RowColumnPivotHierarchy を既定値にリセットします。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|名前または ID に基づいて RowColumnPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|名前に基づいて RowColumnPivotHierarchy を取得します。 RowColumnPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[ランタイム](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|現在の作業ウィンドウまたはコンテンツアドインで JavaScript イベントを切り替えます。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース ピボット フィールド。それ以外の場合は null 値です。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース項目。それ以外の場合は null 値です。|
||[計算](/javascript/api/excel/excel.showasrule#calculation)|データ ピボット フィールドに使用する ShowAs 計算。 詳細については、「Excel ShowAsCalculation」を参照してください。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|スタイルで適用されるテキストの向き。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|automatic が true に設定されている場合、小計を設定する際に、他の値はすべて無視されます。|
||[所要](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[上限](/javascript/api/excel/excel.subtotals#max)||
||[最短](/javascript/api/excel/excel.subtotals#min)||
||[プロダクト](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[総和](/javascript/api/excel/excel.subtotals#sum)||
||[var](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|数値 id を返します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|特定のワークシート上のテーブルの変更された領域を表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|特定のワークシート上のテーブルの変更された領域を表す範囲を取得します。 null オブジェクトを返すこともあります。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|true の場合、ブックが読み取り専用モードで開かれます。 読み取り専用です。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|ワークシートが計算されるときに発生します。|
||[ショーの目盛線](/javascript/api/excel/excel.worksheet#showgridlines)|ワークシートの gridlines フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|ワークシートの headings フラグを取得または設定します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|計算対象のワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|ブック内の任意のワークシートが計算されるときに発生します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel?view=excel-js-1.8)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
