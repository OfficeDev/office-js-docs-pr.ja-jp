---
title: Excel JavaScript API 要件セット1.8
description: ExcelApi 1.8 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a5adcf56654070ca2a8336385f73062c34e90e1d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772010"
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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[策定](/javascript/api/excel/excel.chartaxisdata#alignment)|指定した軸の目盛ラベルの配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|項目の境界で数値軸が項目軸と交差するかどうかを表します。|
||[階層](/javascript/api/excel/excel.chartaxisdata#multilevel)|軸がマルチレベルかどうかを表します。|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|軸の目盛ラベルの書式コードを表します。|
||[交互](/javascript/api/excel/excel.chartaxisdata#offset)|ラベルのレベル間の距離、および先頭レベルと軸線との距離を表します。 値は 0 から 1000 の範囲内でなければなりません。|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|指定した軸と他の軸との交差位置を表します。 詳細については、「ChartAxisPosition」を参照してください。|
||[positionAt](/javascript/api/excel/excel.chartaxisdata#positionat)|指定した軸と他の軸との交差位置を表します。 このプロパティを設定するには、SetPositionAt(double) メソッドを使用する必要があります。|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|軸の目盛ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|グラフの塗りつぶしの書式設定を表します。 読み取り専用です。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[策定](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|指定した軸の目盛ラベルの配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|項目の境界で数値軸が項目軸と交差するかどうかを表します。|
||[階層](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|軸がマルチレベルかどうかを表します。|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|軸の目盛ラベルの書式コードを表します。|
||[交互](/javascript/api/excel/excel.chartaxisloadoptions#offset)|ラベルのレベル間の距離、および先頭レベルと軸線との距離を表します。 値は 0 から 1000 の範囲内でなければなりません。|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|指定した軸と他の軸との交差位置を表します。 詳細については、「ChartAxisPosition」を参照してください。|
||[positionAt](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|指定した軸と他の軸との交差位置を表します。 このプロパティを設定するには、SetPositionAt(double) メソッドを使用する必要があります。|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|軸の目盛ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|A1 スタイルの表記法を使用するグラフの軸タイトルの数式を表す文字列値。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[罫線](/javascript/api/excel/excel.chartaxistitleformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|グラフの塗りつぶしの書式設定を表します。|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[罫線](/javascript/api/excel/excel.chartaxistitleformatdata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[罫線](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[罫線](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[策定](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|指定した軸の目盛ラベルの配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|項目の境界で数値軸が項目軸と交差するかどうかを表します。|
||[階層](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|軸がマルチレベルかどうかを表します。|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|軸の目盛ラベルの書式コードを表します。|
||[交互](/javascript/api/excel/excel.chartaxisupdatedata#offset)|ラベルのレベル間の距離、および先頭レベルと軸線との距離を表します。 値は 0 から 1000 の範囲内でなければなりません。|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|指定した軸と他の軸との交差位置を表します。 詳細については、「ChartAxisPosition」を参照してください。|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|軸の目盛ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|グラフ要素の罫線の書式設定をクリアします。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|グラフがアクティブになったときに発生します。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|新しいグラフがワークシートに追加されるときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|グラフが非アクティブになったときに発生します。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|グラフが削除されるときに発生します。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|コレクション内の各アイテムについて: を参照する Chartsets Labellevel 列挙定数を取得または設定します。|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|コレクション内の各アイテムについて: グラフに空白セルをプロットする方法を設定または返します。 読み取り/書き込み可能。|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|コレクション内の各アイテムについて: グラフの plotArea を表します。|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|コレクション内の各アイテムについて、グラフのデータ系列として列または行を使用する方法を設定します。 読み取り/書き込み可能。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|コレクション内の各アイテムについて: True の場合、可視セルのみがプロットされます。false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|コレクション内の各アイテムについて: を参照する ChartSeriesNameLevel 列挙定数を設定します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|コレクション内の各アイテムについて: 値が数値軸の最大値より大きい場合にデータラベルを表示するかどうかを表します。|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|コレクション内の各項目について、次のようにします。グラフのスタイルを取得または設定します。 読み取り/書き込み可能。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|を参照する Chartsets Labellevel 列挙定数を設定または返します。|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|空白のセルがグラフでプロットされる方法を返すか設定します。 読み取り/書き込み可能。|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|グラフのプロット エリアを表します。|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|グラフ上で列または行がデータ系列として使用される方法を返すか設定します。 読み取り/書き込み可能。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartdata#plotvisibleonly)|true の場合、可視セルだけがプロットされます。false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|ChartSeriesNameLevel を参照する列挙定数を設定または返します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|
||[style](/javascript/api/excel/excel.chartdata#style)|グラフのグラフ スタイルを返すか設定します。 読み取り/書き込み可能。|
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
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[スパイク](/javascript/api/excel/excel.chartdatalabeldata#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|グラフのデータ ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[罫線](/javascript/api/excel/excel.chartdatalabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[罫線](/javascript/api/excel/excel.chartdatalabelformatdata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[罫線](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[罫線](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[スパイク](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|グラフのデータ ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|グラフのデータ ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|グラフのデータ ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフのデータ ラベルが表示されない場合は null 値となります。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[スパイク](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|グラフのデータ ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|A1 スタイルの表記法を使用するグラフのデータ ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|グラフのデータ ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|データ ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|グラフのデータ ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|グラフのデータ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|グラフのデータ ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのデータ ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[スパイク](/javascript/api/excel/excel.chartdatalabels#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|データ ラベルの書式コードを表します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[スパイク](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|データ ラベルの書式コードを表します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[スパイク](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|データ ラベルの書式コードを表します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[スパイク](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|データ ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|グラフのデータ ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|データ ラベルの書式コードを表します。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|データ ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|グラフのデータ ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
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
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|コレクション内の各アイテムについて: グラフの凡例の legendEntry の高さを表します。|
||[index](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|コレクション内の各アイテムについて: グラフの凡例の legendEntry のインデックスを表します。|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|コレクション内の各アイテムについて: グラフ legendEntry の左側を表します。|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|コレクション内の各アイテムについて: グラフ legendEntry の上部を表します。|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|コレクション内の各アイテムについて: グラフの凡例の legendEntry の幅を表します。|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|グラフの凡例に表示される凡例エントリの高さを表します。|
||[index](/javascript/api/excel/excel.chartlegendentrydata#index)|グラフの凡例に含まれる凡例エントリのインデックスを表します。|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|グラフの凡例エントリの左を表します。|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|グラフの凡例エントリの上を表します。|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|グラフの凡例に表示される凡例エントリの幅を表します。|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|グラフの凡例に表示される凡例エントリの高さを表します。|
||[index](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|グラフの凡例に含まれる凡例エントリのインデックスを表します。|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|グラフの凡例エントリの左を表します。|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|グラフの凡例エントリの上を表します。|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|グラフの凡例に表示される凡例エントリの幅を表します。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[罫線](/javascript/api/excel/excel.chartlegendformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[罫線](/javascript/api/excel/excel.chartlegendformatdata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。 読み取り専用です。|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[罫線](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[罫線](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|を参照する Chartsets Labellevel 列挙定数を設定または返します。|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|空白のセルがグラフでプロットされる方法を返すか設定します。 読み取り/書き込み可能。|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|グラフのプロット エリアを表します。|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|グラフ上で列または行がデータ系列として使用される方法を返すか設定します。 読み取り/書き込み可能。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|true の場合、可視セルだけがプロットされます。false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|ChartSeriesNameLevel を参照する列挙定数を設定または返します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|グラフのグラフ スタイルを返すか設定します。 読み取り/書き込み可能。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|プロット エリアの height 値を表します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|プロット エリアの insideHeight 値を表します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|プロット エリアの insideLeft 値を表します。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|プロット エリアの insideTop 値を表します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|プロット エリアの insideWidth 値を表します。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|プロット エリアの left 値を表します。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|プロット エリアの位置を表します。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|グラフ プロット エリアの書式設定を表します。|
||[set (properties: ChartPlotArea)](/javascript/api/excel/excel.chartplotarea#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: Officeextension.error?: Chartの場合は、オプション?: UpdateOptions)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|プロット エリアの top 値を表します。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|プロット エリアの width 値を表します。|
|[Chart/Tareadata](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|グラフ プロット エリアの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|プロット エリアの height 値を表します。|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|プロット エリアの insideHeight 値を表します。|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|プロット エリアの insideLeft 値を表します。|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|プロット エリアの insideTop 値を表します。|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|プロット エリアの insideWidth 値を表します。|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|プロット エリアの left 値を表します。|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|プロット エリアの位置を表します。|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|プロット エリアの top 値を表します。|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|プロット エリアの width 値を表します。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[罫線](/javascript/api/excel/excel.chartplotareaformat#border)|グラフ プロット エリアの罫線の属性を表します。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[set (プロパティ: Excel. Chartrelog Tareaformat)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartPlotAreaFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[Chartrelog Tareaformatdata](/javascript/api/excel/excel.chartplotareaformatdata)|[罫線](/javascript/api/excel/excel.chartplotareaformatdata#border)|グラフ プロット エリアの罫線の属性を表します。|
|[Chartrelog Tareaformatloadoptions](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[罫線](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|グラフ プロット エリアの罫線の属性を表します。|
|[ChartPlotAreaFormatUpdateData](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[罫線](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|グラフ プロット エリアの罫線の属性を表します。|
|[Chartrelog Tarealoadoptions](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|グラフ プロット エリアの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|プロット エリアの height 値を表します。|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|プロット エリアの insideHeight 値を表します。|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|プロット エリアの insideLeft 値を表します。|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|プロット エリアの insideTop 値を表します。|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|プロット エリアの insideWidth 値を表します。|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|プロット エリアの left 値を表します。|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|プロット エリアの位置を表します。|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|プロット エリアの top 値を表します。|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|プロット エリアの width 値を表します。|
|[Chartrelog Tareたデータ](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|グラフ プロット エリアの書式設定を表します。|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|プロット エリアの height 値を表します。|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|プロット エリアの insideHeight 値を表します。|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|プロット エリアの insideLeft 値を表します。|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|プロット エリアの insideTop 値を表します。|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|プロット エリアの insideWidth 値を表します。|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|プロット エリアの left 値を表します。|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|プロット エリアの位置を表します。|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|プロット エリアの top 値を表します。|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|プロット エリアの width 値を表します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|指定した系列のグループを取得または設定します。 読み取り/書き込み|
||[切り出し](/javascript/api/excel/excel.chartseries#explosion)|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|横棒と縦棒の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|系列に含まれるすべてのデータ ラベルのコレクションを表します。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[Chartcharts Collectionloadoptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|コレクション内の各アイテムについて: 指定されたデータ系列のグループを取得または設定します。 読み取り/書き込み|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|コレクション内の各アイテムについて: 系列内のすべての dataLabels のコレクションを表します。|
||[切り出し](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|コレクション内の各アイテムについて: 円グラフまたはドーナツグラフのスライスの切り出し値を取得または設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|コレクション内の各項目について、次のようにします。最初の円グラフまたはドーナツグラフのスライスの角度を、縦から90°の角度で取得または設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|コレクション内の各項目について: True の場合、Microsoft Excel は負の値に対応するときにパターンを反転します。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|コレクション内の各アイテムについて: バーと列の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|コレクション内の各アイテムについて: 補助円グラフ付き円グラフまたは補助縦棒付き円グラフのいずれかの第2セクションのサイズを、プライマリ円グラフのサイズに対する割合で設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|コレクション内の各アイテムについて: 補助円グラフ付き円グラフまたは補助縦棒付き円グラフのいずれかの2つのセクションを分割する方法を設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|コレクション内の各項目について: True の場合、Microsoft Excel は各データマーカーに異なる色またはパターンを割り当てます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[Chart系列データ](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|指定した系列のグループを取得または設定します。 読み取り/書き込み|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|系列に含まれるすべてのデータ ラベルのコレクションを表します。|
||[切り出し](/javascript/api/excel/excel.chartseriesdata#explosion)|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseriesdata#overlap)|横棒と縦棒の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[Chart系列 Loadoptions](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|指定した系列のグループを取得または設定します。 読み取り/書き込み|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|系列に含まれるすべてのデータ ラベルのコレクションを表します。|
||[切り出し](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|横棒と縦棒の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|指定した系列のグループを取得または設定します。 読み取り/書き込み|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|系列に含まれるすべてのデータ ラベルのコレクションを表します。|
||[切り出し](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|円グラフまたはドーナツ グラフのスライス切り出し表示の値を返すか設定します。 切り出し表示は行われず、スライスの先端が円の中心と一致する場合、0 を返します。 読み取り/書き込み可能。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|円グラフまたはドーナツ グラフの最初のスライスの角度 (縦の中心から時計回りでの度数) を返すか設定します。 円グラフ、3-D 円グラフ、およびドーナツ グラフにのみ適用されます。 0 から 360 の範囲内で値を指定できます。 読み取り/書き込み|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|true の場合、Microsoft Excel により、負の数値に対応する項目でパターンが反転されます。 読み取り/書き込み可能。|
||[overlap](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|横棒と縦棒の配置方法を指定します。 -100 から100までの値を指定できます。 2-D 横棒グラフと 2-D 縦棒グラフにのみ適用されます。 読み取り/書き込み可能。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフのセカンダリ セクションのサイズを、プライマリ セクションのサイズのパーセンテージとして返すか設定します。 5 から 200 の範囲内で値を指定できます。 読み取り/書き込み可能。|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|補助円グラフ付き円グラフまたは補助縦棒グラフ付き円グラフを 2 つの部分に分割する方法を返すか設定します。 読み取り/書き込み可能。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|true の場合、Microsoft Excel により、データ マーカーごとに異なる色またはパターンが割り当てられます。 グラフに含まれるデータ系列は 1 つだけでなければなりません。 読み取り/書き込み可能。|
|[ChartTrendline 曲線](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendline#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendline#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|コレクション内の各アイテムについて: 近似曲線が後方に延長される期間を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|コレクション内の各アイテムについて: 近似曲線を前方に延長する時間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|コレクション内の各アイテムについて: グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|コレクション内の各アイテムについて: True の場合、近似曲線の数式がグラフに表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|コレクション内の各項目についての場合: True の場合、グラフに近似曲線の R-2 乗が表示されます。|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendlinedata#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendlinedata#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[スパイク](/javascript/api/excel/excel.charttrendlinelabel#autotext)|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|グラフの近似曲線ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|近似曲線ラベルの書式コードを表す文字列値。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|グラフの近似曲線ラベルの書式設定を表します。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[set (properties: ChartTrendlineLabel)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTrendlineLabelUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|グラフの近似曲線ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[スパイク](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|グラフの近似曲線ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|グラフの近似曲線ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|近似曲線ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|グラフの近似曲線ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[罫線](/javascript/api/excel/excel.charttrendlinelabelformat#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|現在のグラフの近似曲線ラベルの塗りつぶしの書式設定を表します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
||[set (properties: ChartTrendlineLabelFormat)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTrendlineLabelFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartTrendlineLabelFormatData](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[罫線](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartTrendlineLabelFormatLoadOptions](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[罫線](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartTrendlineLabelFormatUpdateData](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[罫線](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|グラフの罫線の書式設定 (色、線のスタイル、線の太さなど) を表します。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|グラフの近似曲線ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[スパイク](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|グラフの近似曲線ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|グラフの近似曲線ラベルの高さ (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|グラフの近似曲線ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|近似曲線ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|グラフの近似曲線ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|グラフの近似曲線ラベルの幅 (ポイント数) を返します。 読み取り専用です。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[スパイク](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|近似曲線ラベルでコンテキストに基づく適切なテキストを自動的に生成するかどうかを表すブール値。|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|グラフの近似曲線ラベルの書式設定を表します。|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|A1 スタイルの表記法を使用するグラフの近似曲線ラベルの数式を表す文字列値。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|グラフの近似曲線ラベルの水平方向の配置を表します。 詳細については、「ChartTextHorizontalAlignment」を参照してください。|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|グラフの近似曲線ラベルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|近似曲線ラベルの書式コードを表す文字列値。|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|グラフの近似曲線ラベルのテキストを表す文字列。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|グラフの近似曲線ラベルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|グラフの近似曲線ラベルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフの近似曲線ラベルが表示されない場合は null 値となります。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|グラフの近似曲線ラベルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendlineloadoptions#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|近似曲線を後方へ拡張するときの区間数を表します。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|近似曲線を前方へ拡張するときの区間数を表します。|
||[付け](/javascript/api/excel/excel.charttrendlineupdatedata#label)|グラフの近似曲線のラベルを表します。|
||[showequ](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|true の場合、グラフに近似曲線の数式が表示されます。|
||[showRSquared](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|true の場合、グラフに近似曲線の R-2 乗値が表示されます。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|を参照する Chartsets Labellevel 列挙定数を設定または返します。|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|空白のセルがグラフでプロットされる方法を返すか設定します。 読み取り/書き込み可能。|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|グラフのプロット エリアを表します。|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|グラフ上で列または行がデータ系列として使用される方法を返すか設定します。 読み取り/書き込み可能。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|true の場合、可視セルだけがプロットされます。false の場合、可視セルと非表示セルの両方がプロットされます。 読み取り/書き込み可能。|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|ChartSeriesNameLevel を参照する列挙定数を設定または返します。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|値が数値軸の最大値より大きい場合にデータ ラベルを表示するかどうかを表します。|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|グラフのグラフ スタイルを返すか設定します。 読み取り/書き込み可能。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|ユーザーの入力規則のカスタム数式。 これにより、重複を防止したり、セル範囲内の合計を制限したりする特別な入力ルールが作成されます。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy の位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy の ID。|
||[set (properties: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: DataPivotHierarchyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setToDefault ()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|DataPivotHierarchy を既定値にリセットします。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|データを特定の集計計算として表示するかどうかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|名前または ID に基づいて DataPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|名前に基づいて DataPivotHierarchy を取得します。 DataPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|現在の軸からピボット階層を削除します。|
|[DataPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|コレクション内の各アイテムについて: DataPivotHierarchy に関連付けられているピボットフィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|コレクション内の各アイテムについて: DataPivotHierarchy の Id。|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|コレクション内の各項目について: DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|コレクション内の各項目について: DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|コレクション内の各アイテムについて: DataPivotHierarchy の位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|コレクション内の各アイテムについて: データを特定のサマリー計算として表示するかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|コレクション内の各アイテムについて: DataPivotHierarchy のすべてのアイテムを表示するかどうかを指定します。|
|[DataPivotHierarchyData](/javascript/api/excel/excel.datapivothierarchydata)|[field](/javascript/api/excel/excel.datapivothierarchydata#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|DataPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|DataPivotHierarchy の位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|データを特定の集計計算として表示するかどうかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|
|[DataPivotHierarchyLoadOptions](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|DataPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|DataPivotHierarchy の位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|データを特定の集計計算として表示するかどうかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|
|[DataPivotHierarchyUpdateData](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[field](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|DataPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|DataPivotHierarchy の名前。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|DataPivotHierarchy の数値形式。|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|DataPivotHierarchy の位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|データを特定の集計計算として表示するかどうかどうかを指定します。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|DataPivotHierarchy のすべての項目を表示するかどうかを指定します。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|現在の範囲からデータの入力規則をクリアします。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|
||[Prompt](/javascript/api/excel/excel.datavalidation#prompt)|ユーザーがセルを選択したときにメッセージを表示します。|
||[type](/javascript/api/excel/excel.datavalidation#type)|データの入力規則の種類。詳細については、Excel.DataValidationType を参照してください。|
||[有効な](/javascript/api/excel/excel.datavalidation#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
||[除外](/javascript/api/excel/excel.datavalidation#rule)|さまざまな種類のデータ検証条件を含むデータ入力規則。|
||[set (properties: DataValidation)](/javascript/api/excel/excel.datavalidation#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: DataValidationUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[DataValidationData](/javascript/api/excel/excel.datavalidationdata)|[errorAlert](/javascript/api/excel/excel.datavalidationdata#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|
||[Prompt](/javascript/api/excel/excel.datavalidationdata#prompt)|ユーザーがセルを選択したときにメッセージを表示します。|
||[除外](/javascript/api/excel/excel.datavalidationdata#rule)|さまざまな種類のデータ検証条件を含むデータ入力規則。|
||[type](/javascript/api/excel/excel.datavalidationdata#type)|データの入力規則の種類。詳細については、Excel.DataValidationType を参照してください。|
||[有効な](/javascript/api/excel/excel.datavalidationdata#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|エラー警告メッセージを表します。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|無効なデータが入力されたときにエラー警告ダイアログを表示するかどうかを指定します。 既定値は true です。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|データの入力規則に対する警告の種類を表します。詳細については、Excel.DataValidationAlertStyle を参照してください。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|エラー警告ダイアログのタイトルを表します。|
|[DataValidationLoadOptions](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[errorAlert](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|
||[Prompt](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|ユーザーがセルを選択したときにメッセージを表示します。|
||[除外](/javascript/api/excel/excel.datavalidationloadoptions#rule)|さまざまな種類のデータ検証条件を含むデータ入力規則。|
||[type](/javascript/api/excel/excel.datavalidationloadoptions#type)|データの入力規則の種類。詳細については、Excel.DataValidationType を参照してください。|
||[有効な](/javascript/api/excel/excel.datavalidationloadoptions#valid)|すべてのセルの値がデータの入力規則に従っているかどうかを表します。|
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
|[DataValidationUpdateData](/javascript/api/excel/excel.datavalidationupdatedata)|[errorAlert](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|無効なデータが入力された場合のエラー警告。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|空白を無視します。つまり、空白のセルではデータの入力規則が検証されません。既定では true に設定されます。|
||[Prompt](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|ユーザーがセルを選択したときにメッセージを表示します。|
||[除外](/javascript/api/excel/excel.datavalidationupdatedata#rule)|さまざまな種類のデータ検証条件を含むデータ入力規則。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Operator プロパティが GreaterThan などのバイナリ演算子に設定されている場合、右のオペランドを指定します (左のオペランドは、ユーザーがセルに入力しようとした値です)。 三項演算子と NotBetween の間の and notbetween で、下限オペランドを指定します。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|三項演算子を and NotBetween で指定すると、上限オペランドが指定されます。 は、GreaterThan などの二項演算子では使用されません。|
||[演算子](/javascript/api/excel/excel.datetimedatavalidation#operator)|データの検証に使用する演算子。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy の ID。|
||[set (properties: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: FilterPivotHierarchyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setToDefault ()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|FilterPivotHierarchy を既定値にリセットします。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|名前または ID に基づいて FilterPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|名前に基づいて FilterPivotHierarchy を取得します。 FilterPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[FilterPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|コレクション内の各アイテムについて: 複数のフィルターアイテムを許可するかどうかを指定します。|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|コレクション内の各アイテムについて: FilterPivotHierarchy の Id。|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|コレクション内の各項目について: FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|コレクション内の各アイテムについて: FilterPivotHierarchy の位置。|
|[FilterPivotHierarchyData](/javascript/api/excel/excel.filterpivothierarchydata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|FilterPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|FilterPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|FilterPivotHierarchy の位置。|
|[FilterPivotHierarchyLoadOptions](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|FilterPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|FilterPivotHierarchy の位置。|
|[FilterPivotHierarchyUpdateData](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|複数のフィルター項目を許可するかどうかを指定します。|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|FilterPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|FilterPivotHierarchy の位置。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|セルのドロップダウンにリストを表示するかどうかを指定します。既定では true に設定されます。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|データの入力規則のリストのソース。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField の名前。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField の ID。|
||[items](/javascript/api/excel/excel.pivotfield#items)|PivotField で構成される PivotItems を返します。|
||[set (properties: Excel. PivotField)](/javascript/api/excel/excel.pivotfield#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotFieldUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField を並べ替えます。 DataPivotHierarchy を指定すると、そのピボット階層に基づいて並べ替えが適用されます。指定しない場合、ピボット フィールド自体が並べ替えの基準になります。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField の小計。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|コレクション内のピボットフィールドの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|名前または id でピボットフィールドを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|名前によってピボットフィールドを取得します。 PivotField が存在しない場合は、null オブジェクトが返されます。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotFieldCollectionLoadOptions](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|コレクション内の各アイテムについて: PivotField の Id。|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|コレクション内の各アイテムについて: PivotField の名前。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|コレクション内の各アイテムについて: ピボットフィールドのすべてのアイテムを表示するかどうかを指定します。|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|コレクション内の各アイテムについて: PivotField の小計。|
|[PivotFieldData](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|PivotField の ID。|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|PivotField に関連付けられているピボット フィールドを返します。|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|PivotField の名前。|
||[showAllItems](/javascript/api/excel/excel.pivotfielddata#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|PivotField の小計。|
|[PivotFieldLoadOptions](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|PivotField の ID。|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|PivotField の名前。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|PivotField の小計。|
|[PivotFieldUpdateData](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|PivotField の名前。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|PivotField のすべての項目を表示するかどうかを指定します。|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|PivotField の小計。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy の名前。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy の ID。|
||[set (properties: PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotHierarchyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|名前または ID に基づいて PivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|名前に基づいて PivotHierarchy を取得します。 PivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|コレクション内の各アイテムについて: PivotHierarchy の Id。|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|コレクション内の各項目について: PivotHierarchy の名前。|
|[PivotHierarchyData](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|PivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|PivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|PivotHierarchy の名前。|
|[PivotHierarchyLoadOptions](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|PivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|PivotHierarchy の名前。|
|[PivotHierarchyUpdateData](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|PivotHierarchy の名前。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem の名前。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem の ID。|
||[set (プロパティ: Excel. PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotItemUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|PivotItem を表示するかどうかを指定します。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|コレクション内のピボットアイテムの数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|名前または id で、PivotItem を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|名前によって PivotItem を取得します。 PivotItem が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotItemCollectionLoadOptions](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|コレクション内の各アイテムについて: PivotItem の Id。|
||[isExpanded](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|コレクション内の各アイテムについて: 子アイテムを表示するようにアイテムが展開されているかどうか、または折りたたまれて子アイテムが非表示になっているかどうかを決定します。|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|コレクション内の各アイテムについて: PivotItem の名前。|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|コレクション内の各アイテムについて: PivotItem を表示するかどうかを指定します。|
|[PivotItemData](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|PivotItem の ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitemdata#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|PivotItem の名前。|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|PivotItem を表示するかどうかを指定します。|
|[PivotItemLoadOptions](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|PivotItem の ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|PivotItem の名前。|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|PivotItem を表示するかどうかを指定します。|
|[PivotItemUpdateData](/javascript/api/excel/excel.pivotitemupdatedata)|[isExpanded](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|項目を展開して子項目を表示するか、または項目を折りたたんで子項目を非表示にするかを指定します。|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|PivotItem の名前。|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|PivotItem を表示するかどうかを指定します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange ()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|ピボットテーブルの列ラベルが存在する範囲を返します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|ピボットテーブルのデータ値が存在する範囲を返します。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|ピボットテーブルのフィルター エリアの範囲を返します。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|フィルター エリアを除く、ピボットテーブルが存在する範囲を返します。|
||[getRowLabelRange ()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|ピボットテーブルの行ラベルが存在する範囲を返します。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|
||[set (properties: PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotLayoutUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|ピボットテーブルレポートに列の総計を表示するかどうかを指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|ピボットテーブルレポートで行の総計を表示するかどうかを指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[layoutType](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|ピボットテーブルレポートに列の総計を表示するかどうかを指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|ピボットテーブルレポートで行の総計を表示するかどうかを指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[layoutType](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|ピボットテーブルレポートに列の総計を表示するかどうかを指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|ピボットテーブルレポートで行の総計を表示するかどうかを指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[layoutType](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|このプロパティは、ピボットテーブルのすべてのフィールドの PivotLayoutType を示します。 フィールドによって状態が異なる場合は null 値になります。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|ピボットテーブルレポートに列の総計を表示するかどうかを指定します。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|ピボットテーブルレポートで行の総計を表示するかどうかを指定します。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|このプロパティは、ピボットテーブルのすべてのフィールドの SubtotalLocationType を示します。 フィールドによって状態が異なる場合は null 値になります。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|ピボットテーブルを削除します。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|ピボットテーブルのデータ ピボット階層。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottable#hierarchies)|ピボットテーブルのピボット階層。|
||[配列](/javascript/api/excel/excel.pivottable#layout)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|ピボットテーブルの行ピボット階層。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name: string, source: Range \| string \| Table, destination: range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|指定したソース データに基づくピボットテーブルを追加し、コピー先範囲の左上のセルに挿入します。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[配列](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|コレクション内の各アイテムについて: PivotLayout。ピボットテーブルのレイアウトと視覚的な構造を説明します。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[columnHierarchies](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|ピボットテーブルの列ピボット階層。|
||[dataHierarchies](/javascript/api/excel/excel.pivottabledata#datahierarchies)|ピボットテーブルのデータ ピボット階層。|
||[filterHierarchies](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|ピボットテーブルのフィルター ピボット階層。|
||[階層](/javascript/api/excel/excel.pivottabledata#hierarchies)|ピボットテーブルのピボット階層。|
||[rowHierarchies](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|ピボットテーブルの行ピボット階層。|
|[ピボットのオプション](/javascript/api/excel/excel.pivottableloadoptions)|[配列](/javascript/api/excel/excel.pivottableloadoptions#layout)|ピボットテーブルのレイアウトとビジュアル構造を記述する PivotLayout。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|dataValidation オブジェクトを返します。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|dataValidation オブジェクトを返します。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|dataValidation オブジェクトを返します。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|dataValidation オブジェクトを返します。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy の位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy の ID。|
||[set (properties: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RowColumnPivotHierarchyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setToDefault ()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|RowColumnPivotHierarchy を既定値にリセットします。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|現在の軸にピボット階層を追加します。 階層が行、列、またはフィルター軸の他の場所に存在する場合は、その場所から削除されます。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|コレクションに含まれるピボット階層の数を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|名前または ID に基づいて RowColumnPivotHierarchy を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|名前に基づいて RowColumnPivotHierarchy を取得します。 RowColumnPivotHierarchy が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove (rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|現在の軸からピボット階層を削除します。|
|[RowColumnPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|コレクション内の各アイテムについて: RowColumnPivotHierarchy の Id。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|コレクション内の各項目について: RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|コレクション内の各アイテムについて: RowColumnPivotHierarchy の位置。|
|[RowColumnPivotHierarchyData](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|RowColumnPivotHierarchy に関連付けられているピボット フィールドを返します。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|RowColumnPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|RowColumnPivotHierarchy の位置。|
|[RowColumnPivotHierarchyLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|RowColumnPivotHierarchy の ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|RowColumnPivotHierarchy の位置。|
|[RowColumnPivotHierarchyUpdateData](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|RowColumnPivotHierarchy の名前。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|RowColumnPivotHierarchy の位置。|
|[ランタイム](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|現在の作業ウィンドウまたはコンテンツアドインで JavaScript イベントを切り替えます。|
|[RuntimeData](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|現在の作業ウィンドウまたはコンテンツアドインで JavaScript イベントを切り替えます。|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|現在の作業ウィンドウまたはコンテンツアドインで JavaScript イベントを切り替えます。|
|[RuntimeUpdateData](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|現在の作業ウィンドウまたはコンテンツアドインで JavaScript イベントを切り替えます。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース ピボット フィールド。それ以外の場合は null 値です。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|ShowAsCalculation 型に基づき、該当する場合は ShowAs 計算の基準となるベース項目。それ以外の場合は null 値です。|
||[計算](/javascript/api/excel/excel.showasrule#calculation)|データ ピボット フィールドに使用する ShowAs 計算。 詳細については、「Excel ShowAsCalculation」を参照してください。|
|[スタイル](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|スタイルで適用されるテキストの向き。|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|コレクション内の各アイテムについて: セル内のテキストの配置が均等分布に設定されている場合に、文字列を自動的にインデントするかどうかを示します。|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|コレクション内の各アイテムについて: スタイルのテキストの向き。|
|[スタイルデータ](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|スタイルで適用されるテキストの向き。|
|[スタイル Loadオプション](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|スタイルで適用されるテキストの向き。|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|セル内のテキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|スタイルで適用されるテキストの向き。|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|コレクション内の各アイテムについて: 数値 id を返します。|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|数値 id を返します。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|数値 id を返します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|true の場合、ブックが読み取り専用モードで開かれます。 読み取り専用です。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[WorkbookData](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|true の場合、ブックが読み取り専用モードで開かれます。 読み取り専用です。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|true の場合、ブックが読み取り専用モードで開かれます。 読み取り専用です。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|ワークシートが計算されるときに発生します。|
||[ショーの目盛線](/javascript/api/excel/excel.worksheet#showgridlines)|ワークシートの gridlines フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|ワークシートの headings フラグを取得または設定します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|計算対象のワークシートの ID を取得します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|ブック内の任意のワークシートが計算されるときに発生します。|
|[ワークシート Collectionloadoptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[ショーの目盛線](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|コレクション内の各アイテムについて: ワークシートの枠線フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|コレクション内の各アイテムについて: ワークシートの見出しフラグを取得または設定します。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[ショーの目盛線](/javascript/api/excel/excel.worksheetdata#showgridlines)|ワークシートの gridlines フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheetdata#showheadings)|ワークシートの headings フラグを取得または設定します。|
|[ワークシート Loadoptions](/javascript/api/excel/excel.worksheetloadoptions)|[ショーの目盛線](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|ワークシートの gridlines フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|ワークシートの headings フラグを取得または設定します。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[ショーの目盛線](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|ワークシートの gridlines フラグを取得または設定します。|
||[showHeadings](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|ワークシートの headings フラグを取得または設定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
