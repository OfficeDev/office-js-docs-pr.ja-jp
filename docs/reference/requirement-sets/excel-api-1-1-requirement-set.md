---
title: Excel JavaScript API 要件セット1.1
description: ExcelApi 1.1 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 921a67b4242150d767fdac057d21c6fc510d98b3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772052"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要件セット1.1

Excel JavaScript API 1.1 は、API の最初のバージョンです。 Excel 2016 でサポートされている唯一の Excel 固有の要件セットです。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[計算 (電卓 Ationtype: "再計算\| " "full \| rebuild")](/javascript/api/excel/excel.application#calculate-calculationtype-)|Excel で現在開いているすべてのブックを再計算します。|
||[計算 (電卓 Ationtype: Excel. 電卓 Ationtype)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Excel で現在開いているすべてのブックを再計算します。|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。 使用可能な値`Automatic`は次のとおりです。 Excel では、再計算が制御されます。`AutomaticExceptTables`、Excel は再計算を制御しますが、テーブル内の変更は無視します。`Manual`、ユーザーが要求すると、計算が行われます。|
||[set (properties: Excel)](/javascript/api/excel/excel.application#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ApplicationUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.application#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationMode](/javascript/api/excel/excel.applicationdata#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。 使用可能な値`Automatic`は次のとおりです。 Excel では、再計算が制御されます。`AutomaticExceptTables`、Excel は再計算を制御しますが、テーブル内の変更は無視します。`Manual`、ユーザーが要求すると、計算が行われます。|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[$all](/javascript/api/excel/excel.applicationloadoptions#$all)||
||[calculationMode](/javascript/api/excel/excel.applicationloadoptions#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。 使用可能な値`Automatic`は次のとおりです。 Excel では、再計算が制御されます。`AutomaticExceptTables`、Excel は再計算を制御しますが、テーブル内の変更は無視します。`Manual`、ユーザーが要求すると、計算が行われます。|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[calculationMode](/javascript/api/excel/excel.applicationupdatedata#calculationmode)|CalculationMode の定数によって定義されている、ブックで使用されている計算モードを返します。 使用可能な値`Automatic`は次のとおりです。 Excel では、再計算が制御されます。`AutomaticExceptTables`、Excel は再計算を制御しますが、テーブル内の変更は無視します。`Manual`、ユーザーが要求すると、計算が行われます。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|バインディングによって表される範囲を返します。バインドが正しい型ではない場合、エラーがスローされます。|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|バインドによって表されるテーブルを返します。バインドが正しい型ではない場合、エラーがスローされます。|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|バインドによって表されるテキストを返します。 バインドが正しい型ではない場合、エラーがスローされます。|
||[id](/javascript/api/excel/excel.binding#id)|バインド識別子を表します。 読み取り専用。|
||[type](/javascript/api/excel/excel.binding#type)|バインドの種類を返します。 詳細については、「Excel. BindingType」を参照してください。 読み取り専用です。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|ID によってバインド オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|項目の配列内の位置に基づいて、バインド オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|コレクション内にあるバインドの数を取得します。 値の取得のみ可能です。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[BindingCollectionLoadOptions](/javascript/api/excel/excel.bindingcollectionloadoptions)|[$all](/javascript/api/excel/excel.bindingcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingcollectionloadoptions#id)|コレクション内の各アイテムについて: バインド識別子を表します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.bindingcollectionloadoptions#type)|コレクション内の各アイテムについて: バインドの種類を返します。 詳細については、「Excel. BindingType」を参照してください。 読み取り専用です。|
|[BindingData](/javascript/api/excel/excel.bindingdata)|[id](/javascript/api/excel/excel.bindingdata#id)|バインド識別子を表します。 読み取り専用。|
||[type](/javascript/api/excel/excel.bindingdata#type)|バインドの種類を返します。 詳細については、「Excel. BindingType」を参照してください。 読み取り専用です。|
|[BindingLoadOptions](/javascript/api/excel/excel.bindingloadoptions)|[$all](/javascript/api/excel/excel.bindingloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingloadoptions#id)|バインド識別子を表します。 読み取り専用。|
||[type](/javascript/api/excel/excel.bindingloadoptions#type)|バインドの種類を返します。 詳細については、「Excel. BindingType」を参照してください。 読み取り専用です。|
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
||[set (properties: Excel. Chart)](/javascript/api/excel/excel.chart#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chart#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setData (sourceData: Range、系列 By?: "Auto" \| "列" \| "Rows")](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|グラフの元データをリセットします。|
||[setData (sourceData: Range、系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|グラフの元データをリセットします。|
||[setPosition (startCell: Range \| String, endcell?: 範囲\|文字列)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|ワークシート上のセルを基準にしてグラフを配置します。|
||[top](/javascript/api/excel/excel.chart#top)|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.chart#width)|グラフ オブジェクトの幅をポイント単位で表します。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[set (プロパティ: Excel. ChartAreaFormat)](/javascript/api/excel/excel.chartareaformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAreaFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartareaformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[font](/javascript/api/excel/excel.chartareaformatdata#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartareaformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartareaformatloadoptions#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[font](/javascript/api/excel/excel.chartareaformatupdatedata#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|グラフの項目軸を表します。 読み取り専用です。|
||[系列軸](/javascript/api/excel/excel.chartaxes#seriesaxis)|3 次元グラフの系列軸を表します。 読み取り専用です。|
||[数値軸](/javascript/api/excel/excel.chartaxes#valueaxis)|軸の数値軸を表します。 読み取り専用です。|
||[set (properties: Excel. ChartAxes)](/javascript/api/excel/excel.chartaxes#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAxesUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartaxes#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartAxesData](/javascript/api/excel/excel.chartaxesdata)|[categoryAxis](/javascript/api/excel/excel.chartaxesdata#categoryaxis)|グラフの項目軸を表します。 読み取り専用です。|
||[系列軸](/javascript/api/excel/excel.chartaxesdata#seriesaxis)|3 次元グラフの系列軸を表します。 読み取り専用です。|
||[数値軸](/javascript/api/excel/excel.chartaxesdata#valueaxis)|軸の数値軸を表します。 読み取り専用です。|
|[ChartAxesLoadOptions](/javascript/api/excel/excel.chartaxesloadoptions)|[$all](/javascript/api/excel/excel.chartaxesloadoptions#$all)||
||[categoryAxis](/javascript/api/excel/excel.chartaxesloadoptions#categoryaxis)|グラフの項目軸を表します。|
||[系列軸](/javascript/api/excel/excel.chartaxesloadoptions#seriesaxis)|3 次元グラフの系列軸を表します。|
||[数値軸](/javascript/api/excel/excel.chartaxesloadoptions#valueaxis)|軸の数値軸を表します。|
|[ChartAxesUpdateData](/javascript/api/excel/excel.chartaxesupdatedata)|[categoryAxis](/javascript/api/excel/excel.chartaxesupdatedata#categoryaxis)|グラフの項目軸を表します。|
||[系列軸](/javascript/api/excel/excel.chartaxesupdatedata#seriesaxis)|3 次元グラフの系列軸を表します。|
||[数値軸](/javascript/api/excel/excel.chartaxesupdatedata#valueaxis)|軸の数値軸を表します。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|数値軸の最小値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|2 つの小さい目盛の間隔を表します。 数値の値または空の文字列を設定できます (軸の値が自動の場合)。 戻り値は常に数値です。|
||[format](/javascript/api/excel/excel.chartaxis#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。 読み取り専用です。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。 値の取得のみ可能です。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。 読み取り専用です。|
||[title](/javascript/api/excel/excel.chartaxis#title)|軸タイトルを表します。 読み取り専用です。|
||[set (properties: ChartAxis)](/javascript/api/excel/excel.chartaxis#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAxisUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartaxis#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[format](/javascript/api/excel/excel.chartaxisdata#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。 読み取り専用です。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisdata#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。 値の取得のみ可能です。|
||[majorUnit](/javascript/api/excel/excel.chartaxisdata#majorunit)|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
||[maximum](/javascript/api/excel/excel.chartaxisdata#maximum)|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minimum](/javascript/api/excel/excel.chartaxisdata#minimum)|数値軸の最小値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisdata#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。 読み取り専用です。|
||[minorUnit](/javascript/api/excel/excel.chartaxisdata#minorunit)|2 つの小さい目盛の間隔を表します。 数値の値または空の文字列を設定できます (軸の値が自動の場合)。 戻り値は常に数値です。|
||[title](/javascript/api/excel/excel.chartaxisdata#title)|軸タイトルを表します。 読み取り専用です。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|グラフの線の書式設定を表します。 読み取り専用です。|
||[set (properties: ChartAxisFormat)](/javascript/api/excel/excel.chartaxisformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAxisFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartaxisformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartAxisFormatData](/javascript/api/excel/excel.chartaxisformatdata)|[font](/javascript/api/excel/excel.chartaxisformatdata#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[line](/javascript/api/excel/excel.chartaxisformatdata#line)|グラフの線の書式設定を表します。 読み取り専用です。|
|[ChartAxisFormatLoadOptions](/javascript/api/excel/excel.chartaxisformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxisformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxisformatloadoptions#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を表します。|
||[line](/javascript/api/excel/excel.chartaxisformatloadoptions#line)|グラフの線の書式設定を表します。|
|[ChartAxisFormatUpdateData](/javascript/api/excel/excel.chartaxisformatupdatedata)|[font](/javascript/api/excel/excel.chartaxisformatupdatedata#font)|グラフ軸要素のフォント属性 (フォント名、フォント サイズ、色など) を表します。|
||[line](/javascript/api/excel/excel.chartaxisformatupdatedata#line)|グラフの線の書式設定を表します。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[$all](/javascript/api/excel/excel.chartaxisloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxisloadoptions#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。|
||[majorUnit](/javascript/api/excel/excel.chartaxisloadoptions#majorunit)|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
||[maximum](/javascript/api/excel/excel.chartaxisloadoptions#maximum)|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minimum](/javascript/api/excel/excel.chartaxisloadoptions#minimum)|数値軸の最小値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。|
||[minorUnit](/javascript/api/excel/excel.chartaxisloadoptions#minorunit)|2 つの小さい目盛の間隔を表します。 数値の値または空の文字列を設定できます (軸の値が自動の場合)。 戻り値は常に数値です。|
||[title](/javascript/api/excel/excel.chartaxisloadoptions#title)|軸タイトルを表します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|グラフ軸のタイトルの書式設定を表します。 読み取り専用です。|
||[set (properties: ChartAxisTitle)](/javascript/api/excel/excel.chartaxistitle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAxisTitleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartaxistitle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|軸タイトルを表します。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|軸のタイトルの表示/非表示を指定するブール型の値です。|
|[ChartAxisTitleData](/javascript/api/excel/excel.chartaxistitledata)|[format](/javascript/api/excel/excel.chartaxistitledata#format)|グラフ軸のタイトルの書式設定を表します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.chartaxistitledata#text)|軸タイトルを表します。|
||[visible](/javascript/api/excel/excel.chartaxistitledata#visible)|軸のタイトルの表示/非表示を指定するブール型の値です。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|グラフの軸タイトルのオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[set (properties: ChartAxisTitleFormat)](/javascript/api/excel/excel.chartaxistitleformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartAxisTitleFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartaxistitleformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[font](/javascript/api/excel/excel.chartaxistitleformatdata#font)|グラフの軸タイトルのオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxistitleformatloadoptions#font)|グラフの軸タイトルのオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[font](/javascript/api/excel/excel.chartaxistitleformatupdatedata#font)|グラフの軸タイトルのオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartAxisTitleLoadOptions](/javascript/api/excel/excel.chartaxistitleloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxistitleloadoptions#format)|グラフ軸のタイトルの書式設定を表します。|
||[text](/javascript/api/excel/excel.chartaxistitleloadoptions#text)|軸タイトルを表します。|
||[visible](/javascript/api/excel/excel.chartaxistitleloadoptions#visible)|軸のタイトルの表示/非表示を指定するブール型の値です。|
|[ChartAxisTitleUpdateData](/javascript/api/excel/excel.chartaxistitleupdatedata)|[format](/javascript/api/excel/excel.chartaxistitleupdatedata#format)|グラフ軸のタイトルの書式設定を表します。|
||[text](/javascript/api/excel/excel.chartaxistitleupdatedata#text)|軸タイトルを表します。|
||[visible](/javascript/api/excel/excel.chartaxistitleupdatedata#visible)|軸のタイトルの表示/非表示を指定するブール型の値です。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[format](/javascript/api/excel/excel.chartaxisupdatedata#format)|グラフオブジェクトの書式を表します。これには、行とフォントの書式設定が含まれます。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#majorgridlines)|指定された軸の大きい目盛線を表す Gridlines オブジェクトを返します。|
||[majorUnit](/javascript/api/excel/excel.chartaxisupdatedata#majorunit)|2 つの大きい目盛の間隔を表します。数値の値または空の文字列を設定できます。戻り値は常に数値です。|
||[maximum](/javascript/api/excel/excel.chartaxisupdatedata#maximum)|数値軸の最大値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minimum](/javascript/api/excel/excel.chartaxisupdatedata#minimum)|数値軸の最小値を表します。数値の値または空の文字列を設定できます (軸の値が自動の場合)。戻り値は常に数値です。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#minorgridlines)|指定された軸の小さい目盛線を表す gridlines オブジェクトを返します。|
||[minorUnit](/javascript/api/excel/excel.chartaxisupdatedata#minorunit)|2 つの小さい目盛の間隔を表します。 数値の値または空の文字列を設定できます (軸の値が自動の場合)。 戻り値は常に数値です。|
||[title](/javascript/api/excel/excel.chartaxisupdatedata#title)|軸タイトルを表します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add (type: "Invalid" \| "columnclustered" \| "columnclustered" \| "ColumnStacked100" \| "3dcolumnclustered" \| "3dcolumnclustered" \| "3DColumnStacked100" \| "barclustered" \| "barclustered"\| "BarStacked100" \| "3dbarclustered" \| "3dbarclustered" \| "3DBarStacked100" \| "LineStacked" \| "LineStacked100" \| "LineMarkers" \| "LineMarkersStacked" \| "LineMarkersStacked100 " \| " PieOfPie " \| " PieExploded " \| " 3DPieExploded " \| " barofpie " \| " XYScatterSmooth " \| " XYScatterSmoothNoMarkers " \| " XYScatterLines " \| "XYScatterLinesNoMarkers " \| \| " がいっぱいになってい\|ます "AreaStacked100" " \| 3dアレス積み上げ" "3dareastac3d100" \| "DoughnutExploded" \| " \|レーダー arマーカー" " \|レーダー ar塗りつぶし" "Surface " \| " SurfaceWireframe " \| " SurfaceTopView " \| " SurfaceTopViewWireframe " \| " バブル " \| " Bubble3DEffect " \| " StockHLC " \| " StockOHLC " \| " StockVHLC " \| "StockVOHLC " \| " CylinderColClustered " \| " CylinderColStacked " \| " CylinderColStacked100 " \| " CylinderBarClustered " \| " CylinderBarStacked " \| " CylinderBarStacked100 " \| "CylinderCol " \| \| " conecolclustered "" conecolclustered \| "" ConeColStacked100 " \| " conecolclustered \| "" conecolclustered " \| " ConeBarStacked100 " \| " conecol " \| "PyramidColClustered " \| " PyramidColStacked " \| " PyramidColStacked100 " \| " PyramidBarClustered " \| " PyramidBarStacked " \| " PyramidBarStacked100 " \| " PyramidCol " \| " 3dcolumn "\| "行" \| "3dline" \| "3dline \| " "扇形\| " "XYScatter" \| "3dline" \| "Area" \| "ドーナツ" \| "レーダー" \| "ヒストグラム" \| "boxwhisker" \| "パレートの\| "" regionmap \| "" ツリー \|マップ "" \|ウォーターフォール " \| " サンバースト "" じょうご ", sourceData: Range, 系列 by? \| :" Auto \| "" 列 "" 行 ")](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|新しいグラフを作成します。|
||[add (type: ChartType, sourceData: Range, 系列 By?: Excel. Chart系列 By)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|新しいグラフを作成します。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|グラフ名を使用してグラフを取得します。 同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|コレクション内での位置を基にグラフを取得します。|
||[count](/javascript/api/excel/excel.chartcollection#count)|ワークシート上のグラフの数を返します。 値の取得のみ可能です。|
||[items](/javascript/api/excel/excel.chartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[$all](/javascript/api/excel/excel.chartcollectionloadoptions#$all)||
||[直交](/javascript/api/excel/excel.chartcollectionloadoptions#axes)|コレクション内の各アイテムについて: グラフ軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chartcollectionloadoptions#datalabels)|コレクション内の各アイテムについて: グラフの datalabels を表します。|
||[format](/javascript/api/excel/excel.chartcollectionloadoptions#format)|コレクション内の各アイテムについて: グラフエリアの書式設定プロパティをカプセル化します。|
||[height](/javascript/api/excel/excel.chartcollectionloadoptions#height)|コレクション内の各項目について、: グラフオブジェクトの高さをポイント単位で表します。|
||[left](/javascript/api/excel/excel.chartcollectionloadoptions#left)|コレクション内の各項目について、グラフの左側からワークシートの原点までの距離をポイント単位で指定します。|
||[まつわる](/javascript/api/excel/excel.chartcollectionloadoptions#legend)|コレクション内の各アイテムについて: グラフの凡例を表します。|
||[name](/javascript/api/excel/excel.chartcollectionloadoptions#name)|コレクション内の各アイテムについて: chart オブジェクトの名前を表します。|
||[級数](/javascript/api/excel/excel.chartcollectionloadoptions#series)|コレクション内の各アイテムについて、グラフ内の1つのデータ系列または系列のコレクションを表します。|
||[title](/javascript/api/excel/excel.chartcollectionloadoptions#title)|コレクション内の各項目について: タイトルのテキスト、表示、位置、書式設定を含む、指定されたグラフのタイトルを表します。|
||[top](/javascript/api/excel/excel.chartcollectionloadoptions#top)|コレクション内の各項目について、: オブジェクトの上端から (ワークシートの) 行1の上端またはグラフエリアの上端までの距離 (ポイント単位) を表します。|
||[width](/javascript/api/excel/excel.chartcollectionloadoptions#width)|コレクション内の各項目について、: グラフオブジェクトの幅をポイント単位で表します。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[直交](/javascript/api/excel/excel.chartdata#axes)|グラフの軸を表します。 読み取り専用です。|
||[dataLabels](/javascript/api/excel/excel.chartdata#datalabels)|グラフのデータラベルを表します。 読み取り専用です。|
||[format](/javascript/api/excel/excel.chartdata#format)|グラフ領域の書式設定プロパティをカプセル化します。 読み取り専用です。|
||[height](/javascript/api/excel/excel.chartdata#height)|グラフ オブジェクトの高さをポイント単位で表します。|
||[left](/javascript/api/excel/excel.chartdata#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[まつわる](/javascript/api/excel/excel.chartdata#legend)|グラフの凡例を表します。 読み取り専用。|
||[name](/javascript/api/excel/excel.chartdata#name)|グラフ オブジェクトの名前を表します。|
||[級数](/javascript/api/excel/excel.chartdata#series)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。 読み取り専用です。|
||[title](/javascript/api/excel/excel.chartdata#title)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.chartdata#top)|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.chartdata#width)|グラフ オブジェクトの幅をポイント単位で表します。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|現在のグラフのデータ ラベルの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
||[set (properties: ChartDataLabelFormat)](/javascript/api/excel/excel.chartdatalabelformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartDataLabelFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartdatalabelformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[font](/javascript/api/excel/excel.chartdatalabelformatdata#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。 読み取り専用です。|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartdatalabelformatloadoptions#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[font](/javascript/api/excel/excel.chartdatalabelformatupdatedata#font)|グラフのデータ ラベルのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。 値の取得のみ可能です。|
||[記号](/javascript/api/excel/excel.chartdatalabels#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[set (properties: ChartDataLabels)](/javascript/api/excel/excel.chartdatalabels#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartDataLabelsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartdatalabels#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[format](/javascript/api/excel/excel.chartdatalabelsdata#format)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。 値の取得のみ可能です。|
||[position](/javascript/api/excel/excel.chartdatalabelsdata#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabelsdata#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsdata#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsdata#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsdata#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsdata#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsdata#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsdata#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelsloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartdatalabelsloadoptions#format)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。|
||[position](/javascript/api/excel/excel.chartdatalabelsloadoptions#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabelsloadoptions#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsloadoptions#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsloadoptions#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsloadoptions#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsloadoptions#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[format](/javascript/api/excel/excel.chartdatalabelsupdatedata#format)|グラフのデータ ラベルの書式 (塗りつぶしとフォントの書式設定を含む) を表します。|
||[position](/javascript/api/excel/excel.chartdatalabelsupdatedata#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabelsupdatedata#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列を設定します。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsupdatedata#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsupdatedata#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsupdatedata#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsupdatedata#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|グラフ要素の塗りつぶしの色をクリアします。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|グラフ要素の塗りつぶしの書式設定を均一な色に設定します。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfont#name)|フォント名 (例: "Calibri")|
||[set (プロパティ: Excel. ChartFont)](/javascript/api/excel/excel.chartfont#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartFontUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartfont#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[size](/javascript/api/excel/excel.chartfont#size)|フォント サイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|フォントに適用する下線の種類。 詳細については、「Excel のグラフ」を参照してください。|
|[ChartFontData](/javascript/api/excel/excel.chartfontdata)|[bold](/javascript/api/excel/excel.chartfontdata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfontdata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.chartfontdata#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfontdata#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.chartfontdata#size)|フォント サイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfontdata#underline)|フォントに適用する下線の種類。 詳細については、「Excel のグラフ」を参照してください。|
|[ChartFontLoadOptions](/javascript/api/excel/excel.chartfontloadoptions)|[$all](/javascript/api/excel/excel.chartfontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.chartfontloadoptions#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfontloadoptions#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.chartfontloadoptions#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfontloadoptions#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.chartfontloadoptions#size)|フォント サイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfontloadoptions#underline)|フォントに適用する下線の種類。 詳細については、「Excel のグラフ」を参照してください。|
|[ChartFontUpdateData](/javascript/api/excel/excel.chartfontupdatedata)|[bold](/javascript/api/excel/excel.chartfontupdatedata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.chartfontupdatedata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.chartfontupdatedata#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.chartfontupdatedata#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.chartfontupdatedata#size)|フォント サイズ (例: 11)|
||[underline](/javascript/api/excel/excel.chartfontupdatedata#underline)|フォントに適用する下線の種類。 詳細については、「Excel のグラフ」を参照してください。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|グラフの目盛線の書式設定を表します。 値の取得のみ可能です。|
||[set (properties: エクセル目盛線)](/javascript/api/excel/excel.chartgridlines#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartGridlinesUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartgridlines#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|軸の目盛線を表示するか非表示にするかを表すブール型の値。|
|[ChartGridlinesData](/javascript/api/excel/excel.chartgridlinesdata)|[format](/javascript/api/excel/excel.chartgridlinesdata#format)|グラフの目盛線の書式設定を表します。 値の取得のみ可能です。|
||[visible](/javascript/api/excel/excel.chartgridlinesdata#visible)|軸の目盛線を表示するか非表示にするかを表すブール型の値。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|グラフの線の書式設定を表します。 読み取り専用です。|
||[set (properties: ChartGridlinesFormat)](/javascript/api/excel/excel.chartgridlinesformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartGridlinesFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartgridlinesformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartGridlinesFormatData](/javascript/api/excel/excel.chartgridlinesformatdata)|[line](/javascript/api/excel/excel.chartgridlinesformatdata#line)|グラフの線の書式設定を表します。 読み取り専用です。|
|[ChartGridlinesFormatLoadOptions](/javascript/api/excel/excel.chartgridlinesformatloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartgridlinesformatloadoptions#line)|グラフの線の書式設定を表します。|
|[ChartGridlinesFormatUpdateData](/javascript/api/excel/excel.chartgridlinesformatupdatedata)|[line](/javascript/api/excel/excel.chartgridlinesformatupdatedata#line)|グラフの線の書式設定を表します。|
|[ChartGridlinesLoadOptions](/javascript/api/excel/excel.chartgridlinesloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartgridlinesloadoptions#format)|グラフの目盛線の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartgridlinesloadoptions#visible)|軸の目盛線を表示するか非表示にするかを表すブール型の値。|
|[ChartGridlinesUpdateData](/javascript/api/excel/excel.chartgridlinesupdatedata)|[format](/javascript/api/excel/excel.chartgridlinesupdatedata#format)|グラフの目盛線の書式設定を表します。|
||[visible](/javascript/api/excel/excel.chartgridlinesupdatedata#visible)|軸の目盛線を表示するか非表示にするかを表すブール型の値。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|
||[position](/javascript/api/excel/excel.chartlegend#position)|グラフの凡例の位置を表します。 詳細については、「ChartLegendPosition」を参照してください。|
||[format](/javascript/api/excel/excel.chartlegend#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。 読み取り専用です。|
||[set (プロパティ: Excel. ChartLegend)](/javascript/api/excel/excel.chartlegend#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartLegendUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartlegend#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[format](/javascript/api/excel/excel.chartlegenddata#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。 読み取り専用です。|
||[overlay](/javascript/api/excel/excel.chartlegenddata#overlay)|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|
||[position](/javascript/api/excel/excel.chartlegenddata#position)|グラフの凡例の位置を表します。 詳細については、「ChartLegendPosition」を参照してください。|
||[visible](/javascript/api/excel/excel.chartlegenddata#visible)|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|グラフの凡例のフォント属性 (フォント名、フォント サイズ、色など) を表します。 値の取得のみ可能です。|
||[set (properties: ChartLegendFormat)](/javascript/api/excel/excel.chartlegendformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartLegendFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartlegendformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[font](/javascript/api/excel/excel.chartlegendformatdata#font)|グラフの凡例のフォント属性 (フォント名、フォント サイズ、色など) を表します。 値の取得のみ可能です。|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[$all](/javascript/api/excel/excel.chartlegendformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartlegendformatloadoptions#font)|グラフの凡例のフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[font](/javascript/api/excel/excel.chartlegendformatupdatedata#font)|グラフの凡例のフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[$all](/javascript/api/excel/excel.chartlegendloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartlegendloadoptions#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。|
||[overlay](/javascript/api/excel/excel.chartlegendloadoptions#overlay)|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|
||[position](/javascript/api/excel/excel.chartlegendloadoptions#position)|グラフの凡例の位置を表します。 詳細については、「ChartLegendPosition」を参照してください。|
||[visible](/javascript/api/excel/excel.chartlegendloadoptions#visible)|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[format](/javascript/api/excel/excel.chartlegendupdatedata#format)|塗りつぶしとフォントの書式設定を含む、グラフの凡例の書式設定を表します。|
||[overlay](/javascript/api/excel/excel.chartlegendupdatedata#overlay)|グラフの凡例をグラフの本体に重ねるかどうかを指定するブール型の値です。|
||[position](/javascript/api/excel/excel.chartlegendupdatedata#position)|グラフの凡例の位置を表します。 詳細については、「ChartLegendPosition」を参照してください。|
||[visible](/javascript/api/excel/excel.chartlegendupdatedata#visible)|ChartLegend オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|グラフ要素の線の書式をクリアします。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|グラフの線の色を表す HTML カラー コード。|
||[set (properties: ChartLineFormat)](/javascript/api/excel/excel.chartlineformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartLineFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartlineformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[color](/javascript/api/excel/excel.chartlineformatdata#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[$all](/javascript/api/excel/excel.chartlineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartlineformatloadoptions#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[color](/javascript/api/excel/excel.chartlineformatupdatedata#color)|グラフの線の色を表す HTML カラー コード。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[$all](/javascript/api/excel/excel.chartloadoptions#$all)||
||[直交](/javascript/api/excel/excel.chartloadoptions#axes)|グラフの軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chartloadoptions#datalabels)|グラフのデータラベルを表します。|
||[format](/javascript/api/excel/excel.chartloadoptions#format)|グラフ領域の書式設定プロパティをカプセル化します。|
||[height](/javascript/api/excel/excel.chartloadoptions#height)|グラフ オブジェクトの高さをポイント単位で表します。|
||[left](/javascript/api/excel/excel.chartloadoptions#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[まつわる](/javascript/api/excel/excel.chartloadoptions#legend)|グラフの凡例を表します。|
||[name](/javascript/api/excel/excel.chartloadoptions#name)|グラフ オブジェクトの名前を表します。|
||[級数](/javascript/api/excel/excel.chartloadoptions#series)|グラフの 1 つのデータ系列またはデータ系列のコレクションを表します。|
||[title](/javascript/api/excel/excel.chartloadoptions#title)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。|
||[top](/javascript/api/excel/excel.chartloadoptions#top)|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.chartloadoptions#width)|グラフ オブジェクトの幅をポイント単位で表します。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|グラフのポイントの書式設定プロパティをカプセル化します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.chartpoint#value)|グラフのポイントの値を返します。 読み取り専用です。|
||[set (プロパティ: Excel. ChartPoint)](/javascript/api/excel/excel.chartpoint#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartPointUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartpoint#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[format](/javascript/api/excel/excel.chartpointdata#format)|グラフのポイントの書式設定プロパティをカプセル化します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.chartpointdata#value)|グラフのポイントの値を返します。 読み取り専用です。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|背景の書式設定情報を含むグラフの塗りつぶしの書式を表します。 読み取り専用です。|
||[set (プロパティ: Excel. ChartPointFormat)](/javascript/api/excel/excel.chartpointformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartPointFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartpointformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[$all](/javascript/api/excel/excel.chartpointformatloadoptions#$all)||
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[$all](/javascript/api/excel/excel.chartpointloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointloadoptions#format)|グラフのポイントの書式設定プロパティをカプセル化します。|
||[value](/javascript/api/excel/excel.chartpointloadoptions#value)|グラフのポイントの値を返します。 読み取り専用です。|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[format](/javascript/api/excel/excel.chartpointupdatedata#format)|グラフのポイントの書式設定プロパティをカプセル化します。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|データ系列内の位置に基づくポイントを取得します。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|系列内にあるグラフのポイントの数を取得します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[$all](/javascript/api/excel/excel.chartpointscollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointscollectionloadoptions#format)|コレクション内の各アイテムについて: 書式プロパティのグラフポイントをカプセル化します。|
||[value](/javascript/api/excel/excel.chartpointscollectionloadoptions#value)|コレクション内の各項目について: グラフ要素の値を返します。 読み取り専用です。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|グラフのデータ系列の名前を表します。|
||[format](/javascript/api/excel/excel.chartseries#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。 読み取り専用です。|
||[数](/javascript/api/excel/excel.chartseries#points)|データ系列にあるすべてのポイントのコレクションを返します。 読み取り専用です。|
||[set (プロパティ: Excel. ChartSeries)](/javascript/api/excel/excel.chartseries#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartSeriesUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartseries#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|コレクション内の位置に基づいてデータ系列を取得します。|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|コレクション内にあるデータ系列の数を取得します。 値の取得のみ可能です。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Chartcharts Collectionloadoptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[$all](/javascript/api/excel/excel.chartseriescollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriescollectionloadoptions#format)|コレクション内の各項目について: 塗りつぶしと線の書式設定を含む、グラフ系列の書式設定を表します。|
||[name](/javascript/api/excel/excel.chartseriescollectionloadoptions#name)|コレクション内の各アイテムについて: グラフの系列の名前を表します。|
||[数](/javascript/api/excel/excel.chartseriescollectionloadoptions#points)|コレクション内の各アイテムについて: 系列内のすべてのポイントのコレクションを表します。|
|[Chart系列データ](/javascript/api/excel/excel.chartseriesdata)|[format](/javascript/api/excel/excel.chartseriesdata#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。 読み取り専用です。|
||[name](/javascript/api/excel/excel.chartseriesdata#name)|グラフのデータ系列の名前を表します。|
||[数](/javascript/api/excel/excel.chartseriesdata#points)|データ系列にあるすべてのポイントのコレクションを返します。 読み取り専用です。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|背景の書式設定情報を含むグラフ系列の塗りつぶしの書式を表します。 読み取り専用です。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|線の書式設定を表します。 読み取り専用です。|
||[set (プロパティ: Excel. Chartの形式)](/javascript/api/excel/excel.chartseriesformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartSeriesFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartseriesformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[Chart系列 Formatdata](/javascript/api/excel/excel.chartseriesformatdata)|[line](/javascript/api/excel/excel.chartseriesformatdata#line)|線の書式設定を表します。 読み取り専用です。|
|[Chart系列 Formatloadoptions](/javascript/api/excel/excel.chartseriesformatloadoptions)|[$all](/javascript/api/excel/excel.chartseriesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartseriesformatloadoptions#line)|線の書式設定を表します。|
|[ChartSeriesFormatUpdateData](/javascript/api/excel/excel.chartseriesformatupdatedata)|[line](/javascript/api/excel/excel.chartseriesformatupdatedata#line)|線の書式設定を表します。|
|[Chart系列 Loadoptions](/javascript/api/excel/excel.chartseriesloadoptions)|[$all](/javascript/api/excel/excel.chartseriesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriesloadoptions#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。|
||[name](/javascript/api/excel/excel.chartseriesloadoptions#name)|グラフのデータ系列の名前を表します。|
||[数](/javascript/api/excel/excel.chartseriesloadoptions#points)|データ系列にあるすべてのポイントのコレクションを返します。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[format](/javascript/api/excel/excel.chartseriesupdatedata#format)|グラフの系列の書式設定を表します。これには、塗りつぶしと線の書式設定が含まれます。|
||[name](/javascript/api/excel/excel.chartseriesupdatedata#name)|グラフのデータ系列の名前を表します。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
||[format](/javascript/api/excel/excel.charttitle#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。 読み取り専用です。|
||[set (properties: Excel. ChartTitle)](/javascript/api/excel/excel.charttitle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTitleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttitle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[text](/javascript/api/excel/excel.charttitle#text)|グラフのタイトルのテキストを表します。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|
|[Chartタイトルデータ](/javascript/api/excel/excel.charttitledata)|[format](/javascript/api/excel/excel.charttitledata#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。 読み取り専用です。|
||[overlay](/javascript/api/excel/excel.charttitledata#overlay)|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
||[text](/javascript/api/excel/excel.charttitledata#text)|グラフのタイトルのテキストを表します。|
||[visible](/javascript/api/excel/excel.charttitledata#visible)|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 値の取得のみ可能です。|
||[set (プロパティ: Excel. Charttitle 形式)](/javascript/api/excel/excel.charttitleformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTitleFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttitleformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[Chartタイトル Formatdata](/javascript/api/excel/excel.charttitleformatdata)|[font](/javascript/api/excel/excel.charttitleformatdata#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。 値の取得のみ可能です。|
|[Chartタイトル Formatloadoptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[$all](/javascript/api/excel/excel.charttitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.charttitleformatloadoptions#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[font](/javascript/api/excel/excel.charttitleformatupdatedata#font)|オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[Chartタイトル Loadoptions](/javascript/api/excel/excel.charttitleloadoptions)|[$all](/javascript/api/excel/excel.charttitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttitleloadoptions#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。|
||[overlay](/javascript/api/excel/excel.charttitleloadoptions#overlay)|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
||[text](/javascript/api/excel/excel.charttitleloadoptions#text)|グラフのタイトルのテキストを表します。|
||[visible](/javascript/api/excel/excel.charttitleloadoptions#visible)|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[format](/javascript/api/excel/excel.charttitleupdatedata#format)|塗りつぶしとフォントの書式設定を含む、グラフタイトルの書式設定を表します。|
||[overlay](/javascript/api/excel/excel.charttitleupdatedata#overlay)|グラフのタイトルをグラフに重ねるかどうかを表すブール型の値。|
||[text](/javascript/api/excel/excel.charttitleupdatedata#text)|グラフのタイトルのテキストを表します。|
||[visible](/javascript/api/excel/excel.charttitleupdatedata#visible)|ChartTitle オブジェクトを表示または非表示にするかを表すブール型の値。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[直交](/javascript/api/excel/excel.chartupdatedata#axes)|グラフの軸を表します。|
||[dataLabels](/javascript/api/excel/excel.chartupdatedata#datalabels)|グラフのデータラベルを表します。|
||[format](/javascript/api/excel/excel.chartupdatedata#format)|グラフ領域の書式設定プロパティをカプセル化します。|
||[height](/javascript/api/excel/excel.chartupdatedata#height)|グラフ オブジェクトの高さをポイント単位で表します。|
||[left](/javascript/api/excel/excel.chartupdatedata#left)|グラフの左側からワークシートの原点までの距離 (ポイント単位)。|
||[まつわる](/javascript/api/excel/excel.chartupdatedata#legend)|グラフの凡例を表します。|
||[name](/javascript/api/excel/excel.chartupdatedata#name)|グラフ オブジェクトの名前を表します。|
||[title](/javascript/api/excel/excel.chartupdatedata#title)|指定したグラフのタイトル (タイトルのテキスト、表示/非表示、位置、書式設定など) を表します。|
||[top](/javascript/api/excel/excel.chartupdatedata#top)|オブジェクトの上端から (ワークシートの) 1 行目の上部または (グラフの) グラフ領域の上部までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.chartupdatedata#width)|グラフ オブジェクトの幅をポイント単位で表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|名前に関連付けられている範囲オブジェクトを返します。 名前付きアイテムの型が範囲でない場合、エラーをスローします。|
||[name](/javascript/api/excel/excel.nameditem#name)|オブジェクトの名前。 値の取得のみ可能です。|
||[type](/javascript/api/excel/excel.nameditem#type)|名前の数式によって返される値の型を示します。 詳細については、「Excel. NamedItemType」を参照してください。 読み取り専用です。|
||[value](/javascript/api/excel/excel.nameditem#value)|名前の数式で計算された値を表します。 名前付き範囲の場合は範囲のアドレスを返します。 読み取り専用です。|
||[set (properties: Excel. NamedItem)](/javascript/api/excel/excel.nameditem#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: NamedItemUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.nameditem#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|オブジェクトを表示するかどうかを指定します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|名前を使用して、NamedItem オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[$all](/javascript/api/excel/excel.nameditemcollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemcollectionloadoptions#name)|コレクション内の各アイテムについて: オブジェクトの名前。 読み取り専用です。|
||[type](/javascript/api/excel/excel.nameditemcollectionloadoptions#type)|コレクション内の各アイテムについて: 名前の数式によって返される値の型を示します。 詳細については、「Excel. NamedItemType」を参照してください。 読み取り専用です。|
||[value](/javascript/api/excel/excel.nameditemcollectionloadoptions#value)|コレクション内の各アイテムについて: 名前の数式で計算された値を表します。 名前付き範囲の場合は範囲のアドレスを返します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.nameditemcollectionloadoptions#visible)|コレクション内の各アイテムについて: オブジェクトを表示するかどうかを指定します。|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[name](/javascript/api/excel/excel.nameditemdata#name)|オブジェクトの名前。 値の取得のみ可能です。|
||[type](/javascript/api/excel/excel.nameditemdata#type)|名前の数式によって返される値の型を示します。 詳細については、「Excel. NamedItemType」を参照してください。 読み取り専用です。|
||[value](/javascript/api/excel/excel.nameditemdata#value)|名前の数式で計算された値を表します。 名前付き範囲の場合は範囲のアドレスを返します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.nameditemdata#visible)|オブジェクトを表示するかどうかを指定します。|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[$all](/javascript/api/excel/excel.nameditemloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemloadoptions#name)|オブジェクトの名前。 値の取得のみ可能です。|
||[type](/javascript/api/excel/excel.nameditemloadoptions#type)|名前の数式によって返される値の型を示します。 詳細については、「Excel. NamedItemType」を参照してください。 読み取り専用です。|
||[value](/javascript/api/excel/excel.nameditemloadoptions#value)|名前の数式で計算された値を表します。 名前付き範囲の場合は範囲のアドレスを返します。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.nameditemloadoptions#visible)|オブジェクトを表示するかどうかを指定します。|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[visible](/javascript/api/excel/excel.nameditemupdatedata#visible)|オブジェクトを表示するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.range#clear-applyto-)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|範囲の値、書式、塗りつぶし、罫線などをクリアします。|
||[delete (shift: "Up" \| "Left")](/javascript/api/excel/excel.range#delete-shift-)|範囲に関連付けられているセルを削除します。|
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
||[insert (shift: "Down" \| "Right")](/javascript/api/excel/excel.range#insert-shift-)|この範囲を占めるセルまたはセルの範囲をワークシートに挿入し、領域を空けるために他のセルをシフトします。この時点で空き領域に位置する、新しい Range オブジェクトが返されます。|
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
||[set (properties: Excel. Range)](/javascript/api/excel/excel.range#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.range#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[track()](/javascript/api/excel/excel.range#track--)|ドキュメントの環境変更に基づいて自動的に調整する目的でオブジェクトを追跡します。 これは context.trackedObjects.add(thisObject) 呼び出しの省略形です。 ".sync" 呼び出し間で、かつ ".run" バッチの連続実行の外でこのオブジェクトを使用しているとき、オブジェクトであるプロパティを設定したか、あるメソッドを呼び出したときに "InvalidObjectPath" エラーが表示される場合、オブジェクトを最初に作成したときに、追跡対象オブジェクトの集まりにそのオブジェクトを追加しておく必要がありました。|
||[untrack()](/javascript/api/excel/excel.range#untrack--)|前に追跡されていた場合、このオブジェクトに関連付けられているメモリを解放します。 これは context.trackedObjects.remove(thisObject) 呼び出しの省略形です。 追跡対象オブジェクトが多いとホスト アプリケーションの動作が遅くなります。追加したオブジェクトが不要になったら、必ずそれを解放してください。 メモリ リリースを有効にするには、"context.sync()" を先に呼び出す必要があります。|
||[values](/javascript/api/excel/excel.range#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel BorderIndex」を参照してください。 読み取り専用です。|
||[set (プロパティ: Excel. RangeBorder)](/javascript/api/excel/excel.rangeborder#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeBorderUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangeborder#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[style](/javascript/api/excel/excel.rangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|範囲周辺の罫線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "edgeright" \| "insidevertical \| " "insidevertical" \| "DiagonalDown" \| "DiagonalUp")](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。 読み取り専用です。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.rangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangebordercollectionloadoptions#color)|コレクション内の各項目について、フォーム #RRGGBB の境界線の色を表す HTML カラーコード ("FFA500" など) または名前付きの HTML 色 (例: "オレンジ")。|
||[sideIndex](/javascript/api/excel/excel.rangebordercollectionloadoptions#sideindex)|コレクション内の各項目について、境界線の特定の辺を示す定数値。 詳細については、「Excel BorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangebordercollectionloadoptions#style)|コレクション内の各アイテムについて: 境界線のスタイルを指定する線スタイルの定数のいずれかです。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangebordercollectionloadoptions#weight)|コレクション内の各アイテムについて: 範囲を囲む境界線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[color](/javascript/api/excel/excel.rangeborderdata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.rangeborderdata#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel BorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangeborderdata#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangeborderdata#weight)|範囲周辺の罫線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[$all](/javascript/api/excel/excel.rangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangeborderloadoptions#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.rangeborderloadoptions#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel BorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangeborderloadoptions#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangeborderloadoptions#weight)|範囲周辺の罫線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[color](/javascript/api/excel/excel.rangeborderupdatedata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[style](/javascript/api/excel/excel.rangeborderupdatedata#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.rangeborderupdatedata#weight)|範囲周辺の罫線の太さを指定します。 詳細については、「Excel BorderWeight」を参照してください。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[address](/javascript/api/excel/excel.rangedata#address)|A1 スタイルの範囲参照を表します。 Address 値にはシート参照が含まれます (例: "Sheet1!A1: B4 ") 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangedata#addresslocal)|ユーザーの言語で指定された範囲の範囲参照を表します。 読み取り専用です。|
||[cellCount](/javascript/api/excel/excel.rangedata#cellcount)|範囲に含まれるセルの数。 セルの数が 2^31-1 (2,147,483,647) を超えると、この API は -1 を返します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangedata#columncount)|範囲に含まれる列の合計数を表します。 読み取り専用です。|
||[columnIndex](/javascript/api/excel/excel.rangedata#columnindex)|範囲に含まれる最初のセルの列番号を表します。 0 を起点とする番号になります。 読み取り専用。|
||[format](/javascript/api/excel/excel.rangedata#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。 読み取り専用です。|
||[formulas](/javascript/api/excel/excel.rangedata#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangedata#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[numberFormat](/javascript/api/excel/excel.rangedata#numberformat)|指定された範囲の Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangedata#rowcount)|範囲に含まれる行の合計数を返します。 読み取り専用です。|
||[rowIndex](/javascript/api/excel/excel.rangedata#rowindex)|範囲に含まれる最初のセルの行番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangedata#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangedata#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangedata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|範囲の背景をリセットします。|
||[color](/javascript/api/excel/excel.rangefill#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ")|
||[set (properties: エクセル Fill)](/javascript/api/excel/excel.rangefill#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeFillUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangefill#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[color](/javascript/api/excel/excel.rangefilldata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ")|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[$all](/javascript/api/excel/excel.rangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangefillloadoptions#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ")|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[color](/javascript/api/excel/excel.rangefillupdatedata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.rangefont#name)|フォント名 (例: "Calibri")|
||[set (properties: エクセルフォント)](/javascript/api/excel/excel.rangefont#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeFontUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangefont#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[size](/javascript/api/excel/excel.rangefont#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefont#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Range過小な Linestyle」を参照してください。|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[bold](/javascript/api/excel/excel.rangefontdata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefontdata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.rangefontdata#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.rangefontdata#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.rangefontdata#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefontdata#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Range過小な Linestyle」を参照してください。|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[$all](/javascript/api/excel/excel.rangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.rangefontloadoptions#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefontloadoptions#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.rangefontloadoptions#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.rangefontloadoptions#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.rangefontloadoptions#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefontloadoptions#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Range過小な Linestyle」を参照してください。|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[bold](/javascript/api/excel/excel.rangefontupdatedata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.rangefontupdatedata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.rangefontupdatedata#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/excel/excel.rangefontupdatedata#name)|フォント名 (例: "Calibri")|
||[size](/javascript/api/excel/excel.rangefontupdatedata#size)|フォント サイズ。|
||[underline](/javascript/api/excel/excel.rangefontupdatedata#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Range過小な Linestyle」を参照してください。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.rangeformat#font)|範囲全体に定義されたフォント オブジェクトを返します。 読み取り専用です。|
||[set (properties: エクセル形式)](/javascript/api/excel/excel.rangeformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangeformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|オブジェクト内のテキストを Excel でラップするかどうかを表します。 null 値は、範囲全体に一様なラップ設定がないことを表します。|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[borders](/javascript/api/excel/excel.rangeformatdata#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.rangeformatdata#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.rangeformatdata#font)|範囲全体に定義されたフォント オブジェクトを返します。 読み取り専用です。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatdata#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatdata#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.rangeformatdata#wraptext)|オブジェクト内のテキストを Excel でラップするかどうかを表します。 null 値は、範囲全体に一様なラップ設定がないことを表します。|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[$all](/javascript/api/excel/excel.rangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.rangeformatloadoptions#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.rangeformatloadoptions#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。|
||[font](/javascript/api/excel/excel.rangeformatloadoptions#font)|範囲全体に定義されたフォント オブジェクトを返します。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.rangeformatloadoptions#wraptext)|オブジェクト内のテキストを Excel でラップするかどうかを表します。 null 値は、範囲全体に一様なラップ設定がないことを表します。|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[borders](/javascript/api/excel/excel.rangeformatupdatedata#borders)|選択した範囲全体に適用する境界線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.rangeformatupdatedata#fill)|範囲全体に定義された塗りつぶしオブジェクトを返します。|
||[font](/javascript/api/excel/excel.rangeformatupdatedata#font)|範囲全体に定義されたフォント オブジェクトを返します。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#horizontalalignment)|指定したオブジェクトの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#verticalalignment)|指定したオブジェクトの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.rangeformatupdatedata#wraptext)|オブジェクト内のテキストを Excel でラップするかどうかを表します。 null 値は、範囲全体に一様なラップ設定がないことを表します。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[$all](/javascript/api/excel/excel.rangeloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeloadoptions#address)|A1 スタイルの範囲参照を表します。 Address 値にはシート参照が含まれます (例: "Sheet1!A1: B4 ") 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeloadoptions#addresslocal)|ユーザーの言語で指定された範囲の範囲参照を表します。 読み取り専用です。|
||[cellCount](/javascript/api/excel/excel.rangeloadoptions#cellcount)|範囲に含まれるセルの数。 セルの数が 2^31-1 (2,147,483,647) を超えると、この API は -1 を返します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangeloadoptions#columncount)|範囲に含まれる列の合計数を表します。 読み取り専用です。|
||[columnIndex](/javascript/api/excel/excel.rangeloadoptions#columnindex)|範囲に含まれる最初のセルの列番号を表します。 0 を起点とする番号になります。 読み取り専用。|
||[format](/javascript/api/excel/excel.rangeloadoptions#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。|
||[formulas](/javascript/api/excel/excel.rangeloadoptions#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeloadoptions#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[numberFormat](/javascript/api/excel/excel.rangeloadoptions#numberformat)|指定された範囲の Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangeloadoptions#rowcount)|範囲に含まれる行の合計数を返します。 読み取り専用です。|
||[rowIndex](/javascript/api/excel/excel.rangeloadoptions#rowindex)|範囲に含まれる最初のセルの行番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangeloadoptions#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangeloadoptions#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangeloadoptions#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[worksheet](/javascript/api/excel/excel.rangeloadoptions#worksheet)|現在の範囲を含んでいるワークシート。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[format](/javascript/api/excel/excel.rangeupdatedata#format)|Format オブジェクト (範囲のフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化するオブジェクト) を返します。|
||[formulas](/javascript/api/excel/excel.rangeupdatedata#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeupdatedata#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[numberFormat](/javascript/api/excel/excel.rangeupdatedata#numberformat)|指定された範囲の Excel の数値書式コードを表します。|
||[values](/javascript/api/excel/excel.rangeupdatedata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|テーブルを削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.table#getdatabodyrange--)|テーブルのデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|テーブルのヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|テーブル全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.table#gettotalrowrange--)|テーブルの集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.table#name)|テーブルの名前。|
||[列](/javascript/api/excel/excel.table#columns)|テーブルに含まれるすべての列のコレクションを表します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.table#id)|指定されたブックのテーブルを一意に識別する値を返します。識別子の値は、テーブルの名前が変更された場合も変わりません。読み取り専用です。|
||[rows](/javascript/api/excel/excel.table#rows)|テーブルに含まれるすべての行のコレクションを表します。 読み取り専用です。|
||[set (properties: Excel. Table)](/javascript/api/excel/excel.table#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TableUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.table#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|ヘッダー行を表示するかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|集計行を表示するかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.table#style)|テーブル スタイルを表す定数値。使用可能な値は次のとおりです。TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address: Range \| String, hasheaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|新しいテーブルを作成します。範囲オブジェクトまたはソース アドレスにより、テーブルが追加されるワークシートが判断されます。テーブルが追加できない場合 (たとえば、アドレスが無効な場合や、テーブルが別のテーブルと重複している場合) は、エラーがスローされます。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|コレクション内の位置に基づいてテーブルを取得します。|
||[count](/javascript/api/excel/excel.tablecollection#count)|ブックに含まれるテーブルの数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecollectionloadoptions#$all)||
||[列](/javascript/api/excel/excel.tablecollectionloadoptions#columns)|コレクション内の各アイテムについて: テーブル内のすべての列のコレクションを表します。|
||[id](/javascript/api/excel/excel.tablecollectionloadoptions#id)|コレクション内の各アイテムについて: 指定されたブック内のテーブルを一意に識別する値を返します。 識別子の値は、テーブルの名前が変更された場合も変わりません。 読み取り専用です。|
||[name](/javascript/api/excel/excel.tablecollectionloadoptions#name)|コレクション内の各項目について: テーブルの名前。|
||[rows](/javascript/api/excel/excel.tablecollectionloadoptions#rows)|コレクション内の各アイテムについて: テーブル内のすべての行のコレクションを表します。|
||[showHeaders](/javascript/api/excel/excel.tablecollectionloadoptions#showheaders)|コレクション内の各アイテムについて: ヘッダー行が表示されるかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.tablecollectionloadoptions#showtotals)|コレクション内の各アイテムについて: [合計] 行が表示されるかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.tablecollectionloadoptions#style)|コレクション内の各項目について、表のスタイルを表す定数値。 使用可能な値は次のとおりです。 TableStyleLight1 スルー TableStyleLight21、TableStyleMedium1 スルー TableStyleMedium28、TableStyleStyleDark1 スルー TableStyleStyleDark11。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|テーブルから列を削除します。|
||[Getの Odyrange ()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|列のデータ本体に関連付けられた範囲オブジェクトを取得します。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|列のヘッダー行に関連付けられた範囲オブジェクトを取得します。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|列全体に関連付けられた範囲オブジェクトを取得します。|
||[getTotalRowRange)](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|列の集計行に関連付けられた範囲オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|テーブル列の名前を表します。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|テーブル内の列を識別する一意のキーを返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|テーブルの列コレクション内の列のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用。|
||[set (properties: TableColumn)](/javascript/api/excel/excel.tablecolumn#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TableColumnUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.tablecolumn#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index?: number, values?: Array<Array<boolean \| \|文字列番号>> \| boolean \|文字列\| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|テーブルに新しい列を追加します。|
||[getItem (key: number \|文字列)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|名前または ID を使用して列オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|コレクション内の位置に基づいて列を取得します。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|テーブルの列数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableColumnCollectionLoadOptions](/javascript/api/excel/excel.tablecolumncollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecolumncollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumncollectionloadoptions#id)|コレクション内の各アイテムについて: テーブル内の列を識別する一意のキーを返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.tablecolumncollectionloadoptions#index)|コレクション内の各アイテムについて: テーブルの columns コレクション内の列のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用。|
||[name](/javascript/api/excel/excel.tablecolumncollectionloadoptions#name)|コレクション内の各アイテムについて: テーブルの列の名前を表します。|
||[values](/javascript/api/excel/excel.tablecolumncollectionloadoptions#values)|コレクション内の各項目について: 指定された範囲の生の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableColumnData](/javascript/api/excel/excel.tablecolumndata)|[id](/javascript/api/excel/excel.tablecolumndata#id)|テーブル内の列を識別する一意のキーを返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.tablecolumndata#index)|テーブルの列コレクション内の列のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用。|
||[name](/javascript/api/excel/excel.tablecolumndata#name)|テーブル列の名前を表します。|
||[values](/javascript/api/excel/excel.tablecolumndata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableColumnLoadOptions](/javascript/api/excel/excel.tablecolumnloadoptions)|[$all](/javascript/api/excel/excel.tablecolumnloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumnloadoptions#id)|テーブル内の列を識別する一意のキーを返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.tablecolumnloadoptions#index)|テーブルの列コレクション内の列のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[name](/javascript/api/excel/excel.tablecolumnloadoptions#name)|テーブル列の名前を表します。|
||[values](/javascript/api/excel/excel.tablecolumnloadoptions#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableColumnUpdateData](/javascript/api/excel/excel.tablecolumnupdatedata)|[name](/javascript/api/excel/excel.tablecolumnupdatedata#name)|テーブル列の名前を表します。|
||[values](/javascript/api/excel/excel.tablecolumnupdatedata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableData](/javascript/api/excel/excel.tabledata)|[列](/javascript/api/excel/excel.tabledata#columns)|テーブルに含まれるすべての列のコレクションを表します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.tabledata#id)|指定されたブックのテーブルを一意に識別する値を返します。識別子の値は、テーブルの名前が変更された場合も変わりません。読み取り専用です。|
||[name](/javascript/api/excel/excel.tabledata#name)|テーブルの名前。|
||[rows](/javascript/api/excel/excel.tabledata#rows)|テーブルに含まれるすべての行のコレクションを表します。 読み取り専用です。|
||[showHeaders](/javascript/api/excel/excel.tabledata#showheaders)|ヘッダー行を表示するかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.tabledata#showtotals)|集計行を表示するかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.tabledata#style)|テーブル スタイルを表す定数値。使用可能な値は次のとおりです。TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[$all](/javascript/api/excel/excel.tableloadoptions#$all)||
||[列](/javascript/api/excel/excel.tableloadoptions#columns)|テーブルに含まれるすべての列のコレクションを表します。|
||[id](/javascript/api/excel/excel.tableloadoptions#id)|指定されたブックのテーブルを一意に識別する値を返します。識別子の値は、テーブルの名前が変更された場合も変わりません。読み取り専用です。|
||[name](/javascript/api/excel/excel.tableloadoptions#name)|テーブルの名前。|
||[rows](/javascript/api/excel/excel.tableloadoptions#rows)|テーブルに含まれるすべての行のコレクションを表します。|
||[showHeaders](/javascript/api/excel/excel.tableloadoptions#showheaders)|ヘッダー行を表示するかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.tableloadoptions#showtotals)|集計行を表示するかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.tableloadoptions#style)|テーブル スタイルを表す定数値。使用可能な値は次のとおりです。TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|テーブルから行を削除します。|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|行全体に関連付けられた範囲オブジェクトを返します。|
||[index](/javascript/api/excel/excel.tablerow#index)|テーブルの行コレクション内の行のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[set (プロパティ: Excel. TableRow)](/javascript/api/excel/excel.tablerow#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TableRowUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.tablerow#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[values](/javascript/api/excel/excel.tablerow#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index?: number, values?: Array<Array<boolean \|文字列\|番号>> \| boolean \|文字列\|番号)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|テーブルに 1 つ以上の行を追加します。 戻りオブジェクトは新しく追加された行の先頭になります。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|コレクション内の位置を基に行を取得します。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|テーブルの行数を返します。 読み取り専用です。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRowCollectionLoadOptions](/javascript/api/excel/excel.tablerowcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablerowcollectionloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowcollectionloadoptions#index)|コレクション内の各アイテムについて: テーブルの rows コレクション内の行のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[values](/javascript/api/excel/excel.tablerowcollectionloadoptions#values)|コレクション内の各項目について: 指定された範囲の生の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableRowData](/javascript/api/excel/excel.tablerowdata)|[index](/javascript/api/excel/excel.tablerowdata#index)|テーブルの行コレクション内の行のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[values](/javascript/api/excel/excel.tablerowdata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableRowLoadOptions](/javascript/api/excel/excel.tablerowloadoptions)|[$all](/javascript/api/excel/excel.tablerowloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowloadoptions#index)|テーブルの行コレクション内の行のインデックス番号を返します。 0 を起点とする番号になります。 読み取り専用です。|
||[values](/javascript/api/excel/excel.tablerowloadoptions#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableRowUpdateData](/javascript/api/excel/excel.tablerowupdatedata)|[values](/javascript/api/excel/excel.tablerowupdatedata#values)|指定した範囲の Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[name](/javascript/api/excel/excel.tableupdatedata#name)|テーブルの名前。|
||[showHeaders](/javascript/api/excel/excel.tableupdatedata#showheaders)|ヘッダー行を表示するかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.tableupdatedata#showtotals)|集計行を表示するかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[style](/javascript/api/excel/excel.tableupdatedata#style)|テーブル スタイルを表す定数値。使用可能な値は次のとおりです。TableStyleLight1 から TableStyleLight21、TableStyleMedium1 から TableStyleMedium28、TableStyleStyleDark1 から TableStyleStyleDark11。ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|ブックから現在選択されている1つのセル範囲を取得します。 複数の範囲が選択されている場合、このメソッドはエラーをスローします。|
||[application](/javascript/api/excel/excel.workbook#application)|このブックを含む Excel アプリケーションインスタンスを表します。 読み取り専用です。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|ブックの一部であるバインドのコレクションを表します。 読み取り専用。|
||[names](/javascript/api/excel/excel.workbook#names)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。 読み取り専用です。|
||[テーブル](/javascript/api/excel/excel.workbook#tables)|ブックに関連付けられているテーブルのコレクションを表します。 読み取り専用です。|
||[what-if](/javascript/api/excel/excel.workbook#worksheets)|ブックに関連付けられているワークシートのコレクションを表します。 読み取り専用です。|
||[set (プロパティ: Excel. Workbook)](/javascript/api/excel/excel.workbook#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: WorkbookUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.workbook#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[bindings](/javascript/api/excel/excel.workbookdata#bindings)|ブックの一部であるバインドのコレクションを表します。 読み取り専用。|
||[names](/javascript/api/excel/excel.workbookdata#names)|ブック スコープの名前付き項目 (名前付き範囲と名前付き定数) のコレクションを表します。 読み取り専用です。|
||[テーブル](/javascript/api/excel/excel.workbookdata#tables)|ブックに関連付けられているテーブルのコレクションを表します。 読み取り専用です。|
||[what-if](/javascript/api/excel/excel.workbookdata#worksheets)|ブックに関連付けられているワークシートのコレクションを表します。 読み取り専用です。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[$all](/javascript/api/excel/excel.workbookloadoptions#$all)||
||[application](/javascript/api/excel/excel.workbookloadoptions#application)|このブックを含む Excel アプリケーションインスタンスを表します。|
||[bindings](/javascript/api/excel/excel.workbookloadoptions#bindings)|ブックの一部であるバインドのコレクションを表します。|
||[テーブル](/javascript/api/excel/excel.workbookloadoptions#tables)|ブックに関連付けられているテーブルのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Excel UI でワークシートをアクティブにします。|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|ブックからワークシートを削除します。 ワークシートの可視性が "非常に非表示" に設定されている場合は、削除操作が一般の例外によって失敗することに注意してください。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|行と列の番号に基づいて、1 つのセルを含んだ範囲オブジェクトを取得します。 ワークシートのグリッド内に収まるセルは、親の範囲の境界の外側にある場合があります。|
||[getRange (address?: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|アドレスまたは名前で指定された、セルの単一の四角形のブロックを表す range オブジェクトを取得します。|
||[name](/javascript/api/excel/excel.worksheet#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheet#position)|0 を起点とした、ブック内のワークシートの位置。|
||[管理図](/javascript/api/excel/excel.worksheet#charts)|ワークシートの一部になっているグラフのコレクションを返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.worksheet#id)|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|
||[テーブル](/javascript/api/excel/excel.worksheet#tables)|ワークシートの一部になっているグラフのコレクション。 読み取り専用です。|
||[set (properties: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: WorksheetUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.worksheet#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|ワークシートの可視性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add (name?: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|新しいワークシートをブックに追加します。ワークシートは、既存のワークシートの末尾に追加されます。新しく追加したワークシートをアクティブにする場合は、そのワークシートに対して ".activate() を呼び出します。|
||[getActiveWorksheet ()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|ブックの、現在作業中のワークシートを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|名前または ID を使用して、ワークシート オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ワークシート Collectionloadoptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[$all](/javascript/api/excel/excel.worksheetcollectionloadoptions#$all)||
||[管理図](/javascript/api/excel/excel.worksheetcollectionloadoptions#charts)|コレクション内の各アイテムについて: ワークシートの一部であるグラフのコレクションを返します。|
||[id](/javascript/api/excel/excel.worksheetcollectionloadoptions#id)|コレクション内の各アイテムについて: 指定されたブック内のワークシートを一意に識別する値を返します。 この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。 値の取得のみ可能です。|
||[name](/javascript/api/excel/excel.worksheetcollectionloadoptions#name)|コレクション内の各アイテムについて: ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheetcollectionloadoptions#position)|コレクション内の各アイテムについて: ブック内のワークシートの0から始まる位置。|
||[テーブル](/javascript/api/excel/excel.worksheetcollectionloadoptions#tables)|コレクション内の各項目について、ワークシートの一部であるテーブルのコレクション。|
||[visibility](/javascript/api/excel/excel.worksheetcollectionloadoptions#visibility)|コレクション内の各アイテムについて: ワークシートの可視性。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[管理図](/javascript/api/excel/excel.worksheetdata#charts)|ワークシートの一部になっているグラフのコレクションを返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.worksheetdata#id)|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|
||[name](/javascript/api/excel/excel.worksheetdata#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheetdata#position)|0 を起点とした、ブック内のワークシートの位置。|
||[テーブル](/javascript/api/excel/excel.worksheetdata#tables)|ワークシートの一部になっているグラフのコレクション。 読み取り専用です。|
||[visibility](/javascript/api/excel/excel.worksheetdata#visibility)|ワークシートの可視性。|
|[ワークシート Loadoptions](/javascript/api/excel/excel.worksheetloadoptions)|[$all](/javascript/api/excel/excel.worksheetloadoptions#$all)||
||[管理図](/javascript/api/excel/excel.worksheetloadoptions#charts)|ワークシートの一部になっているグラフのコレクションを返します。|
||[id](/javascript/api/excel/excel.worksheetloadoptions#id)|指定されたブックのワークシートを一意に識別する値を返します。この識別子の値は、ワークシートの名前を変更したり移動したりしても同じままです。値の取得のみ可能です。|
||[name](/javascript/api/excel/excel.worksheetloadoptions#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheetloadoptions#position)|0 を起点とした、ブック内のワークシートの位置。|
||[テーブル](/javascript/api/excel/excel.worksheetloadoptions#tables)|ワークシートの一部になっているグラフのコレクション。|
||[visibility](/javascript/api/excel/excel.worksheetloadoptions#visibility)|ワークシートの可視性。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[name](/javascript/api/excel/excel.worksheetupdatedata#name)|ワークシートの表示名。|
||[position](/javascript/api/excel/excel.worksheetupdatedata#position)|0 を起点とした、ブック内のワークシートの位置。|
||[visibility](/javascript/api/excel/excel.worksheetupdatedata#visibility)|ワークシートの可視性。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
