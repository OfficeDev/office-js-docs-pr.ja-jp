---
title: Excel JavaScript API 要件セット1.7
description: ExcelApi 1.7 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c84d099982225bae11cb3deba8a0503da0695aed
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771989"
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
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[罫線](/javascript/api/excel/excel.chartareaformatdata#border)|グラフエリアの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。 読み取り専用です。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[罫線](/javascript/api/excel/excel.chartareaformatloadoptions#border)|グラフエリアの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[罫線](/javascript/api/excel/excel.chartareaformatupdatedata#border)|グラフエリアの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (型: "Invalid" \| "Category" \| "Value" \| "Series", group?: "Primary" \| "Secondary")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|種類とグループで識別された特定の軸を返します。|
||[getItem (type: ChartAxisType, group?: ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|種類とグループで識別された特定の軸を返します。|
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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|指定した軸に対応するグループを表します。 詳細については、「ChartAxisGroup」を参照してください。 読み取り専用です。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|指定された項目軸の基本単位を返すか設定します。|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|項目軸の種類を返すか設定します。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|カスタム軸の表示単位の値を表します。 読み取り専用です。 このプロパティを設定するには、SetCustomDisplayUnit(double) メソッドを使用してください。|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|軸の表示単位を表します。 詳細については、「ChartAxisDisplayUnit」を参照してください。|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|グラフ軸の高さ (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|軸の左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[logBase](/javascript/api/excel/excel.chartaxisdata#logbase)|対数目盛りを使用する場合の対数の底を表します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|指定した軸の目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の目盛のスケール値を返すか設定します。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|指定した軸の補助目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の補助目盛のスケール値を返すか設定します。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|Microsoft Excel でデータ ポイントを最後から最初への順でプロットするかどうかを表します。|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|数値軸のスケールの種類を表します。 詳細については、「ChartAxisScaleType」を参照してください。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|軸の表示単位のラベルを表示するかどうかを表します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|指定した軸での目盛ラベルの位置を示します。 詳細については、「ChartAxisTickLabelPosition」を参照してください。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|目盛ラベル間の項目または系列の数を表します。 1 から 31999 の範囲内で値を設定できます。自動的に設定する場合は、空の文字列にします。 戻り値は常に数値です。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|目盛間の項目または系列の数を表します。|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|軸の上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[type](/javascript/api/excel/excel.chartaxisdata#type)|軸の種類を表します。 詳細については、「ChartAxisType」を参照してください。|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|軸を表示するかどうかを表すブール値。|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|グラフ軸の幅 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|指定した軸に対応するグループを表します。 詳細については、「ChartAxisGroup」を参照してください。 読み取り専用です。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|指定された項目軸の基本単位を返すか設定します。|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|項目軸の種類を返すか設定します。|
||[越える](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[使用されていない場合、既存のファーストパーティのソリューションでのバックセーフのため保持されます。 代わりに、 `Position`を使用してください。|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[使用されていない場合、既存のファーストパーティのソリューションでのバックセーフのため保持されます。 代わりに、 `PositionAt`を使用してください。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|カスタム軸の表示単位の値を表します。 読み取り専用です。 このプロパティを設定するには、SetCustomDisplayUnit(double) メソッドを使用してください。|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|軸の表示単位を表します。 詳細については、「ChartAxisDisplayUnit」を参照してください。|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|グラフ軸の高さ (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|軸の左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[logBase](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|対数目盛りを使用する場合の対数の底を表します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|指定した軸の目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の目盛のスケール値を返すか設定します。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|指定した軸の補助目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の補助目盛のスケール値を返すか設定します。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|Microsoft Excel でデータ ポイントを最後から最初への順でプロットするかどうかを表します。|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|数値軸のスケールの種類を表します。 詳細については、「ChartAxisScaleType」を参照してください。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|軸の表示単位のラベルを表示するかどうかを表します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|指定した軸での目盛ラベルの位置を示します。 詳細については、「ChartAxisTickLabelPosition」を参照してください。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|目盛ラベル間の項目または系列の数を表します。 1 から 31999 の範囲内で値を設定できます。自動的に設定する場合は、空の文字列にします。 戻り値は常に数値です。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|目盛間の項目または系列の数を表します。|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|軸の上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
||[type](/javascript/api/excel/excel.chartaxisloadoptions#type)|軸の種類を表します。 詳細については、「ChartAxisType」を参照してください。|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|軸を表示するかどうかを表すブール値。|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|グラフ軸の幅 (ポイント数) を表します。 軸が非表示の場合は Null。 読み取り専用です。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[baseTimeUnit](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|指定された項目軸の基本単位を返すか設定します。|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|項目軸の種類を返すか設定します。|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|軸の表示単位を表します。 詳細については、「ChartAxisDisplayUnit」を参照してください。|
||[logBase](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|対数目盛りを使用する場合の対数の底を表します。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|指定した軸の目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の目盛のスケール値を返すか設定します。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|指定した軸の補助目盛の種類を表します。 詳細については、「ChartAxisTickMark」を参照してください。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|CategoryType プロパティが lTimeScale に設定されている場合、項目軸の補助目盛のスケール値を返すか設定します。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|Microsoft Excel でデータ ポイントを最後から最初への順でプロットするかどうかを表します。|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|数値軸のスケールの種類を表します。 詳細については、「ChartAxisScaleType」を参照してください。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|軸の表示単位のラベルを表示するかどうかを表します。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|指定した軸での目盛ラベルの位置を示します。 詳細については、「ChartAxisTickLabelPosition」を参照してください。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|目盛ラベル間の項目または系列の数を表します。 1 から 31999 の範囲内で値を設定できます。自動的に設定する場合は、空の文字列にします。 戻り値は常に数値です。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|目盛間の項目または系列の数を表します。|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|軸を表示するかどうかを表すブール値。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|罫線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[set (プロパティ: Excel. ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartBorderUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartborder#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartBorderData](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|罫線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartBorderLoadOptions](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|罫線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartBorderUpdateData](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|グラフの罫線の色を表す HTML カラー コード。|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|罫線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|罫線の太さ (ポイント数) を表します。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|コレクション内の各アイテムについて: グラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|コレクション内の各アイテムについて: chart の一意の id。 読み取り専用です。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|コレクション内の各アイテムについて: ピボットグラフのすべてのフィールドボタンを表示するかどうかを表します。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|グラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[id](/javascript/api/excel/excel.chartdata#id)|グラフの一意の ID。 読み取り専用です。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|ピボットグラフにすべてのフィールド ボタンを表示するかどうかを示します。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabel#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[set (properties: ChartDataLabel)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartDataLabelUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabeldata#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|データ ラベルの位置を表す DataLabelPosition 値。 詳細については、「ChartDataLabelPosition」を参照してください。|
||[記号](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|グラフのデータ ラベルに使用される区切り文字を表す文字列。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|データ ラベルのバブルのサイズを表示または非表示にするかを表すブール値。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|データ ラベルのカテゴリ名を表示するか非表示にするかを表すブール値。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|データ ラベルの凡例マーカーを表示するか非表示にするかを表すブール値。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|データ ラベルのパーセンテージを表示するか非表示にするかを表すブール値。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|データ ラベルの系列名を表示するか非表示にするかを表すブール値。|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|データ ラベルの値を表示するか非表示にするかを表すブール値。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|フォント名、フォントサイズ、色など、グラフの文字オブジェクトのフォントの属性を表します。|
||[set (プロパティ: Excel. ChartFormatString)](/javascript/api/excel/excel.chartformatstring#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartFormatStringUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartFormatStringData](/javascript/api/excel/excel.chartformatstringdata)|[font](/javascript/api/excel/excel.chartformatstringdata#font)|フォント名、フォントサイズ、色など、グラフの文字オブジェクトのフォントの属性を表します。|
|[ChartFormatStringLoadOptions](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartformatstringloadoptions#font)|フォント名、フォントサイズ、色など、グラフの文字オブジェクトのフォントの属性を表します。|
|[ChartFormatStringUpdateData](/javascript/api/excel/excel.chartformatstringupdatedata)|[font](/javascript/api/excel/excel.chartformatstringupdatedata#font)|フォント名、フォントサイズ、色など、グラフの文字オブジェクトのフォントの属性を表します。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|グラフの凡例の高さをポイント単位で表します。 凡例が表示されていない場合は Null。|
||[left](/javascript/api/excel/excel.chartlegend#left)|グラフの凡例の左のポイントを表します。 凡例が表示されていない場合は Null。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|凡例に含まれる凡例エントリのコレクションを表します。 読み取り専用です。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|凡例がグラフに影付きかどうかを表します。|
||[top](/javascript/api/excel/excel.chartlegend#top)|グラフ凡例の上を表します。|
||[width](/javascript/api/excel/excel.chartlegend#width)|グラフの凡例の幅をポイント単位で表します。 凡例が表示されていない場合は Null。|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|グラフの凡例の高さをポイント単位で表します。 凡例が表示されていない場合は Null。|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|グラフの凡例の左のポイントを表します。 凡例が表示されていない場合は Null。|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|凡例に含まれる凡例エントリのコレクションを表します。 読み取り専用です。|
||[showShadow](/javascript/api/excel/excel.chartlegenddata#showshadow)|凡例がグラフに影付きかどうかを表します。|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|グラフ凡例の上を表します。|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|グラフの凡例の幅をポイント単位で表します。 凡例が表示されていない場合は Null。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[set (properties: ChartLegendEntry)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartLegendEntryUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|コレクションに含まれる凡例エントリの数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|指定されたインデックスに位置する凡例エントリを返します。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|コレクション内の各アイテムについて: グラフの凡例項目の表示を表します。|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendEntryUpdateData](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|グラフの凡例エントリを表示するかどうかを表します。|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|グラフの凡例の高さをポイント単位で表します。 凡例が表示されていない場合は Null。|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|グラフの凡例の左のポイントを表します。 凡例が表示されていない場合は Null。|
||[showShadow](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|凡例がグラフに影付きかどうかを表します。|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|グラフ凡例の上を表します。|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|グラフの凡例の幅をポイント単位で表します。 凡例が表示されていない場合は Null。|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|グラフの凡例の高さをポイント単位で表します。 凡例が表示されていない場合は Null。|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|グラフの凡例の左のポイントを表します。 凡例が表示されていない場合は Null。|
||[showShadow](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|凡例がグラフに影付きかどうかを表します。|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|グラフ凡例の上を表します。|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|グラフの凡例の幅をポイント単位で表します。 凡例が表示されていない場合は Null。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|線の太さ (ポイント数) を表します。|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|線の太さ (ポイント数) を表します。|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|線の太さ (ポイント数) を表します。|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|線のスタイルを表します。 詳細については、「Excel の ChartLineStyle」を参照してください。|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|線の太さ (ポイント数) を表します。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|グラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|グラフの一意の ID。 読み取り専用です。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|ピボットグラフにすべてのフィールド ボタンを表示するかどうかを示します。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|データ ポイントのマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|グラフ データ ポイントのデータ ラベルを返します。 読み取り専用です。|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|グラフ データ ポイントのデータ ラベルを返します。 読み取り専用です。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpointdata#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpointdata#markerstyle)|データ ポイントのマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[罫線](/javascript/api/excel/excel.chartpointformat#border)|色、スタイル、および重みの情報を含むグラフデータポイントの罫線の書式を表します。 読み取り専用です。|
|[ChartPointFormatData](/javascript/api/excel/excel.chartpointformatdata)|[罫線](/javascript/api/excel/excel.chartpointformatdata#border)|色、スタイル、および重みの情報を含むグラフデータポイントの罫線の書式を表します。 読み取り専用です。|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[罫線](/javascript/api/excel/excel.chartpointformatloadoptions#border)|色、スタイル、および重みの情報を含むグラフデータポイントの罫線の書式を表します。|
|[ChartPointFormatUpdateData](/javascript/api/excel/excel.chartpointformatupdatedata)|[罫線](/javascript/api/excel/excel.chartpointformatupdatedata#border)|色、スタイル、および重みの情報を含むグラフデータポイントの罫線の書式を表します。|
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|グラフ データ ポイントのデータ ラベルを返します。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpointloadoptions#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|データ ポイントのマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|グラフ データ ポイントのデータ ラベルを返します。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|データ ポイントのマーカー背景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|データ ポイントのマーカー前景色を表す HTML カラー コード。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpointupdatedata#markersize)|データ ポイントのマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|データ ポイントのマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|コレクション内の各項目について: グラフ要素のデータラベルを返します。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|コレクション内の各アイテムについて: データポイントにデータラベルがあるかどうかを表します。 等高線グラフには適用されません。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|コレクション内の各項目について: データ要素のマーカーの背景色の HTML 色コード表現。 例: #FF0000 は赤を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|コレクション内の各アイテムについて: データポイントのマーカーの前景色を HTML カラーコードで表現したもの。 例: #FF0000 は赤を表します。|
||[markerSize](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|コレクション内の各アイテムについて: データポイントのマーカーサイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|コレクション内の各アイテムについて: グラフデータポイントのマーカースタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
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
|[Chartcharts Collectionloadoptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|コレクション内の各アイテムについて: 系列のグラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|コレクション内の各アイテムについて: グラフ系列のドーナツの穴のサイズを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|
||[対象](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|コレクション内の各項目について: 系列がフィルター処理されているかどうかを表すブール型 (Boolean) の値。 等高線グラフには適用されません。|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|コレクション内の各アイテムについて: グラフ系列の間隔幅を表します。  横棒グラフと縦棒グラフでのみ有効です。|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|コレクション内の各項目について: 系列にデータラベルがあるかどうかを表すブール型 (Boolean) の値。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|コレクション内の各アイテムについて: グラフ系列のマーカーの背景色を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|コレクション内の各項目について: グラフ系列のマーカー前の色を表します。|
||[markerSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|コレクション内の各アイテムについて: グラフ系列のマーカーサイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|コレクション内の各アイテムについて: グラフ系列のマーカースタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[plotOrder](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|コレクション内の各アイテムについて: グラフグループ内のグラフ系列のプロット順序を表します。|
||[showShadow](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|コレクション内の各項目について: 系列に影があるかどうかを表すブール型 (Boolean) の値。|
||[smooth](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|コレクション内の各項目について: 系列が滑らかさであるかどうかを表すブール型 (Boolean) の値。 折れ線グラフおよび散布図にのみ適用されます。|
|[Chart系列データ](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|グラフ系列の種類を表します。 詳細については、「ChartType」を参照してください。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|グラフ系列のドーナツの穴の大きさを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|
||[対象](/javascript/api/excel/excel.chartseriesdata#filtered)|データ系列がフィルター処理されるかどうかを表すブール値。 等高線グラフには適用されません。|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|グラフ系列間に設けられる間隔を表します。  横棒グラフと縦棒グラフでのみ有効です。|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|系列のデータ ラベルの有無を表すブール値。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|グラフ系列のマーカー背景色を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|グラフ系列のマーカー前景色を表します。|
||[markerSize](/javascript/api/excel/excel.chartseriesdata#markersize)|グラフ系列のマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartseriesdata#markerstyle)|グラフ系列のマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[plotOrder](/javascript/api/excel/excel.chartseriesdata#plotorder)|グラフ グループ内でのグラフ系列のプロット順序を表します。|
||[showShadow](/javascript/api/excel/excel.chartseriesdata#showshadow)|系列に影があるかどうかを表すブール型 (Boolean) の値を指定します。|
||[smooth](/javascript/api/excel/excel.chartseriesdata#smooth)|系列が平滑化されるかどうかを表すブール値。 折れ線グラフおよび散布図にのみ適用されます。|
||[曲線](/javascript/api/excel/excel.chartseriesdata#trendlines)|データ系列に含まれる近似曲線のコレクションを表します。 読み取り専用です。|
|[Chart系列 Loadoptions](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|グラフ系列の種類を表します。 詳細については、「ChartType」を参照してください。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|グラフ系列のドーナツの穴の大きさを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|
||[対象](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|データ系列がフィルター処理されるかどうかを表すブール値。 等高線グラフには適用されません。|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|グラフ系列間に設けられる間隔を表します。  横棒グラフと縦棒グラフでのみ有効です。|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|系列のデータ ラベルの有無を表すブール値。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|グラフ系列のマーカー背景色を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|グラフ系列のマーカー前景色を表します。|
||[markerSize](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|グラフ系列のマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|グラフ系列のマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[plotOrder](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|グラフ グループ内でのグラフ系列のプロット順序を表します。|
||[showShadow](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|系列に影があるかどうかを表すブール型 (Boolean) の値を指定します。|
||[smooth](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|系列が平滑化されるかどうかを表すブール値。 折れ線グラフおよび散布図にのみ適用されます。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|グラフ系列の種類を表します。 詳細については、「ChartType」を参照してください。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|グラフ系列のドーナツの穴の大きさを表します。  ドーナツ グラフと doughnutExploded グラフでのみ有効です。|
||[対象](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|データ系列がフィルター処理されるかどうかを表すブール値。 等高線グラフには適用されません。|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|グラフ系列間に設けられる間隔を表します。  横棒グラフと縦棒グラフでのみ有効です。|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|系列のデータ ラベルの有無を表すブール値。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|グラフ系列のマーカー背景色を表します。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|グラフ系列のマーカー前景色を表します。|
||[markerSize](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|グラフ系列のマーカー サイズを表します。|
||[markerStyle](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|グラフ系列のマーカー スタイルを表します。 詳細については、「ChartMarkerStyle」を参照してください。|
||[plotOrder](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|グラフ グループ内でのグラフ系列のプロット順序を表します。|
||[showShadow](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|系列に影があるかどうかを表すブール型 (Boolean) の値を指定します。|
||[smooth](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|系列が平滑化されるかどうかを表すブール値。 折れ線グラフおよび散布図にのみ適用されます。|
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
|[Chartタイトルデータ](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|グラフ タイトルの高さ (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|グラフ タイトルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttitledata#left)|グラフ タイトルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[position](/javascript/api/excel/excel.charttitledata#position)|グラフ タイトルの位置を表します。 詳細については、「Excel Charttitle Position」を参照してください。|
||[showShadow](/javascript/api/excel/excel.charttitledata#showshadow)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|グラフ タイトルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttitledata#top)|グラフ タイトルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|グラフ タイトルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.charttitledata#width)|グラフ タイトルの幅 (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[罫線](/javascript/api/excel/excel.charttitleformat#border)|グラフタイトルの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。 読み取り専用です。|
|[Chartタイトル Formatdata](/javascript/api/excel/excel.charttitleformatdata)|[罫線](/javascript/api/excel/excel.charttitleformatdata#border)|グラフタイトルの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。 読み取り専用です。|
|[Chartタイトル Formatloadoptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[罫線](/javascript/api/excel/excel.charttitleformatloadoptions#border)|グラフタイトルの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[罫線](/javascript/api/excel/excel.charttitleformatupdatedata#border)|グラフタイトルの罫線の書式を表します。これには、色、linestyle、およびウエイトが含まれます。|
|[Chartタイトル Loadoptions](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|グラフ タイトルの高さ (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|グラフ タイトルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|グラフ タイトルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|グラフ タイトルの位置を表します。 詳細については、「Excel Charttitle Position」を参照してください。|
||[showShadow](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|グラフ タイトルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|グラフ タイトルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|グラフ タイトルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|グラフ タイトルの幅 (ポイント数) を返します。 グラフのタイトルが表示されていない場合は Null。 読み取り専用です。|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|グラフ タイトルの水平方向の配置を表します。|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|グラフ タイトルの左端からグラフ エリアの左端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|グラフ タイトルの位置を表します。 詳細については、「Excel Charttitle Position」を参照してください。|
||[showShadow](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|グラフ タイトルが影付きにされるかどうかを指定するブール値を表します。|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|グラフ タイトルのテキストの向きを表します。 値は -90 から 90 の範囲内の整数か、縦書きテキストの場合は 180 でなければなりません。|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|グラフ タイトルの上端からグラフ エリアの上端までの距離 (ポイント数) を表します。 グラフのタイトルが表示されていない場合は Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|グラフ タイトルの垂直方向の配置を表します。 詳細については、「Excel Charttext縦書きの配置」を参照してください。|
|[ChartTrendline 曲線](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|trendline オブジェクトを削除します。|
||[y](/javascript/api/excel/excel.charttrendline#intercept)|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendline#name)|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[format](/javascript/api/excel/excel.charttrendline#format)|グラフの近似曲線の書式設定を表します。|
||[set (プロパティ: Excel. ChartTrendline 曲線)](/javascript/api/excel/excel.charttrendline#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTrendlineUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[type](/javascript/api/excel/excel.charttrendline#type)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add (type?: "Linear" \| "指数" \| "対数" \| "movingaverage" \| "多項式" \| "Power")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|近似曲線のコレクションに新しい近似曲線を追加します。|
||[add (type?: ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|近似曲線のコレクションに新しい近似曲線を追加します。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|コレクションに含まれる近似曲線の数を返します。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|インデックス (項目配列内の挿入順序) に基づいて trendline オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|コレクション内の各項目について: グラフの近似曲線の書式設定を表します。|
||[y](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|コレクション内の各アイテムについて: 近似曲線の切片値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|コレクション内の各アイテムについて: グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|コレクション内の各アイテムについて: 近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|コレクション内の各項目について: グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[type](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|コレクション内の各項目について: グラフの近似曲線の種類を表します。|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|グラフの近似曲線の書式設定を表します。|
||[y](/javascript/api/excel/excel.charttrendlinedata#intercept)|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[type](/javascript/api/excel/excel.charttrendlinedata#type)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|グラフの線の書式設定を表します。 値の取得のみ可能です。|
||[set (properties: ChartTrendlineFormat)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartTrendlineFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartTrendlineFormatData](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|グラフの線の書式設定を表します。 値の取得のみ可能です。|
|[ChartTrendlineFormatLoadOptions](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|グラフの線の書式設定を表します。|
|[ChartTrendlineFormatUpdateData](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|グラフの線の書式設定を表します。|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|グラフの近似曲線の書式設定を表します。|
||[y](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[type](/javascript/api/excel/excel.charttrendlineloadoptions#type)|グラフの近似曲線の種類を表します。|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|グラフの近似曲線の書式設定を表します。|
||[y](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|近似曲線の切片の値を表します。 数値または空の文字列を設定できます (値を自動的に設定する場合)。 戻り値は常に数値です。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|グラフの近似曲線の期間を表します。 MovingAverage 型の近似曲線にのみ適用されます。|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|近似曲線の名前を表します。 文字列値または null 値 (値を自動的に設定する場合) に設定できます。 戻り値は常に文字列です。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|グラフの近似曲線の順序を表します。 多項式型の近似曲線にのみ適用されます。|
||[type](/javascript/api/excel/excel.charttrendlineupdatedata#type)|グラフの近似曲線の種類を表します。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|グラフの種類を表します。 詳細については、「ChartType」を参照してください。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|ピボットグラフにすべてのフィールド ボタンを表示するかどうかを示します。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.customproperty#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.customproperty#type)|カスタム プロパティの値の型を取得します。 読み取り専用。|
||[set (properties: Excel. CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: CustomPropertyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.customproperty#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[value](/javascript/api/excel/excel.customproperty#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomPropertyCollectionLoadOptions](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|コレクション内の各アイテムについて: カスタムプロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|コレクション内の各アイテムについて: カスタムプロパティの値の種類を取得します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|コレクション内の各アイテムについて: カスタムプロパティの値を取得または設定します。|
|[CustomPropertyData](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.custompropertydata#type)|カスタム プロパティの値の型を取得します。 読み取り専用。|
||[value](/javascript/api/excel/excel.custompropertydata#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyLoadOptions](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.custompropertyloadoptions#type)|カスタム プロパティの値の型を取得します。 読み取り専用。|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyUpdateData](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|カスタム プロパティの値を取得または設定します。|
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
||[set (properties: DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: DocumentPropertiesUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|ブックの件名を取得または設定します。|
||[title](/javascript/api/excel/excel.documentproperties#title)|ブックのタイトルを取得または設定します。|
|[DocumentPropertiesData](/javascript/api/excel/excel.documentpropertiesdata)|[判別](/javascript/api/excel/excel.documentpropertiesdata#author)|ブックの作成者を取得または設定します。|
||[項目](/javascript/api/excel/excel.documentpropertiesdata#category)|ブックのカテゴリを取得または設定します。|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|ブックのコメントを取得または設定します。|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|ブックの会社を取得または設定します。|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|ブックの作成日を取得します。 読み取り専用です。|
||[配色](/javascript/api/excel/excel.documentpropertiesdata#custom)|ブックのカスタム プロパティのコレクションを取得します。 読み取り専用です。|
||[キーワード](/javascript/api/excel/excel.documentpropertiesdata#keywords)|ブックのキーワードを取得または設定します。|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|ブックの最後の作成者を取得します。 読み取り専用です。|
||[manager](/javascript/api/excel/excel.documentpropertiesdata#manager)|ブックのマネージャーを取得または設定します。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|ブックのリビジョン番号を取得します。 読み取り専用です。|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|ブックの件名を取得または設定します。|
||[title](/javascript/api/excel/excel.documentpropertiesdata#title)|ブックのタイトルを取得または設定します。|
|[DocumentPropertiesLoadOptions](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[判別](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|ブックの作成者を取得または設定します。|
||[項目](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|ブックのカテゴリを取得または設定します。|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|ブックのコメントを取得または設定します。|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|ブックの会社を取得または設定します。|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|ブックの作成日を取得します。 読み取り専用です。|
||[キーワード](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|ブックのキーワードを取得または設定します。|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|ブックの最後の作成者を取得します。 読み取り専用です。|
||[manager](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|ブックのマネージャーを取得または設定します。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|ブックのリビジョン番号を取得します。 読み取り専用です。|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|ブックの件名を取得または設定します。|
||[title](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|ブックのタイトルを取得または設定します。|
|[DocumentPropertiesUpdateData](/javascript/api/excel/excel.documentpropertiesupdatedata)|[判別](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|ブックの作成者を取得または設定します。|
||[項目](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|ブックのカテゴリを取得または設定します。|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|ブックのコメントを取得または設定します。|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|ブックの会社を取得または設定します。|
||[キーワード](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|ブックのキーワードを取得または設定します。|
||[manager](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|ブックのマネージャーを取得または設定します。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|ブックのリビジョン番号を取得します。 読み取り専用です。|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|ブックの件名を取得または設定します。|
||[title](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|ブックのタイトルを取得または設定します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|名前付きアイテムの値と型を含むオブジェクトを返します。 読み取り専用です。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|名前付きアイテムの配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[Nameditemarraypct データ](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|名前付きアイテムの配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[Nameditemarraypct Loadoptions](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|名前付きアイテムの配列内の各アイテムの型を表します。|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|名前付きアイテムの配列に含まれる各アイテムの値を表します。読み取り専用。|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|コレクション内の各アイテムについて: 名前付きアイテムの値と型を含むオブジェクトを返します。|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|コレクション内の各アイテムについて: 名前付きアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[arrayValues](/javascript/api/excel/excel.nameditemdata#arrayvalues)|名前付きアイテムの値と型を含むオブジェクトを返します。 読み取り専用です。|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|名前付きアイテムの値と型を含むオブジェクトを返します。|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|名前付きのアイテムの数式を取得または設定します。  数式は常に '=' 記号で始まります。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: number, Numrows: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|現在の Range オブジェクトと左上のセルが同じで、指定した数の行と列を含む Range オブジェクトを取得します。|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|範囲を base64 でエンコードされた png 画像としてレンダリングします。|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|指定された範囲の左上のセルを囲む領域を表す Range オブジェクトを返します。 周囲の領域は、この範囲に相対の空白の行と空白の列の任意の組み合わせで囲まれた範囲です。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|アクティブ セルに多数の値が含まれる場合、そのセルのカードを表示します。|
||[style](/javascript/api/excel/excel.range#style)|現在の範囲のスタイルを表します。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hyperlink](/javascript/api/excel/excel.rangedata#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|
||[style](/javascript/api/excel/excel.rangedata#style)|現在の範囲のスタイルを表します。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|該当する範囲内のすべてのセルのテキストの向きを設定します。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Range オブジェクトの列の幅が、シートの標準の幅と等しいかどうかを示します。|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|該当する範囲内のすべてのセルのテキストの向きを設定します。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|Range オブジェクトの列の幅が、シートの標準の幅と等しいかどうかを示します。|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|該当する範囲内のすべてのセルのテキストの向きを設定します。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|Range オブジェクトの列の幅が、シートの標準の幅と等しいかどうかを示します。|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|該当する範囲内のすべてのセルのテキストの向きを設定します。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|Range オブジェクトの行の高さを、シートの標準の高さと等しくするかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|Range オブジェクトの列の幅が、シートの標準の幅と等しいかどうかを示します。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|ハイパーリンクの URL ターゲットを表します。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|ハイパーリンクのドキュメント参照先を表します。|
||[ポップヒント](/javascript/api/excel/excel.rangehyperlink#screentip)|ハイパーリンクの上にカーソルを合わせると表示される文字列を表します。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|該当する範囲内の左上端のセルに表示される文字列を表します。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hyperlink](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|現在の範囲のスタイルを表します。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[hyperlink](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|現在の範囲のハイパーリンクを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|ユーザーの言語で文字列として指定された範囲に対応する、Excel の数値形式コードを表します。|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|現在の範囲のスタイルを表します。|
|[スタイル](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|このスタイルを削除します。|
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
||[set (properties: Excel)](/javascript/api/excel/excel.style#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: StyleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.style#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|スタイルで適用される垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|コレクションに新しいスタイルを追加します。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|名前に基づいてスタイルを取得します。|
||[items](/javascript/api/excel/excel.stylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|コレクション内の各アイテムについて: 4 つの罫線のスタイルを表す、4つの Border オブジェクトの Border コレクション。|
||[Unset](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|コレクション内の各アイテムについて: スタイルが組み込みのスタイルであるかどうかを示します。|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|コレクション内の各アイテムについて: スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.stylecollectionloadoptions#font)|コレクション内の各項目について: スタイルのフォントを表す Font オブジェクト。|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|コレクション内の各アイテムについて: ワークシートが保護されているときに、数式を非表示にするかどうかを示します。|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|コレクション内の各アイテムについて: スタイルの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[includeAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|コレクション内の各アイテムについて: スタイルに、AutoIndent、水平アラインメント、垂直アラインメント、WrapText、IndentLevel、および TextOrientation の各プロパティが含まれているかどうかを示します。|
||[includeBorder](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|コレクション内の各アイテムについて: スタイルに Color、ColorIndex、LineStyle、および Weight の各罫線のプロパティが含まれているかどうかを示します。|
||[includeFont](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|コレクション内の各アイテムについて: スタイルに、Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、打ち消し線、下付き、上付き、および下線のフォントプロパティが含まれているかどうかを示します。|
||[includeNumber](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|コレクション内の各アイテムについて: スタイルに NumberFormat プロパティが含まれているかどうかを示します。|
||[includePatterns](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|コレクション内の各アイテムについて: スタイルに Color、ColorIndex、InvertIfNegative、Pattern、Pattern Color、および PatternColorIndex interior プロパティが含まれているかどうかを示します。|
||[includeProtection](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|コレクション内の各アイテムについて: スタイルに FormulaHidden および Locked protection プロパティが含まれているかどうかを示します。|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|コレクション内の各アイテムについて: スタイルのインデントレベルを示す 0 ~ 250 の整数。|
||[locked](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|コレクション内の各アイテムについて: ワークシートが保護されているときにオブジェクトがロックされているかどうかを示します。|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|コレクション内の各アイテムについて: スタイルの名前。|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|コレクション内の各アイテムについて: スタイルの数値形式の書式設定コード。|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|コレクション内の各アイテムについて: スタイルの数値形式のローカライズされた書式コード。|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|コレクション内の各アイテムについて: スタイルの読み取り順序。|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|コレクション内の各アイテムについて: 使用可能な列幅に合わせて、テキストを自動的に縮小するかどうかを示します。|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|コレクション内の各アイテムについて: スタイルの垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|コレクション内の各アイテムについて: オブジェクト内のテキストを折り返すかどうかを示します。|
|[スタイルデータ](/javascript/api/excel/excel.styledata)|[borders](/javascript/api/excel/excel.styledata#borders)|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。|
||[Unset](/javascript/api/excel/excel.styledata#builtin)|スタイルが組み込みのスタイルであるかどうかを示します。|
||[fill](/javascript/api/excel/excel.styledata#fill)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.styledata#font)|スタイルのフォントを表す Font オブジェクト。|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|スタイルでの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[includeAlignment](/javascript/api/excel/excel.styledata#includealignment)|スタイルに配置のプロパティ (AddIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、および TextOrientation) が含まれるかどうかを示します。|
||[includeBorder](/javascript/api/excel/excel.styledata#includeborder)|スタイルに罫線のプロパティ (Color、ColorIndex、LineStyle、Weight) が含まれているかどうかを示します。|
||[includeFont](/javascript/api/excel/excel.styledata#includefont)|スタイルにフォントのプロパティ (Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript、Underline) が含まれているかどうかを示します。|
||[includeNumber](/javascript/api/excel/excel.styledata#includenumber)|スタイルに NumberFormat プロパティが含まれているかどうかを示します。|
||[includePatterns](/javascript/api/excel/excel.styledata#includepatterns)|スタイルに塗りつぶしのプロパティ (Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、PatternColorIndex) が含まれているかどうかを示します。|
||[includeProtection](/javascript/api/excel/excel.styledata#includeprotection)|スタイルに保護のプロパティ (FormulaHidden および Locked) が含まれているかどうかを示します。|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.styledata#locked)|ワークシートが保護されている場合、オブジェクトがロックされるかどうかを示します。|
||[name](/javascript/api/excel/excel.styledata#name)|スタイルの名前。|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|スタイルで適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|スタイルで適用される垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|
|[スタイル Loadオプション](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[borders](/javascript/api/excel/excel.styleloadoptions#borders)|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。|
||[Unset](/javascript/api/excel/excel.styleloadoptions#builtin)|スタイルが組み込みのスタイルであるかどうかを示します。|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.styleloadoptions#font)|スタイルのフォントを表す Font オブジェクト。|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|スタイルでの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[includeAlignment](/javascript/api/excel/excel.styleloadoptions#includealignment)|スタイルに配置のプロパティ (AddIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、および TextOrientation) が含まれるかどうかを示します。|
||[includeBorder](/javascript/api/excel/excel.styleloadoptions#includeborder)|スタイルに罫線のプロパティ (Color、ColorIndex、LineStyle、Weight) が含まれているかどうかを示します。|
||[includeFont](/javascript/api/excel/excel.styleloadoptions#includefont)|スタイルにフォントのプロパティ (Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript、Underline) が含まれているかどうかを示します。|
||[includeNumber](/javascript/api/excel/excel.styleloadoptions#includenumber)|スタイルに NumberFormat プロパティが含まれているかどうかを示します。|
||[includePatterns](/javascript/api/excel/excel.styleloadoptions#includepatterns)|スタイルに塗りつぶしのプロパティ (Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、PatternColorIndex) が含まれているかどうかを示します。|
||[includeProtection](/javascript/api/excel/excel.styleloadoptions#includeprotection)|スタイルに保護のプロパティ (FormulaHidden および Locked) が含まれているかどうかを示します。|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.styleloadoptions#locked)|ワークシートが保護されている場合、オブジェクトがロックされるかどうかを示します。|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|スタイルの名前。|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|スタイルで適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|スタイルで適用される垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[borders](/javascript/api/excel/excel.styleupdatedata#borders)|4 つの辺の罫線のスタイルを表す、4 つの Border オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|スタイルの塗りつぶし。|
||[font](/javascript/api/excel/excel.styleupdatedata#font)|スタイルのフォントを表す Font オブジェクト。|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|ワークシートが保護されている場合、数式を非表示にするかどうかを示します。|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|スタイルでの水平方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[includeAlignment](/javascript/api/excel/excel.styleupdatedata#includealignment)|スタイルに配置のプロパティ (AddIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel、および TextOrientation) が含まれるかどうかを示します。|
||[includeBorder](/javascript/api/excel/excel.styleupdatedata#includeborder)|スタイルに罫線のプロパティ (Color、ColorIndex、LineStyle、Weight) が含まれているかどうかを示します。|
||[includeFont](/javascript/api/excel/excel.styleupdatedata#includefont)|スタイルにフォントのプロパティ (Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript、Underline) が含まれているかどうかを示します。|
||[includeNumber](/javascript/api/excel/excel.styleupdatedata#includenumber)|スタイルに NumberFormat プロパティが含まれているかどうかを示します。|
||[includePatterns](/javascript/api/excel/excel.styleupdatedata#includepatterns)|スタイルに塗りつぶしのプロパティ (Color、ColorIndex、InvertIfNegative、Pattern、PatternColor、PatternColorIndex) が含まれているかどうかを示します。|
||[includeProtection](/javascript/api/excel/excel.styleupdatedata#includeprotection)|スタイルに保護のプロパティ (FormulaHidden および Locked) が含まれているかどうかを示します。|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|スタイルのインデント レベルを示す 0 から 250 の範囲内の整数。|
||[locked](/javascript/api/excel/excel.styleupdatedata#locked)|ワークシートが保護されている場合、オブジェクトがロックされるかどうかを示します。|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|スタイルで適用される数値形式の表示形式コード。|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|スタイルで適用される数値形式のローカライズされた表示形式コード。|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|スタイルで適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|スタイルで適用される垂直方向の配置を表します。 詳細については、「Excel の配置」を参照してください。|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|Microsoft Excel でオブジェクト内のテキストをラップするかどうかを示します。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|特定の表で、セル内のデータが変更されたときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|特定の表で選択範囲が変更されたときに発生します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|特定のワークシート上のテーブル内で変更されたエリアを表すアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 詳細については、「DataChangeType」を参照してください。|
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
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|ブックの名前を取得します。 読み取り専用です。|
||[プロパティ](/javascript/api/excel/excel.workbookdata#properties)|ブックのプロパティを取得します。 読み取り専用です。|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|ブックの workbookProtection オブジェクトを返します。 読み取り専用です。|
||[線](/javascript/api/excel/excel.workbookdata#styles)|ブックに関連付けられているスタイルのコレクションを表します。 読み取り専用です。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|ブックの名前を取得します。 読み取り専用です。|
||[プロパティ](/javascript/api/excel/excel.workbookloadoptions#properties)|ブックのプロパティを取得します。|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|ブックの workbookProtection オブジェクトを返します。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect (password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|ブックを保護します。 ブックが保護されている場合は失敗します。|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|ブックが保護されているかどうかを示します。 読み取り専用です。|
||[保護の解除 (password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|ブックの保護を解除します。|
|[WorkbookProtectionData](/javascript/api/excel/excel.workbookprotectiondata)|[protected](/javascript/api/excel/excel.workbookprotectiondata#protected)|ブックが保護されているかどうかを示します。 読み取り専用です。|
|[WorkbookProtectionLoadOptions](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[protected](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|ブックが保護されているかどうかを示します。 読み取り専用です。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[プロパティ](/javascript/api/excel/excel.workbookupdatedata#properties)|ブックのプロパティを取得します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy (positionType? \| : "None", Before \| \| \| "" End ", relativeTo?: Excel. ワークシート)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|ワークシートをコピーして、指定した位置に配置します。 コピーしたワークシートを返します。|
||[copy (positionType?: Excel. ワークシートの種類, relativeTo?: Excel)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|ワークシートをコピーして、指定した位置に配置します。 コピーしたワークシートを返します。|
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
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|ブック内のすべてのワークシートがアクティブになったときに発生します。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|新しいワークシートがブックに追加されるときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|ブック内のすべてのワークシートが非アクティブ化されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|ブックからワークシートが削除されるときに発生します。|
|[ワークシート Collectionloadoptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|コレクション内の各アイテムについて: ワークシート内のすべての行の標準 (既定) の高さをポイント単位で返します。 読み取り専用です。|
||[standardWidth](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|コレクション内の各アイテムについて: ワークシート内のすべての列の標準 (既定) 幅を取得または設定します。|
||[tabColor](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|コレクション内の各アイテムについて: ワークシートのタブの色を取得または設定します。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[standardHeight](/javascript/api/excel/excel.worksheetdata#standardheight)|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。 読み取り専用です。|
||[standardWidth](/javascript/api/excel/excel.worksheetdata#standardwidth)|ワークシートのすべての列の標準 (既定) の幅を返すか設定します。|
||[tabColor](/javascript/api/excel/excel.worksheetdata#tabcolor)|ワークシートのタブの色を取得または設定します。|
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
|[ワークシート Loadoptions](/javascript/api/excel/excel.worksheetloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|ワークシート内のすべての行の標準 (既定) の高さ (ポイント数) を返します。 読み取り専用です。|
||[standardWidth](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|ワークシートのすべての列の標準 (既定) の幅を返すか設定します。|
||[tabColor](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|ワークシートのタブの色を取得または設定します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[保護の解除 (password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|ワークシートの保護を解除します。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|オブジェクトの編集を可能にするワークシート保護オプションを表します。|
||[allowEditScenarios シナリオ](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|シナリオの編集を可能にするワークシート保護オプションを表します。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|選択モードのワークシート保護オプションを表します。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|特定のワークシートで選択されたエリアを表す範囲のアドレスを取得します。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|選択範囲が変更されたワークシートの ID を取得します。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[standardWidth](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|ワークシートのすべての列の標準 (既定) の幅を返すか設定します。|
||[tabColor](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|ワークシートのタブの色を取得または設定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
