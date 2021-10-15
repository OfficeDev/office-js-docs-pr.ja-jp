---
title: ExcelJavaScript API 要件セット 1.6
description: ExcelApi 1.6 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 5460a678eb23aeee0be5795856790eb8cccd593e
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367426"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 の新機能

## <a name="conditional-formatting"></a>条件付き書式

範囲の条件付き書式を導入します。 次の種類の条件付き書式を使用できます。

- カラー スケール
- データ バー
- アイコン セット
- カスタム

さらに、次の機能も使用できます。

- 条件付き書式が適用された範囲を返す。
- 条件付き書式を削除する。
- 優先度と機能を `stopifTrue` 提供します。
- 指定した範囲のすべての条件付き書式のコレクションを取得する。
- 現在指定している範囲でアクティブなすべての条件付き書式をクリアする。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.6 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.6 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendApiCalculationUntilNextSync__)|次の値が呼び出されるまで `context.sync()` 計算を中断します。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|この条件付き形式のルール オブジェクトを指定します。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|カラー スケールの条件。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threeColorScale)|場合、カラー スケールには 3 つのポイント (最小、中点、最大値) が含め、それ以外の場合は 2 `true` つ (最小、最大値) になります。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|必要に応じて、条件付き書式ルールを評価する数式。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|必要に応じて、条件付き書式ルールを評価する数式。|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|セル値の条件付き書式の演算子。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|カラー スケール条件の最大ポイント。|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|カラー スケールの基準の中間点 (カラー スケールが 3 色スケールの場合)。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|色スケール基準の最小点。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色スケールの色の HTML カラー コード表現 (赤を表#FF0000など)。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|数値、数式、または `null` (if `type` `lowestValue` is)|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件式の基準の基になる条件。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#borderColor)|境界線の色を表す HTML カラー コード (#RRGGBB"FFA500" など) または名前の付いた HTML 色 ("オレンジ色" など) です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillColor)|塗りつぶしの色を表す HTML カラー コード (#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("オレンジ色" など) として指定します。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveBorderColor)|負のデータ バーが正のデータ バーと同じ罫線の色を持っている場合に指定します。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveFillColor)|負のデータ バーが正のデータ バーと同じ塗りつぶし色を持つ場合に指定します。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#borderColor)|境界線の色を表す HTML カラー コード (#RRGGBB"FFA500" など) または名前の付いた HTML 色 ("オレンジ色" など) です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillColor)|塗りつぶしの色を表す HTML カラー コード (#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("オレンジ色" など) として指定します。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientFill)|データ バーにグラデーションが設定されている場合に指定します。|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|必要に応じて、データ バー ルールを評価する数式。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|データ バーのルールの種類。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#cellValue)|現在の条件付き書式が型の場合、セル値の条件付き書式プロパティを返 `CellValue` します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellValueOrNullObject)|現在の条件付き書式が型の場合、セル値の条件付き書式プロパティを返 `CellValue` します。|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorScale)|現在の条件付き書式が型の場合は、色スケールの条件付き書式プロパティを返 `ColorScale` します。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorScaleOrNullObject)|現在の条件付き書式が型の場合は、色スケールの条件付き書式プロパティを返 `ColorScale` します。|
||[カスタム](/javascript/api/excel/excel.conditionalformat#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customOrNullObject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#dataBar)|現在の条件付き書式がデータ バーの場合は、データ バーのプロパティを返します。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#dataBarOrNullObject)|現在の条件付き書式がデータ バーの場合は、データ バーのプロパティを返します。|
||[delete()](/javascript/api/excel/excel.conditionalformat#delete__)|この条件付き書式を削除します。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getRange__)|条件付き書式が適用された範囲を返す。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getRangeOrNullObject__)|conditonal 形式が適用される範囲を返します。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconSet)|現在の条件付き書式が型の場合は、アイコン セットの条件付き書式プロパティを返 `IconSet` します。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconSetOrNullObject)|現在の条件付き書式が型の場合は、アイコン セットの条件付き書式プロパティを返 `IconSet` します。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|現在の条件付き書式の優先度 `ConditionalFormatCollection` です。|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|事前設定された条件の条件付き書式を返します。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetOrNullObject)|事前設定された条件の条件付き書式を返します。|
||[優先度](/javascript/api/excel/excel.conditionalformat#priority)|この条件付き書式が現在存在する条件付き書式コレクション内の優先度 (またはインデックス)。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopIfTrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textComparison)|現在の条件付き書式がテキスト型の場合は、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textComparisonOrNullObject)|現在の条件付き書式がテキスト型の場合は、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topBottom)|現在の条件付き書式が型の場合は、上/下の条件付き書式プロパティを返 `TopBottom` します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topBottomOrNullObject)|現在の条件付き書式が型の場合は、上/下の条件付き書式プロパティを返 `TopBottom` します。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|条件付き書式の種類。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add_type_)|最初/最優先で新しい条件付き書式をコレクションに追加します。|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearAll__)|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getCount__)|ブック内の条件付き書式の数を返します。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItem_id_)|指定された ID に対応する条件付き書式を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getItemAt_index_)|指定されたインデックスに条件付き書式を返します。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|必要に応じて、条件付き書式ルールを評価する数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulaLocal)|必要に応じて、ユーザーの言語で条件付き書式ルールを評価する数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formulaR1C1)|必要に応じて、R1C1 スタイル表記で条件付き書式ルールを評価する数式。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customIcon)|既定のアイコン セットと異なる場合は、現在の条件のカスタム アイコンが `null` 返されます。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|種類によっては数値または数式。|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` または `greaterThanOrEqual` 、アイコンの条件付き書式のルールの種類ごとに指定します。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|アイコンの条件式は次のものに基づいています。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[条件](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件付き書式の条件。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|境界線の色を表す HTML カラー コード (#RRGGBB"FFA500" など) または名前の付いた HTML 色 ("オレンジ色" など) です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideIndex)|罫線の特定の辺を表す定数値。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|下の罫線を取得します。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。|
||[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getItem_index_)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getItemAt_index_)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|左側の罫線を取得します。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|右の罫線を取得します。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|上の罫線を取得します。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear__)|塗りつぶしをリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|塗りつぶしの色を表す HTML カラー コード (#RRGGBB 形式 ("FFA500"など)、または名前付き HTML 色 ("オレンジ色" など) として指定します。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|フォントが太字の場合に指定します。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear__)|フォントの書式設定をリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|テキストの色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|フォントが italic の場合に指定します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|フォントの取り消し線の状態を指定します。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|フォントに適用される下線の種類。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|条件付き書式範囲全体に適用される罫線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|条件付き書式の範囲全体で定義されている fill オブジェクトを返します。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|条件付き書式の範囲全体で定義されているフォント オブジェクトを返します。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberFormat)|指定したExcelの数値書式コードを表します。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|テキストの条件付き書式の演算子。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件付き書式のテキスト値。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|上または下のランクに基づいて値を書式設定します。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.customconditionalformat#rule)|この条件付き `Rule` 形式のオブジェクトを指定します。|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axisColor)|軸線の色を表す HTML カラー コード (#RRGGBB ("FFA500" など) または名前付き HTML 色 ("オレンジ色" など) です。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisFormat)|データ バーに対して軸がどのように決定Excel表現します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#barDirection)|データ バー グラフィックの基になる方向を指定します。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerBoundRule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeFormat)|データ バー内の軸の左側のすべての値Excel表現します。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveFormat)|データ バー内の軸の右側のすべての値Excel表示します。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showDataBarOnly)|If `true` は、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperBoundRule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|ルールの条件とアイコン セットの配列と、条件付きアイコンの潜在的なカスタム アイコン。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseIconOrder)|場合 `true` は、アイコン セットのアイコンの順序を反転します。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showIconOnly)|場合 `true` は、値を非表示にし、アイコンのみを表示します。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|設定されている場合は、条件付き書式のアイコン セット オプションを表示します。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件付き書式のルール。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate__)|ワークシート上のセルの範囲を計算します。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalFormats)|範囲と `ConditionalFormats` 交差するコレクション。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.textconditionalformat#rule)|条件付き書式のルール。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.topbottomconditionalformat#rule)|上/下の条件付き書式の条件。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate_markAllDirty_)|ワークシート上のすべてのセルを計算します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
