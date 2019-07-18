---
title: Excel JavaScript API 要件セット1.6
description: ExcelApi 1.6 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e1a3375d19d8c1cb0fbddac50fabf826b96d7cc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771975"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 の新機能

## <a name="conditional-formatting"></a>条件付き書式

範囲の条件付き書式が導入されています。 次の種類の条件付き書式を使用できます。

* カラー スケール
* データ バー
* アイコン セット
* カスタム

さらに、次の機能も使用できます。

* 条件付き書式が適用された範囲を返す。
* 条件付き書式を削除する。
* 優先順位機能と stopifTrue 機能を提供する。
* 指定した範囲のすべての条件付き書式のコレクションを取得する。
* 現在指定している範囲でアクティブなすべての条件付き書式をクリアする。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|次の "context.sync()" が呼び出されるまで、計算を中断します。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|この条件付き書式の Rule オブジェクトを表します。|
||[set (properties: CellValueConditionalFormat)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: CellValueConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[CellValueConditionalFormatData](/javascript/api/excel/excel.cellvalueconditionalformatdata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatdata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.cellvalueconditionalformatdata#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[CellValueConditionalFormatLoadOptions](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[CellValueConditionalFormatUpdateData](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|カラースケールの基準。 2ポイントのカラースケールを使用している場合、中点はオプションです。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|True の場合、カラースケールには3つのポイント (最小、中点、最大) が設定されます。それ以外の場合は、2つ (最小、最大) が設定されます。|
||[set (properties: ColorScaleConditionalFormat)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ColorScaleConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ColorScaleConditionalFormatData](/javascript/api/excel/excel.colorscaleconditionalformatdata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatdata#criteria)|カラースケールの基準。 2ポイントのカラースケールを使用している場合、中点はオプションです。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatdata#threecolorscale)|True の場合、カラースケールには3つのポイント (最小、中点、最大) が設定されます。それ以外の場合は、2つ (最小、最大) が設定されます。|
|[ColorScaleConditionalFormatLoadOptions](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#criteria)|カラースケールの基準。 2ポイントのカラースケールを使用している場合、中点はオプションです。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#threecolorscale)|True の場合、カラースケールには3つのポイント (最小、中点、最大) が設定されます。それ以外の場合は、2つ (最小、最大) が設定されます。|
|[ColorScaleConditionalFormatUpdateData](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata#criteria)|カラースケールの基準。 2ポイントのカラースケールを使用している場合、中点はオプションです。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|条件付き書式ルールを評価するために必要な場合、数式。|
||[演算子](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|テキスト条件付き書式の演算子を指定します。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|最大ポイントのカラー スケール条件。|
||[地点](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|カラー スケールが 3 色スケールの場合のカラー スケール条件の中間値。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|最小ポイントのカラー スケール条件。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色スケールの色を表す HTML カラーコード。 例: #FF0000 は赤を表します。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|数値、数式、(型が LowestValue の場合は) null。|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件式の基準となる条件式を指定します。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|
||[set (properties: ConditionalDataBarNegativeFormat)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalDataBarNegativeFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ConditionalDataBarNegativeFormatData](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivebordercolor)|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivefillcolor)|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|
|[ConditionalDataBarNegativeFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivebordercolor)|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivefillcolor)|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|
|[ConditionalDataBarNegativeFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivebordercolor)|負の DataBar に正の DataBar と同じ枠線の色があるかどうかを表すブール値。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivefillcolor)|負の DataBar に正の DataBar と同じ塗りつぶしの色があるかどうかを表すブール値。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|DataBar のグラデーションの有無を表すブール値。|
||[set (properties: ConditionalDataBarPositiveFormat)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalDataBarPositiveFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ConditionalDataBarPositiveFormatData](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#gradientfill)|DataBar のグラデーションの有無を表すブール値。|
|[ConditionalDataBarPositiveFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#gradientfill)|DataBar のグラデーションの有無を表すブール値。|
|[ConditionalDataBarPositiveFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#gradientfill)|DataBar のグラデーションの有無を表すブール値。|
|[Conditionalの配列](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|databar のルールを評価するために必要な場合、数式。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Databar のルールの種類。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|この条件付き書式を削除します。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|条件付き書式が適用された範囲を返す。 複数の範囲に条件付き書式を適用すると、エラーがスローされます。 読み取り専用です。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Conditonal 書式が適用される範囲を返します。または、複数の範囲に条件付き書式が適用されている場合は、null オブジェクトを返します。 読み取り専用です。|
||[的](/javascript/api/excel/excel.conditionalformat#priority)|この条件付き書式が現在存在している条件付き書式コレクション内の優先度 (またはインデックス)。 これも変更する|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。 読み取り専用です。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。 読み取り専用です。|
||[配色](/javascript/api/excel/excel.conditionalformat#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。 読み取り専用です。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。 読み取り専用です。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。 読み取り専用です。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。 読み取り専用です。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。 読み取り専用です。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|現在の ConditionalFormatCollection 内での条件付き書式の優先順位。 読み取り専用です。|
||[3-d](/javascript/api/excel/excel.conditionalformat#preset)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|条件付き書式の種類を指定します。 一度に設定できるのは1つだけです。 読み取り専用です。|
||[set (プロパティ: Excel. ConditionalFormat)](/javascript/api/excel/excel.conditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add (type: "Custom" \| "DataBar" \| "ColorScale" \| "IconSet" \| "topbottom" \| "PresetCriteria" \| "ContainsText" \| "cellvalue")](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|新しい条件付き書式をコレクションの先頭/最上位の優先度に追加します。|
||[追加 (種類: ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|新しい条件付き書式をコレクションの先頭/最上位の優先度に追加します。|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|ブック内の条件付き書式の数を返します。 読み取り専用です。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|指定された ID に対応する条件付き書式を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|指定されたインデックスに条件付き書式を返します。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ConditionalFormatCollectionLoadOptions](/javascript/api/excel/excel.conditionalformatcollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalue)|コレクション内の各アイテムについて: 現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalueornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscale)|コレクション内の各アイテムについて: 現在の条件付き書式が ColorScale 型の場合、ColorScale 条件付き書式プロパティを返します。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscaleornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式が ColorScale 型の場合、ColorScale 条件付き書式プロパティを返します。|
||[配色](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#custom)|コレクション内の各アイテムについて: 現在の条件付き書式がカスタム型の場合は、カスタム条件付き書式プロパティを返します。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#customornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式がカスタム型の場合は、カスタム条件付き書式プロパティを返します。|
||[dataBar](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databar)|コレクション内の各アイテムについて: 現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databarornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[iconSet](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconset)|コレクション内の各アイテムについて: 現在の条件付き書式が IconSet 型の場合、IconSet 条件付き書式プロパティを返します。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconsetornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式が IconSet 型の場合、IconSet 条件付き書式プロパティを返します。|
||[id](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#id)|コレクション内の各アイテムについて: 現在の ConditionalFormatCollection 内の条件付き書式の優先度。 読み取り専用です。|
||[3-d](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#preset)|コレクション内の各アイテムについて: 事前設定の条件条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#presetornullobject)|コレクション内の各アイテムについて: 事前設定の条件条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[的](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#priority)|コレクション内の各アイテムについて: この条件付き書式が現在存在している条件付き書式コレクション内の優先度 (またはインデックス)。 これも変更する|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#stopiftrue)|コレクション内の各アイテムについて: この条件付き書式の条件が満たされている場合、そのセルに対して優先度の低い書式は適用されません。|
||[textComparison](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparison)|コレクション内の各アイテムについて: 現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparisonornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottom)|コレクション内の各アイテムについて: 現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式のプロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottomornullobject)|コレクション内の各アイテムについて: 現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式のプロパティを返します。|
||[type](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#type)|コレクション内の各項目について、条件付き書式の種類。 一度に設定できるのは1つだけです。 読み取り専用です。|
|[ConditionalFormatData](/javascript/api/excel/excel.conditionalformatdata)|[cellValue](/javascript/api/excel/excel.conditionalformatdata#cellvalue)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatdata#cellvalueornullobject)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformatdata#colorscale)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。 読み取り専用です。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatdata#colorscaleornullobject)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。 読み取り専用です。|
||[配色](/javascript/api/excel/excel.conditionalformatdata#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。 読み取り専用です。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatdata#customornullobject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。 読み取り専用です。|
||[dataBar](/javascript/api/excel/excel.conditionalformatdata#databar)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。 読み取り専用です。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatdata#databarornullobject)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。 読み取り専用です。|
||[iconSet](/javascript/api/excel/excel.conditionalformatdata#iconset)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。 読み取り専用です。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#iconsetornullobject)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.conditionalformatdata#id)|現在の ConditionalFormatCollection 内での条件付き書式の優先順位。 読み取り専用です。|
||[3-d](/javascript/api/excel/excel.conditionalformatdata#preset)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#presetornullobject)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[的](/javascript/api/excel/excel.conditionalformatdata#priority)|この条件付き書式が現在存在している条件付き書式コレクション内の優先度 (またはインデックス)。 これも変更する|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatdata#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
||[textComparison](/javascript/api/excel/excel.conditionalformatdata#textcomparison)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatdata#textcomparisonornullobject)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformatdata#topbottom)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatdata#topbottomornullobject)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[type](/javascript/api/excel/excel.conditionalformatdata#type)|条件付き書式の種類を指定します。 一度に設定できるのは1つだけです。 読み取り専用です。|
|[ConditionalFormatLoadOptions](/javascript/api/excel/excel.conditionalformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalue)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalueornullobject)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformatloadoptions#colorscale)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#colorscaleornullobject)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[配色](/javascript/api/excel/excel.conditionalformatloadoptions#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#customornullobject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[dataBar](/javascript/api/excel/excel.conditionalformatloadoptions#databar)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#databarornullobject)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[iconSet](/javascript/api/excel/excel.conditionalformatloadoptions#iconset)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#iconsetornullobject)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[id](/javascript/api/excel/excel.conditionalformatloadoptions#id)|現在の ConditionalFormatCollection 内での条件付き書式の優先順位。 読み取り専用です。|
||[3-d](/javascript/api/excel/excel.conditionalformatloadoptions#preset)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#presetornullobject)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[的](/javascript/api/excel/excel.conditionalformatloadoptions#priority)|この条件付き書式が現在存在している条件付き書式コレクション内の優先度 (またはインデックス)。 これも変更する|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatloadoptions#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
||[textComparison](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparison)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparisonornullobject)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformatloadoptions#topbottom)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#topbottomornullobject)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[type](/javascript/api/excel/excel.conditionalformatloadoptions#type)|条件付き書式の種類を指定します。 一度に設定できるのは1つだけです。 読み取り専用です。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
||[set (properties: Excel. ConditionalFormatRule)](/javascript/api/excel/excel.conditionalformatrule#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalFormatRuleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalformatrule#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ConditionalFormatRuleData](/javascript/api/excel/excel.conditionalformatruledata)|[formula](/javascript/api/excel/excel.conditionalformatruledata#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruledata#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruledata#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
|[ConditionalFormatRuleLoadOptions](/javascript/api/excel/excel.conditionalformatruleloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatruleloadoptions#$all)||
||[formula](/javascript/api/excel/excel.conditionalformatruleloadoptions#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleloadoptions#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleloadoptions#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
|[ConditionalFormatRuleUpdateData](/javascript/api/excel/excel.conditionalformatruleupdatedata)|[formula](/javascript/api/excel/excel.conditionalformatruleupdatedata#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleupdatedata#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleupdatedata#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
|[ConditionalFormatUpdateData](/javascript/api/excel/excel.conditionalformatupdatedata)|[cellValue](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalue)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalueornullobject)|現在の条件付き書式が CellValue 型の場合は、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformatupdatedata#colorscale)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#colorscaleornullobject)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[配色](/javascript/api/excel/excel.conditionalformatupdatedata#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#customornullobject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[dataBar](/javascript/api/excel/excel.conditionalformatupdatedata#databar)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#databarornullobject)|現在の条件付き書式がデータバーの場合、データバーのプロパティを返します。|
||[iconSet](/javascript/api/excel/excel.conditionalformatupdatedata#iconset)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#iconsetornullobject)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[3-d](/javascript/api/excel/excel.conditionalformatupdatedata#preset)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#presetornullobject)|事前設定の条件の条件付き書式を返します。 詳細については、「PresetCriteriaConditionalFormat」を参照してください。|
||[的](/javascript/api/excel/excel.conditionalformatupdatedata#priority)|この条件付き書式が現在存在している条件付き書式コレクション内の優先度 (またはインデックス)。 これも変更する|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatupdatedata#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
||[textComparison](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparison)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparisonornullobject)|現在の条件付き書式がテキスト型の場合、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformatupdatedata#topbottom)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#topbottomornullobject)|現在の条件付き書式が TopBottom 型の場合、上位/下位条件付き書式プロパティを返します。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|既定の IconSet と異なる場合は現在の条件のカスタム アイコン、そうでない場合は null が返されます。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|種類によっては数値または数式。|
||[演算子](/javascript/api/excel/excel.conditionaliconcriterion#operator)|アイコンの条件付き書式のルールの種類ごとに、GreaterThan または GreaterThanOrEqual。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|アイコンの条件式は次のものに基づいています。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[条件](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件付き書式の条件を指定します。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel の ConditionalRangeBorderIndex」を参照してください。 読み取り専用です。|
||[set (properties: Excel. ConditionalRangeBorder)](/javascript/api/excel/excel.conditionalrangeborder#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalRangeBorderUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalrangeborder#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "edgeright")](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItem (index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|下罫線を取得します。 読み取り専用です。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。 読み取り専用です。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|左罫線を取得します。 読み取り専用です。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|右罫線を取得します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|上罫線を取得します。 読み取り専用です。|
|[ConditionalRangeBorderCollectionLoadOptions](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#color)|コレクション内の各項目について、フォーム #RRGGBB の境界線の色を表す HTML カラーコード ("FFA500" など) または名前付きの HTML 色 (例: "オレンジ")。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#sideindex)|コレクション内の各項目について、境界線の特定の辺を示す定数値。 詳細については、「Excel の ConditionalRangeBorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#style)|コレクション内の各アイテムについて: 境界線のスタイルを指定する線スタイルの定数のいずれかです。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeBorderCollectionUpdateData](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#bottom)|下罫線を取得します。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#left)|左罫線を取得します。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#right)|右罫線を取得します。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#top)|上罫線を取得します。|
|[ConditionalRangeBorderData](/javascript/api/excel/excel.conditionalrangeborderdata)|[color](/javascript/api/excel/excel.conditionalrangeborderdata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderdata#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel の ConditionalRangeBorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.conditionalrangeborderdata#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeBorderLoadOptions](/javascript/api/excel/excel.conditionalrangeborderloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangeborderloadoptions#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderloadoptions#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel の ConditionalRangeBorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.conditionalrangeborderloadoptions#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeBorderUpdateData](/javascript/api/excel/excel.conditionalrangeborderupdatedata)|[color](/javascript/api/excel/excel.conditionalrangeborderupdatedata#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[style](/javascript/api/excel/excel.conditionalrangeborderupdatedata#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|塗りつぶしをリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[set (properties: Excel. ConditionalRangeFill)](/javascript/api/excel/excel.conditionalrangefill#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalRangeFillUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalrangefill#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ConditionalRangeFillData](/javascript/api/excel/excel.conditionalrangefilldata)|[color](/javascript/api/excel/excel.conditionalrangefilldata#color)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
|[ConditionalRangeFillLoadOptions](/javascript/api/excel/excel.conditionalrangefillloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangefillloadoptions#color)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
|[ConditionalRangeFillUpdateData](/javascript/api/excel/excel.conditionalrangefillupdatedata)|[color](/javascript/api/excel/excel.conditionalrangefillupdatedata#color)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|フォントの太字の状態を表します。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|フォントの書式設定をリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|フォントの斜体の状態を表します。|
||[set (properties: Excel. ConditionalRangeFont)](/javascript/api/excel/excel.conditionalrangefont#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalRangeFontUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalrangefont#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|フォントの取り消し線の状態を表します。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Conditionalrangefont過小認識」を参照してください。|
|[ConditionalRangeFontData](/javascript/api/excel/excel.conditionalrangefontdata)|[bold](/javascript/api/excel/excel.conditionalrangefontdata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.conditionalrangefontdata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.conditionalrangefontdata#italic)|フォントの斜体の状態を表します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontdata#strikethrough)|フォントの取り消し線の状態を表します。|
||[underline](/javascript/api/excel/excel.conditionalrangefontdata#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Conditionalrangefont過小認識」を参照してください。|
|[ConditionalRangeFontLoadOptions](/javascript/api/excel/excel.conditionalrangefontloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.conditionalrangefontloadoptions#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.conditionalrangefontloadoptions#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.conditionalrangefontloadoptions#italic)|フォントの斜体の状態を表します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontloadoptions#strikethrough)|フォントの取り消し線の状態を表します。|
||[underline](/javascript/api/excel/excel.conditionalrangefontloadoptions#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Conditionalrangefont過小認識」を参照してください。|
|[ConditionalRangeFontUpdateData](/javascript/api/excel/excel.conditionalrangefontupdatedata)|[bold](/javascript/api/excel/excel.conditionalrangefontupdatedata#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/excel/excel.conditionalrangefontupdatedata#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.conditionalrangefontupdatedata#italic)|フォントの斜体の状態を表します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontupdatedata#strikethrough)|フォントの取り消し線の状態を表します。|
||[underline](/javascript/api/excel/excel.conditionalrangefontupdatedata#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Conditionalrangefont過小認識」を参照してください。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|指定された範囲の Excel の数値書式コードを表します。 Null が渡された場合はクリアされます。|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|条件付き書式の範囲全体に適用される border オブジェクトのコレクションです。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|条件付き書式の範囲全体で定義される fill オブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|条件付き書式の範囲全体で定義される font オブジェクトを返します。 読み取り専用です。|
||[set (properties: Excel. ConditionalRangeFormat)](/javascript/api/excel/excel.conditionalrangeformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ConditionalRangeFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.conditionalrangeformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ConditionalRangeFormatData](/javascript/api/excel/excel.conditionalrangeformatdata)|[borders](/javascript/api/excel/excel.conditionalrangeformatdata#borders)|条件付き書式の範囲全体に適用される border オブジェクトのコレクションです。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatdata#fill)|条件付き書式の範囲全体で定義される fill オブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.conditionalrangeformatdata#font)|条件付き書式の範囲全体で定義される font オブジェクトを返します。 読み取り専用です。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatdata#numberformat)|指定された範囲の Excel の数値書式コードを表します。 Null が渡された場合はクリアされます。|
|[ConditionalRangeFormatLoadOptions](/javascript/api/excel/excel.conditionalrangeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeformatloadoptions#$all)||
||[borders](/javascript/api/excel/excel.conditionalrangeformatloadoptions#borders)|条件付き書式の範囲全体に適用される border オブジェクトのコレクションです。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatloadoptions#fill)|条件付き書式の範囲全体で定義される fill オブジェクトを返します。|
||[font](/javascript/api/excel/excel.conditionalrangeformatloadoptions#font)|条件付き書式の範囲全体で定義される font オブジェクトを返します。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatloadoptions#numberformat)|指定された範囲の Excel の数値書式コードを表します。 Null が渡された場合はクリアされます。|
|[ConditionalRangeFormatUpdateData](/javascript/api/excel/excel.conditionalrangeformatupdatedata)|[borders](/javascript/api/excel/excel.conditionalrangeformatupdatedata#borders)|条件付き書式の範囲全体に適用される border オブジェクトのコレクションです。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatupdatedata#fill)|条件付き書式の範囲全体で定義される fill オブジェクトを返します。|
||[font](/javascript/api/excel/excel.conditionalrangeformatupdatedata#font)|条件付き書式の範囲全体で定義される font オブジェクトを返します。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatupdatedata#numberformat)|指定された範囲の Excel の数値書式コードを表します。 Null が渡された場合はクリアされます。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[演算子](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|テキスト条件付き書式の演算子を指定します。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件付き書式のテキスト値。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[Rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|上位または下位のランクに基づいて値を書式設定します。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.customconditionalformat#rule)|この条件付き書式の Rule オブジェクトを表します。 読み取り専用です。|
||[set (properties: Excel. CustomConditionalFormat)](/javascript/api/excel/excel.customconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: CustomConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.customconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[CustomConditionalFormatData](/javascript/api/excel/excel.customconditionalformatdata)|[format](/javascript/api/excel/excel.customconditionalformatdata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.customconditionalformatdata#rule)|この条件付き書式の Rule オブジェクトを表します。 読み取り専用です。|
|[CustomConditionalFormatLoadOptions](/javascript/api/excel/excel.customconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.customconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.customconditionalformatloadoptions#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.customconditionalformatloadoptions#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[CustomConditionalFormatUpdateData](/javascript/api/excel/excel.customconditionalformatupdatedata)|[format](/javascript/api/excel/excel.customconditionalformatupdatedata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.customconditionalformatupdatedata#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[[ファイル]](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[軸書式](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Excel データバーの軸をどのように判別するかを表します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|データバーのグラフィックスの基準となる方向を表します。|
||[小 Boundrule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Excel データバーの軸の左側にあるすべての値の表現。 読み取り専用です。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Excel データバーの軸の右側にあるすべての値の表現。 読み取り専用です。|
||[set (properties: エクセル Arconditionalformat)](/javascript/api/excel/excel.databarconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: DataBarConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.databarconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[Showます Aronly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[データ](/javascript/api/excel/excel.databarconditionalformatdata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatdata#axiscolor)|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[軸書式](/javascript/api/excel/excel.databarconditionalformatdata#axisformat)|Excel データバーの軸をどのように判別するかを表します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatdata#bardirection)|データバーのグラフィックスの基準となる方向を表します。|
||[小 Boundrule](/javascript/api/excel/excel.databarconditionalformatdata#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatdata#negativeformat)|Excel データバーの軸の左側にあるすべての値の表現。 読み取り専用です。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatdata#positiveformat)|Excel データバーの軸の右側にあるすべての値の表現。 読み取り専用です。|
||[Showます Aronly](/javascript/api/excel/excel.databarconditionalformatdata#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[' (A) '。](/javascript/api/excel/excel.databarconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.databarconditionalformatloadoptions#$all)||
||[axisColor](/javascript/api/excel/excel.databarconditionalformatloadoptions#axiscolor)|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[軸書式](/javascript/api/excel/excel.databarconditionalformatloadoptions#axisformat)|Excel データバーの軸をどのように判別するかを表します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatloadoptions#bardirection)|データバーのグラフィックスの基準となる方向を表します。|
||[小 Boundrule](/javascript/api/excel/excel.databarconditionalformatloadoptions#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#negativeformat)|Excel データバーの軸の左側にあるすべての値の表現。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#positiveformat)|Excel データバーの軸の右側にあるすべての値の表現。|
||[Showます Aronly](/javascript/api/excel/excel.databarconditionalformatloadoptions#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[DataBarConditionalFormatUpdateData](/javascript/api/excel/excel.databarconditionalformatupdatedata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatupdatedata#axiscolor)|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[軸書式](/javascript/api/excel/excel.databarconditionalformatupdatedata#axisformat)|Excel データバーの軸をどのように判別するかを表します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatupdatedata#bardirection)|データバーのグラフィックスの基準となる方向を表します。|
||[小 Boundrule](/javascript/api/excel/excel.databarconditionalformatupdatedata#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#negativeformat)|Excel データバーの軸の左側にあるすべての値の表現。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#positiveformat)|Excel データバーの軸の右側にあるすべての値の表現。|
||[Showます Aronly](/javascript/api/excel/excel.databarconditionalformatupdatedata#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|ルールの条件および IconSets の配列と、条件付きアイコンのユーザー設定のアイコン。 最初の条件では、カスタムアイコンのみを変更できることに注意してください。設定すると、type、formula、および operator は無視されます。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|True の場合は、IconSet のアイコンオーダーを逆にします。 カスタムアイコンが使用されている場合は、これを設定できないことに注意してください。|
||[set (properties: Excel. IconSetConditionalFormat)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: IconSetConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|設定すると、条件付き書式の IconSet オプションが表示されます。|
|[IconSetConditionalFormatData](/javascript/api/excel/excel.iconsetconditionalformatdata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatdata#criteria)|ルールの条件および IconSets の配列と、条件付きアイコンのユーザー設定のアイコン。 最初の条件では、カスタムアイコンのみを変更できることに注意してください。設定すると、type、formula、および operator は無視されます。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatdata#reverseiconorder)|True の場合は、IconSet のアイコンオーダーを逆にします。 カスタムアイコンが使用されている場合は、これを設定できないことに注意してください。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatdata#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatdata#style)|設定すると、条件付き書式の IconSet オプションが表示されます。|
|[IconSetConditionalFormatLoadOptions](/javascript/api/excel/excel.iconsetconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#criteria)|ルールの条件および IconSets の配列と、条件付きアイコンのユーザー設定のアイコン。 最初の条件では、カスタムアイコンのみを変更できることに注意してください。設定すると、type、formula、および operator は無視されます。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#reverseiconorder)|True の場合は、IconSet のアイコンオーダーを逆にします。 カスタムアイコンが使用されている場合は、これを設定できないことに注意してください。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#style)|設定すると、条件付き書式の IconSet オプションが表示されます。|
|[IconSetConditionalFormatUpdateData](/javascript/api/excel/excel.iconsetconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#criteria)|ルールの条件および IconSets の配列と、条件付きアイコンのユーザー設定のアイコン。 最初の条件では、カスタムアイコンのみを変更できることに注意してください。設定すると、type、formula、および operator は無視されます。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#reverseiconorder)|True の場合は、IconSet のアイコンオーダーを逆にします。 カスタムアイコンが使用されている場合は、これを設定できないことに注意してください。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#style)|設定すると、条件付き書式の IconSet オプションが表示されます。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件付き書式のルール。|
||[set (properties: PresetCriteriaConditionalFormat)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PresetCriteriaConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[PresetCriteriaConditionalFormatData](/javascript/api/excel/excel.presetcriteriaconditionalformatdata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#rule)|条件付き書式のルール。|
|[PresetCriteriaConditionalFormatLoadOptions](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#rule)|条件付き書式のルール。|
|[PresetCriteriaConditionalFormatUpdateData](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#rule)|条件付き書式のルール。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|ワークシート上のセルの範囲を計算します。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|範囲に交差する ConditionalFormats のコレクションです。 読み取り専用です。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[conditionalFormats](/javascript/api/excel/excel.rangedata#conditionalformats)|範囲に交差する ConditionalFormats のコレクションです。 読み取り専用です。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.textconditionalformat#rule)|条件付き書式のルール。|
||[set (properties: Excel. TextConditionalFormat)](/javascript/api/excel/excel.textconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TextConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.textconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[TextConditionalFormatData](/javascript/api/excel/excel.textconditionalformatdata)|[format](/javascript/api/excel/excel.textconditionalformatdata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.textconditionalformatdata#rule)|条件付き書式のルール。|
|[TextConditionalFormatLoadOptions](/javascript/api/excel/excel.textconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.textconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.textconditionalformatloadoptions#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.textconditionalformatloadoptions#rule)|条件付き書式のルール。|
|[TextConditionalFormatUpdateData](/javascript/api/excel/excel.textconditionalformatupdatedata)|[format](/javascript/api/excel/excel.textconditionalformatupdatedata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.textconditionalformatupdatedata#rule)|条件付き書式のルール。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.topbottomconditionalformat#rule)|上位/下位条件付き書式の条件を指定します。|
||[set (プロパティ: Excel. TopBottomConditionalFormat)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TopBottomConditionalFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[TopBottomConditionalFormatData](/javascript/api/excel/excel.topbottomconditionalformatdata)|[format](/javascript/api/excel/excel.topbottomconditionalformatdata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.topbottomconditionalformatdata#rule)|上位/下位条件付き書式の条件を指定します。|
|[TopBottomConditionalFormatLoadOptions](/javascript/api/excel/excel.topbottomconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#rule)|上位/下位条件付き書式の条件を指定します。|
|[TopBottomConditionalFormatUpdateData](/javascript/api/excel/excel.topbottomconditionalformatupdatedata)|[format](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#rule)|上位/下位条件付き書式の条件を指定します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate (markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|ワークシート上のすべてのセルを計算します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
