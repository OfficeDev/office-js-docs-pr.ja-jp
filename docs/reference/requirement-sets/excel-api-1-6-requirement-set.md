---
title: ExcelJavaScript API 要件セット 1.6
description: ExcelApi 1.6 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bc2eb8f182a329808a46f172868b818027f5e367
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350107"
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
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|次の "context.sync()" が呼び出されるまで、計算を中断します。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|この条件付き形式の Rule オブジェクトを指定します。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|カラー スケールの条件。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|true の場合、カラー スケールには 3 つのポイント (最小、中点、最大値) が含め、それ以外の場合は 2 つ (最小、最大値) になります。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|条件付き書式ルールを評価するために必要な場合、数式。|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|セル値の条件付き書式の演算子。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|最大ポイントのカラー スケール条件。|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|カラー スケールが 3 色スケールの場合のカラー スケール条件の中間値。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|最小ポイントのカラー スケール条件。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色スケールの色の HTML カラー コード表現 (赤を表#FF0000など)。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|数値、数式、(型が LowestValue の場合は) null。|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件式の基準の基になる条件。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|境界線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|フォーム #RRGGBB の塗りつぶしの色 ("FFA500" など) を表す HTML カラー コード、または名前付き HTML 色 ("オレンジ色" など) を表します。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|負の DataBar が正の DataBar と同じ罫線の色を持っている場合に指定します。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|負の DataBar が正の DataBar と同じ塗りつぶし色を持つ場合に指定します。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|境界線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|フォーム #RRGGBB の塗りつぶしの色 ("FFA500" など) を表す HTML カラー コード、または名前付き HTML 色 ("オレンジ色" など) を表します。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|DataBar にグラデーションが設定されている場合に指定します。|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|databar のルールを評価するために必要な場合、数式。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|データ バーのルールの種類。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|この条件付き書式を削除します。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|条件付き書式が適用された範囲を返す。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|条件付き書式が複数の範囲に適用されている場合は、conditonal 形式が適用される範囲、または null オブジェクトを返します。|
||[優先度](/javascript/api/excel/excel.conditionalformat#priority)|この条件付き書式が現在存在する条件付き書式コレクション内の優先度 (またはインデックス)。|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|現在の条件付き書式が CellValue 型の場合、セル値の条件付き書式プロパティを返します。|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|現在の条件付き書式が ColorScale 型の場合は、ColorScale 条件付き書式プロパティを返します。|
||[カスタム](/javascript/api/excel/excel.conditionalformat#custom)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|現在の条件付き書式がカスタム型の場合は、カスタムの条件付き書式プロパティを返します。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|現在の条件付き書式がデータ バーの場合は、データ バーのプロパティを返します。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|現在の条件付き書式がデータ バーの場合は、データ バーのプロパティを返します。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|現在の条件付き書式が IconSet 型の場合は、IconSet 条件付き書式プロパティを返します。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|現在の ConditionalFormatCollection 内での条件付き書式の優先順位。|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|事前設定された条件の条件付き書式を返します。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|事前設定された条件の条件付き書式を返します。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|現在の条件付き書式がテキスト型の場合は、特定のテキスト条件付き書式プロパティを返します。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|現在の条件付き書式がテキスト型の場合は、特定のテキスト条件付き書式プロパティを返します。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|現在の条件付き書式が TopBottom 型の場合は、Top/Bottom 条件付き書式プロパティを返します。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|現在の条件付き書式が TopBottom 型の場合は、Top/Bottom 条件付き書式プロパティを返します。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|条件付き書式の種類。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|最初/最優先で新しい条件付き書式をコレクションに追加します。|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|ブック内の条件付き書式の数を返します。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|指定された ID に対応する条件付き書式を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|指定されたインデックスに条件付き書式を返します。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|既定の IconSet と異なる場合は現在の条件のカスタム アイコン、そうでない場合は null が返されます。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|種類によっては数値または数式。|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|Icon 条件付き書式のルールの種類ごとに、GreaterThan または GreaterThanOrEqual を指定します。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|アイコンの条件式は次のものに基づいています。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[条件](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件付き書式の条件。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|境界線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|罫線の特定の辺を表す定数値。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|下の罫線を取得します。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|左側の罫線を取得します。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|右の罫線を取得します。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|上の罫線を取得します。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|塗りつぶしをリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|塗りつぶしの色、フォーム #RRGGBB ("FFA500" など) を表す HTML カラー コード、または名前付き HTML 色 ("オレンジ色" など) を表します。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|フォントが太字の場合に指定します。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|フォントの書式設定をリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|テキストの色の HTML カラー コード表現 (例:赤を#FF0000など)。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|フォントが italic の場合に指定します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|フォントの取り消し線の状態を指定します。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|フォントに適用される下線の種類。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|指定したExcelの数値書式コードを表します。|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|条件付き書式範囲全体に適用される罫線オブジェクトのコレクション。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|条件付き書式の範囲全体で定義されている fill オブジェクトを返します。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|条件付き書式の範囲全体で定義されているフォント オブジェクトを返します。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|テキストの条件付き書式の演算子。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件付き書式のテキスト値。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|上または下のランクに基づいて値を書式設定します。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.customconditionalformat#rule)|この条件付き形式の Rule オブジェクトを指定します。|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|軸線、フォーム #RRGGBB の色 ("FFA500" など) を表す HTML カラー コード、または名前付き HTML 色 ("オレンジ色" など) を表します。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|データ バーに対して軸がどのように決定Excel表現します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|データ バー グラフィックの基になる方向を指定します。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|データ バー内の軸の左側のすべての値Excel表現します。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|データ バー内の軸の右側のすべての値Excel表示します。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|ルールの Criteria と IconSets の配列と、条件付きアイコンの潜在的なカスタム アイコン。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|true の場合は、IconSet のアイコンの順序を反転します。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|設定されている場合は、条件付き書式の IconSet オプションを表示します。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件付き書式のルール。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|ワークシート上のセルの範囲を計算します。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|範囲と交差する ConditionalFormats のコレクション。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.textconditionalformat#rule)|条件付き書式のルール。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|条件付き書式のフォント、塗りつぶし、罫線、その他のプロパティをカプセル化する format オブジェクトを返します。|
||[ルール](/javascript/api/excel/excel.topbottomconditionalformat#rule)|上/下の条件付き書式の条件。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|ワークシート上のすべてのセルを計算します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
