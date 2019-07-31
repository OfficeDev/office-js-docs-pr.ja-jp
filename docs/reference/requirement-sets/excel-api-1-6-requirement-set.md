---
title: Excel JavaScript API 要件セット1.6
description: ExcelApi 1.6 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1c81a2241d7e82e7fd2146189638c43ec1865296
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940767"
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
* 優先度と`stopifTrue`機能を提供します。
* 指定した範囲のすべての条件付き書式のコレクションを取得する。
* 現在指定している範囲でアクティブなすべての条件付き書式をクリアする。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|次の "context.sync()" が呼び出されるまで、計算を中断します。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|この条件付き書式の Rule オブジェクトを表します。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|カラースケールの基準。 2ポイントのカラースケールを使用している場合、中点はオプションです。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|True の場合、カラースケールには3つのポイント (最小、中点、最大) が設定されます。それ以外の場合は、2つ (最小、最大) が設定されます。|
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
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|DataBar のグラデーションの有無を表すブール値。|
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
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|この条件付き書式の条件が満たされた場合、優先順位の低い書式はそのセルに影響を及ぼしません。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[追加 (種類: ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|新しい条件付き書式をコレクションの先頭/最上位の優先度に追加します。|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|現在指定している範囲でアクティブなすべての条件付き書式をクリアする。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|ブック内の条件付き書式の数を返します。 読み取り専用です。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|指定された ID に対応する条件付き書式を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|指定されたインデックスに条件付き書式を返します。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|ユーザーの言語で条件付き書式ルールを評価するために必要な場合、数式。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|R1C1 形式の表記法で条件付き書式ルールを評価するために必要な場合、数式。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|既定の IconSet と異なる場合は現在の条件のカスタム アイコン、そうでない場合は null が返されます。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|種類によっては数値または数式。|
||[演算子](/javascript/api/excel/excel.conditionaliconcriterion#operator)|アイコンの条件付き書式のルールの種類ごとに、GreaterThan または GreaterThanOrEqual。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|アイコンの条件式は次のものに基づいています。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[条件](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件付き書式の条件を指定します。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|枠線の色を表す HTML カラー コード。形式は #RRGGBB (例: "FFA500")、または名前付きの HTML 色 (例: "オレンジ") です。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|罫線の特定の辺を表す定数値。 詳細については、「Excel の ConditionalRangeBorderIndex」を参照してください。 読み取り専用です。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|罫線の線スタイルを指定する、線スタイル定数のいずれか 1 つ。 詳細については、「Excel BorderLineStyle」を参照してください。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|オブジェクトの名前を使用して、境界線オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|オブジェクトのインデックスを使用して、境界線オブジェクトを取得します。|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|下罫線を取得します。 読み取り専用です。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|コレクションに含まれる境界線オブジェクトの数。 読み取り専用です。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|左罫線を取得します。 読み取り専用です。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|右罫線を取得します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|上罫線を取得します。 読み取り専用です。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|塗りつぶしをリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|塗りつぶしの色を表す HTML カラー コード。#RRGGBB 形式 (例: "FFA500")、または名前付きの HTML 色 (例: "orange") として示されます。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|フォントの太字の状態を表します。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|フォントの書式設定をリセットします。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|テキストの色の HTML カラー コード表記。 例: #FF0000 は赤を表します。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|フォントの斜体の状態を表します。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|フォントの取り消し線の状態を表します。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|フォントに適用する下線の種類。 詳細については、「Excel の Conditionalrangefont過小認識」を参照してください。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|指定された範囲の Excel の数値書式コードを表します。 Null が渡された場合はクリアされます。|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|条件付き書式の範囲全体に適用される border オブジェクトのコレクションです。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|条件付き書式の範囲全体で定義される fill オブジェクトを返します。 読み取り専用です。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|条件付き書式の範囲全体で定義される font オブジェクトを返します。 読み取り専用です。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[演算子](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|テキスト条件付き書式の演算子を指定します。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件付き書式のテキスト値。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[Rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|数値のランクに対する 1 から 1000、またはパーセントのランクに対する 1 から 100 のランク。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|上位または下位のランクに基づいて値を書式設定します。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.customconditionalformat#rule)|この条件付き書式の Rule オブジェクトを表します。 読み取り専用です。|
|[[ファイル]](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|軸の線の色を表す HTML カラー コード。形式は #RRGGBB (例:"FFA500")、または名前付きの HTML 色 (例: 「オレンジ」) です。|
||[軸書式](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Excel データバーの軸をどのように判別するかを表します。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|データバーのグラフィックスの基準となる方向を表します。|
||[小 Boundrule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|データ バーの下限値 (および該当する場合はその計算方法) を構成するルール。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Excel データバーの軸の左側にあるすべての値の表現。 読み取り専用です。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Excel データバーの軸の右側にあるすべての値の表現。 読み取り専用です。|
||[Showます Aronly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|true の場合、データ バーが適用されているセルの値を非表示にします。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|データ バーの上限値 (および該当する場合はその計算方法) を構成するルール。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|ルールの条件および IconSets の配列と、条件付きアイコンのユーザー設定のアイコン。 最初の条件では、カスタムアイコンのみを変更できることに注意してください。設定すると、type、formula、および operator は無視されます。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|True の場合は、IconSet のアイコンオーダーを逆にします。 カスタムアイコンが使用されている場合は、これを設定できないことに注意してください。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|true の場合、値は非表示にされて、アイコンのみが表示されます。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|設定すると、条件付き書式の IconSet オプションが表示されます。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。|
||[除外](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件付き書式のルール。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|ワークシート上のセルの範囲を計算します。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|範囲に交差する ConditionalFormats のコレクションです。 読み取り専用です。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.textconditionalformat#rule)|条件付き書式のルール。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|書式設定オブジェクトを返し、条件付き書式のフォント、塗りつぶし、罫線などのプロパティをカプセル化します。 読み取り専用です。|
||[除外](/javascript/api/excel/excel.topbottomconditionalformat#rule)|上位/下位条件付き書式の条件を指定します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate (markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|ワークシート上のすべてのセルを計算します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
