---
title: Excel JavaScript API 要件セット1.9
description: ExcelApi 1.9 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1c7361debe7ba09c3477d39d9337c35bf5df3066
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772003"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Excel JavaScript API 1.9 の新機能

1.9 の要件セットにより、500 件を超える新しい Excel API が 導入されました。 最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | 画像、幾何学な図形、テキスト ボックスを挿入、位置変更、書式設定します。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape) [Image](/javascript/api/excel/excel.image) |
| [オート フィルター](../../excel/excel-add-ins-worksheets.md#filter-data) | 範囲にフィルターを追加します。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [エリア](../../excel/excel-add-ins-multiple-ranges.md) | 連続していない範囲をサポートします。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特別なセル](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | ある範囲内に日付、コメント、数式を含むセルを取得します。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [検索](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | ある範囲またはワークシート内で値や数式を見つけます。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [コピーと貼り付け](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | 範囲間で値、書式、数式をコピーします。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [計算](../../excel/performance.md#suspend-calculation-temporarily) | Excel 計算エンジンを細かく操作できます。 | [アプリケーション](/javascript/api/excel/excel.application) |
| 新しいグラフ | 新しくサポートされたグラフである、マップ、箱ひげ図、ウォーターフォール、サンバースト、パレート、 じょうごをお試しください。 | [Chart](/javascript/api/excel/excel.charttype) |
| 範囲の形式 | 範囲の形式の新しい機能です。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。 読み取り専用です。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|アプリケーションの計算の状態を返します。 詳細については、Excel.CalculationState をご覧ください。 読み取り専用です。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|反復計算の設定を返します。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|次の "context.sync()" が呼び出されるまで画面の更新を一時停止します。|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationEngineVersion](/javascript/api/excel/excel.applicationdata#calculationengineversion)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。 読み取り専用です。|
||[calculationState](/javascript/api/excel/excel.applicationdata#calculationstate)|アプリケーションの計算の状態を返します。 詳細については、Excel.CalculationState をご覧ください。 読み取り専用です。|
||[iterativeCalculation](/javascript/api/excel/excel.applicationdata#iterativecalculation)|反復計算の設定を返します。|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[calculationEngineVersion](/javascript/api/excel/excel.applicationloadoptions#calculationengineversion)|最後の完全な再計算に使用した Excel 計算エンジンのバージョンを返します。 読み取り専用です。|
||[calculationState](/javascript/api/excel/excel.applicationloadoptions#calculationstate)|アプリケーションの計算の状態を返します。 詳細については、Excel.CalculationState をご覧ください。 読み取り専用です。|
||[iterativeCalculation](/javascript/api/excel/excel.applicationloadoptions#iterativecalculation)|反復計算の設定を返します。|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[iterativeCalculation](/javascript/api/excel/excel.applicationupdatedata#iterativecalculation)|反復計算の設定を返します。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|範囲にオートフィルターを適用します。 列インデックスやフィルター条件が指定されている場合、列にフィルターを適用します。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|オートフィルターのフィルター条件がクリアされます。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|AutoFilter が適用される範囲を表す Range オブジェクトを返します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|オートフィルターが適用される範囲を表す Range オブジェクトを返します。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。 読み取り専用です。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|AutoFilter が有効かどうかを示します。 読み取り専用です。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|AutoFilter にフィルター条件が与えられているかどうかを示します。 読み取り専用です。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|その範囲で現在指定されている Autofilter オブジェクトを適用します。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|範囲の AutoFilter を削除します。|
|[AutoFilterData](/javascript/api/excel/excel.autofilterdata)|[criteria](/javascript/api/excel/excel.autofilterdata#criteria)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。 読み取り専用です。|
||[enabled](/javascript/api/excel/excel.autofilterdata#enabled)|AutoFilter が有効かどうかを示します。 読み取り専用です。|
||[isDataFiltered](/javascript/api/excel/excel.autofilterdata#isdatafiltered)|AutoFilter にフィルター条件が与えられているかどうかを示します。 読み取り専用です。|
|[AutoFilterLoadOptions](/javascript/api/excel/excel.autofilterloadoptions)|[$all](/javascript/api/excel/excel.autofilterloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.autofilterloadoptions#criteria)|オートフィルターが適用された範囲のすべてのフィルター条件を保持する配列です。 読み取り専用です。|
||[enabled](/javascript/api/excel/excel.autofilterloadoptions#enabled)|AutoFilter が有効かどうかを示します。 読み取り専用です。|
||[isDataFiltered](/javascript/api/excel/excel.autofilterloadoptions#isdatafiltered)|AutoFilter にフィルター条件が与えられているかどうかを示します。 読み取り専用です。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|1 つの境界線の `color` プロパティを表します。|
||[style](/javascript/api/excel/excel.cellborder#style)|1 つの境界線の `style` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|1 つの境界線の `tintAndShade` プロパティを表します。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|1 つの境界線の `weight` プロパティを表します。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|`format.borders.bottom` プロパティを表します。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|`format.borders.diagonalDown` プロパティを表します。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|`format.borders.diagonalUp` プロパティを表します。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|`format.borders.horizontal` プロパティを表します。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|`format.borders.left` プロパティを表します。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|`format.borders.right` プロパティを表します。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|`format.borders.top` プロパティを表します。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|`format.borders.vertical` プロパティを表します。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|`hidden` プロパティを表します。|
|[CellPropertiesBorderLoadOptions](/javascript/api/excel/excel.cellpropertiesborderloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesborderloadoptions#color)|`color`プロパティに読み込むかどうかを指定します。|
||[style](/javascript/api/excel/excel.cellpropertiesborderloadoptions#style)|`style`プロパティに読み込むかどうかを指定します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesborderloadoptions#tintandshade)|`tintAndShade`プロパティに読み込むかどうかを指定します。|
||[weight](/javascript/api/excel/excel.cellpropertiesborderloadoptions#weight)|`weight`プロパティに読み込むかどうかを指定します。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|`format.fill.color` プロパティを表します。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|`format.fill.pattern` プロパティを表します。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|`format.fill.patternColor` プロパティを表します。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|`format.fill.patternTintAndShade` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|`format.fill.tintAndShade` プロパティを表します。|
|[CellPropertiesFillLoadOptions](/javascript/api/excel/excel.cellpropertiesfillloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesfillloadoptions#color)|`color`プロパティに読み込むかどうかを指定します。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfillloadoptions#pattern)|`pattern`プロパティに読み込むかどうかを指定します。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterncolor)|`patternColor`プロパティに読み込むかどうかを指定します。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterntintandshade)|`patternTintAndShade`プロパティに読み込むかどうかを指定します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#tintandshade)|`tintAndShade`プロパティに読み込むかどうかを指定します。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|`format.font.bold` プロパティを表します。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|`format.font.color` プロパティを表します。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|`format.font.italic` プロパティを表します。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|`format.font.name` プロパティを表します。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|`format.font.size` プロパティを表します。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|`format.font.strikethrough` プロパティを表します。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|`format.font.subscript` プロパティを表します。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|`format.font.superscript` プロパティを表します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|`format.font.tintAndShade` プロパティを表します。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|`format.font.underline` プロパティを表します。|
|[CellPropertiesFontLoadOptions](/javascript/api/excel/excel.cellpropertiesfontloadoptions)|[bold](/javascript/api/excel/excel.cellpropertiesfontloadoptions#bold)|`bold`プロパティに読み込むかどうかを指定します。|
||[color](/javascript/api/excel/excel.cellpropertiesfontloadoptions#color)|`color`プロパティに読み込むかどうかを指定します。|
||[italic](/javascript/api/excel/excel.cellpropertiesfontloadoptions#italic)|`italic`プロパティに読み込むかどうかを指定します。|
||[name](/javascript/api/excel/excel.cellpropertiesfontloadoptions#name)|`name`プロパティに読み込むかどうかを指定します。|
||[size](/javascript/api/excel/excel.cellpropertiesfontloadoptions#size)|`size`プロパティに読み込むかどうかを指定します。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfontloadoptions#strikethrough)|`strikethrough`プロパティに読み込むかどうかを指定します。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#subscript)|`subscript`プロパティに読み込むかどうかを指定します。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#superscript)|`superscript`プロパティに読み込むかどうかを指定します。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfontloadoptions#tintandshade)|`tintAndShade`プロパティに読み込むかどうかを指定します。|
||[underline](/javascript/api/excel/excel.cellpropertiesfontloadoptions#underline)|`underline`プロパティに読み込むかどうかを指定します。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|`autoIndent` プロパティを表します。|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|`borders` プロパティを表します。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|`fill` プロパティを表します。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|`font` プロパティを表します。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|`horizontalAlignment` プロパティを表します。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|`indentLevel` プロパティを表します。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|`protection` プロパティを表します。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|`readingOrder` プロパティを表します。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|`shrinkToFit` プロパティを表します。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|`textOrientation` プロパティを表します。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|`useStandardHeight` プロパティを表します。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|`useStandardWidth` プロパティを表します。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|`verticalAlignment` プロパティを表します。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|`wrapText` プロパティを表します。|
|[CellPropertiesFormatLoadOptions](/javascript/api/excel/excel.cellpropertiesformatloadoptions)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformatloadoptions#autoindent)|`autoIndent`プロパティに読み込むかどうかを指定します。|
||[borders](/javascript/api/excel/excel.cellpropertiesformatloadoptions#borders)|`borders`プロパティに読み込むかどうかを指定します。|
||[fill](/javascript/api/excel/excel.cellpropertiesformatloadoptions#fill)|`fill`プロパティに読み込むかどうかを指定します。|
||[font](/javascript/api/excel/excel.cellpropertiesformatloadoptions#font)|`font`プロパティに読み込むかどうかを指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#horizontalalignment)|`horizontalAlignment`プロパティに読み込むかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformatloadoptions#indentlevel)|`indentLevel`プロパティに読み込むかどうかを指定します。|
||[protection](/javascript/api/excel/excel.cellpropertiesformatloadoptions#protection)|`protection`プロパティに読み込むかどうかを指定します。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformatloadoptions#readingorder)|`readingOrder`プロパティに読み込むかどうかを指定します。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformatloadoptions#shrinktofit)|`shrinkToFit`プロパティに読み込むかどうかを指定します。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformatloadoptions#textorientation)|`textOrientation`プロパティに読み込むかどうかを指定します。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardheight)|`useStandardHeight`プロパティに読み込むかどうかを指定します。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardwidth)|`useStandardWidth`プロパティに読み込むかどうかを指定します。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#verticalalignment)|`verticalAlignment`プロパティに読み込むかどうかを指定します。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformatloadoptions#wraptext)|`wrapText`プロパティに読み込むかどうかを指定します。|
|[CellPropertiesLoadOptions](/javascript/api/excel/excel.cellpropertiesloadoptions)|[address](/javascript/api/excel/excel.cellpropertiesloadoptions#address)|`address`プロパティに読み込むかどうかを指定します。|
||[addressLocal](/javascript/api/excel/excel.cellpropertiesloadoptions#addresslocal)|`addressLocal`プロパティに読み込むかどうかを指定します。|
||[format](/javascript/api/excel/excel.cellpropertiesloadoptions#format)|`format`プロパティに読み込むかどうかを指定します。|
||[hidden](/javascript/api/excel/excel.cellpropertiesloadoptions#hidden)|`hidden`プロパティに読み込むかどうかを指定します。|
||[hyperlink](/javascript/api/excel/excel.cellpropertiesloadoptions#hyperlink)|`hyperlink`プロパティに読み込むかどうかを指定します。|
||[style](/javascript/api/excel/excel.cellpropertiesloadoptions#style)|`style`プロパティに読み込むかどうかを指定します。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|`format.protection.formulaHidden` プロパティを表します。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|`format.protection.locked` プロパティを表します。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|変更後の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|変更前の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|変更後の値の型を表します。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|変更前の値の型を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Excel UI でグラフをアクティブにします。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|ピボット グラフのオプションをカプセル化します。 読み取り専用です。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|グラフの配色を返すか、設定します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|グラフのグラフ エリアの角が丸いかどうかを指定します。 読み取り/書き込み可能。|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[colorScheme](/javascript/api/excel/excel.chartareaformatdata#colorscheme)|グラフの配色を返すか、設定します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatdata#roundedcorners)|グラフのグラフ エリアの角が丸いかどうかを指定します。 読み取り/書き込み可能。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[colorScheme](/javascript/api/excel/excel.chartareaformatloadoptions#colorscheme)|グラフの配色を返すか、設定します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatloadoptions#roundedcorners)|グラフのグラフ エリアの角が丸いかどうかを指定します。 読み取り/書き込み可能。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[colorScheme](/javascript/api/excel/excel.chartareaformatupdatedata#colorscheme)|グラフの配色を返すか、設定します。 読み取り/書き込み可能。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatupdatedata#roundedcorners)|グラフのグラフ エリアの角が丸いかどうかを指定します。 読み取り/書き込み可能。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisdata#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisloadoptions#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisupdatedata#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[set (properties: Excel. ChartBinOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartBinOptionsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartbinoptions#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBinOptionsData](/javascript/api/excel/excel.chartbinoptionsdata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsdata#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsdata#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptionsdata#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsdata#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[type](/javascript/api/excel/excel.chartbinoptionsdata#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsdata#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptionsdata#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBinOptionsLoadOptions](/javascript/api/excel/excel.chartbinoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartbinoptionsloadoptions#$all)||
||[allowOverflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptionsloadoptions#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[type](/javascript/api/excel/excel.chartbinoptionsloadoptions#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptionsloadoptions#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBinOptionsUpdateData](/javascript/api/excel/excel.chartbinoptionsupdatedata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowoverflow)|ヒストグラム図やパレート図でビンのオーバーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowunderflow)|ヒストグラム図やパレート図でビンのアンダーフローが有効になっているかどうかを指定します。 読み取り/書き込み可能。|
||[count](/javascript/api/excel/excel.chartbinoptionsupdatedata#count)|ヒストグラム図やパレート図のビン数を設定または返します。 読み取り/書き込み可能。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#overflowvalue)|ヒストグラム図やパレート図のビンのオーバーフロー値を設定または返します。 読み取り/書き込み可能。|
||[type](/javascript/api/excel/excel.chartbinoptionsupdatedata#type)|ヒストグラム図やパレート図のビンの種類を設定または返します。 読み取り/書き込み可能。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#underflowvalue)|ヒストグラム図やパレート図のビンのアンダーフロー値を設定または返します。 読み取り/書き込み可能。|
||[width](/javascript/api/excel/excel.chartbinoptionsupdatedata#width)|ヒストグラム図やパレート図のビンの幅値を設定または返します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[set (properties: ChartBoxwhiskerOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartBoxwhiskerOptionsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|箱ひげ図で、内部ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|箱ひげ図で、平均線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|箱ひげ図で、平均マーカーが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|箱ひげ図で、特異ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptionsData](/javascript/api/excel/excel.chartboxwhiskeroptionsdata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showinnerpoints)|箱ひげ図で、内部ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanline)|箱ひげ図で、平均線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanmarker)|箱ひげ図で、平均マーカーが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showoutlierpoints)|箱ひげ図で、特異ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptionsLoadOptions](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions)|[$all](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#$all)||
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showinnerpoints)|箱ひげ図で、内部ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanline)|箱ひげ図で、平均線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanmarker)|箱ひげ図で、平均マーカーが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showoutlierpoints)|箱ひげ図で、特異ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
|[ChartBoxwhiskerOptionsUpdateData](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#quartilecalculation)|箱ひげ図の四分位数計算の種類を設定または返します。 読み取り/書き込み可能。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showinnerpoints)|箱ひげ図で、内部ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanline)|箱ひげ図で、平均線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanmarker)|箱ひげ図で、平均マーカーが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showoutlierpoints)|箱ひげ図で、特異ポイントが表示されるかどうかを指定します。 読み取り/書き込み可能。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartcollectionloadoptions#pivotoptions)|コレクション内の各アイテムについて: ピボットグラフのオプションをカプセル化します。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[pivotOptions](/javascript/api/excel/excel.chartdata#pivotoptions)|ピボット グラフのオプションをカプセル化します。 読み取り専用です。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabeldata#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsdata#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#linknumberformat)|数値形式がセルにリンクされているかどうかを表します。 true の場合、セルで数値形式が変更されるとラベルでも数値形式が変更されます。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|誤差範囲に終点のスタイル線端があるかどうかを指定します。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|誤差範囲のどの部分を含めるかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|誤差範囲の書式の種類を指定します。|
||[set (properties: ChartErrorBars)](/javascript/api/excel/excel.charterrorbars#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartErrorBarsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charterrorbars#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|誤差範囲が表示されるかどうかを指定します。|
|[ChartErrorBarsData](/javascript/api/excel/excel.charterrorbarsdata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsdata#endstylecap)|誤差範囲に終点のスタイル線端があるかどうかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbarsdata#format)|誤差範囲の書式の種類を指定します。|
||[include](/javascript/api/excel/excel.charterrorbarsdata#include)|誤差範囲のどの部分を含めるかを指定します。|
||[type](/javascript/api/excel/excel.charterrorbarsdata#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbarsdata#visible)|誤差範囲が表示されるかどうかを指定します。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|グラフの線の書式設定を表します。|
||[set (properties: ChartErrorBarsFormat)](/javascript/api/excel/excel.charterrorbarsformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartErrorBarsFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.charterrorbarsformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartErrorBarsFormatData](/javascript/api/excel/excel.charterrorbarsformatdata)|[line](/javascript/api/excel/excel.charterrorbarsformatdata#line)|グラフの線の書式設定を表します。|
|[ChartErrorBarsFormatLoadOptions](/javascript/api/excel/excel.charterrorbarsformatloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charterrorbarsformatloadoptions#line)|グラフの線の書式設定を表します。|
|[ChartErrorBarsFormatUpdateData](/javascript/api/excel/excel.charterrorbarsformatupdatedata)|[line](/javascript/api/excel/excel.charterrorbarsformatupdatedata#line)|グラフの線の書式設定を表します。|
|[ChartErrorBarsLoadOptions](/javascript/api/excel/excel.charterrorbarsloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsloadoptions#$all)||
||[endStyleCap](/javascript/api/excel/excel.charterrorbarsloadoptions#endstylecap)|誤差範囲に終点のスタイル線端があるかどうかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbarsloadoptions#format)|誤差範囲の書式の種類を指定します。|
||[include](/javascript/api/excel/excel.charterrorbarsloadoptions#include)|誤差範囲のどの部分を含めるかを指定します。|
||[type](/javascript/api/excel/excel.charterrorbarsloadoptions#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbarsloadoptions#visible)|誤差範囲が表示されるかどうかを指定します。|
|[ChartErrorBarsUpdateData](/javascript/api/excel/excel.charterrorbarsupdatedata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsupdatedata#endstylecap)|誤差範囲に終点のスタイル線端があるかどうかを指定します。|
||[format](/javascript/api/excel/excel.charterrorbarsupdatedata#format)|誤差範囲の書式の種類を指定します。|
||[include](/javascript/api/excel/excel.charterrorbarsupdatedata#include)|誤差範囲のどの部分を含めるかを指定します。|
||[type](/javascript/api/excel/excel.charterrorbarsupdatedata#type)|誤差範囲でマークされている範囲の種類。|
||[visible](/javascript/api/excel/excel.charterrorbarsupdatedata#visible)|誤差範囲が表示されるかどうかを指定します。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartloadoptions#pivotoptions)|ピボット グラフのオプションをカプセル化します。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|リージョン マップ グラフの系列マッピング レベルを設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
||[set (プロパティ: Excel. ChartMapOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartMapOptionsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartmapoptions#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[ChartMapOptionsData](/javascript/api/excel/excel.chartmapoptionsdata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsdata#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptionsdata#level)|リージョン マップ グラフの系列マッピング レベルを設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsdata#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
|[ChartMapOptionsLoadOptions](/javascript/api/excel/excel.chartmapoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartmapoptionsloadoptions#$all)||
||[labelStrategy](/javascript/api/excel/excel.chartmapoptionsloadoptions#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptionsloadoptions#level)|リージョン マップ グラフの系列マッピング レベルを設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsloadoptions#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
|[ChartMapOptionsUpdateData](/javascript/api/excel/excel.chartmapoptionsupdatedata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsupdatedata#labelstrategy)|リージョン マップ グラフの系列マップ ラベル方法を設定または返します。 読み取り/書き込み可能。|
||[level](/javascript/api/excel/excel.chartmapoptionsupdatedata#level)|リージョン マップ グラフの系列マッピング レベルを設定または返します。 読み取り/書き込み可能。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsupdatedata#projectiontype)|リージョン マップ グラフの系列投影タイプを設定または返します。 読み取り/書き込み可能。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[set (properties: ChartPivotOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ChartPivotOptionsUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.chartpivotoptions#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを指定します。 ShowAxisFieldButtons プロパティは、ピボットグラフを選択したときに有効になる [分析] タブの [フィールド ボタン] ボックスの一覧にある [軸フィールド ボタンを表示する] コマンドに対応します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを指定します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを指定します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|ピボットグラフに値フィールド ボタンを表示するかどうかを指定します。|
|[ChartPivotOptionsData](/javascript/api/excel/excel.chartpivotoptionsdata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを指定します。 ShowAxisFieldButtons プロパティは、ピボットグラフを選択したときに有効になる [分析] タブの [フィールド ボタン] ボックスの一覧にある [軸フィールド ボタンを表示する] コマンドに対応します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを指定します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを指定します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showvaluefieldbuttons)|ピボットグラフに値フィールド ボタンを表示するかどうかを指定します。|
|[ChartPivotOptionsLoadOptions](/javascript/api/excel/excel.chartpivotoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartpivotoptionsloadoptions#$all)||
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを指定します。 ShowAxisFieldButtons プロパティは、ピボットグラフを選択したときに有効になる [分析] タブの [フィールド ボタン] ボックスの一覧にある [軸フィールド ボタンを表示する] コマンドに対応します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを指定します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを指定します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showvaluefieldbuttons)|ピボットグラフに値フィールド ボタンを表示するかどうかを指定します。|
|[ChartPivotOptionsUpdateData](/javascript/api/excel/excel.chartpivotoptionsupdatedata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showaxisfieldbuttons)|ピボットグラフに軸フィールド ボタンを表示するかどうかを指定します。 ShowAxisFieldButtons プロパティは、ピボットグラフを選択したときに有効になる [分析] タブの [フィールド ボタン] ボックスの一覧にある [軸フィールド ボタンを表示する] コマンドに対応します。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showlegendfieldbuttons)|ピボットグラフに凡例フィールド ボタンを表示するかどうかを指定します。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showreportfilterfieldbuttons)|ピボットグラフにレポート フィルター フィールド ボタンを表示するかどうかを指定します。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showvaluefieldbuttons)|ピボットグラフに値フィールド ボタンを表示するかどうかを指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。 読み取り専用です。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。 読み取り専用です。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|リージョン マップ グラフのオプションをカプセル化します。 読み取り専用です。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|ウォーターフォール図で、コネクタが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|系列の各データ ラベルの引き出し線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
|[Chartcharts Collectionloadoptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#binoptions)|コレクション内の各アイテムについて: ヒストグラムグラフおよびパレート図の bin オプションをカプセル化します。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#boxwhiskeroptions)|コレクション内の各アイテムについて: 箱ひげ図のオプションがカプセル化されています。|
||[bubbleScale](/javascript/api/excel/excel.chartseriescollectionloadoptions#bubblescale)|コレクション内の各項目に対して、既定のサイズの割合を表す 0 (ゼロ) から300の整数値を指定できます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumcolor)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最大値の色を設定します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumtype)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最大値の種類を設定します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumvalue)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最大値を設定します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointcolor)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の中間値の色を設定します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointtype)|コレクション内の各項目について、次の操作を行います。地域マップのグラフ系列の中間値の種類を設定します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointvalue)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の中間値を設定します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumcolor)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最小値の色を設定します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumtype)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最小値の型を設定します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumvalue)|コレクション内の各項目について、次の操作を行います。領域マップグラフの系列の最小値を設定します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientstyle)|コレクション内の各アイテムについて: 領域マップグラフの系列のグラデーションスタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertcolor)|コレクション内の各項目について、次のデータ系列内の負のデータ要素の塗りつぶしの色を設定します。 読み取り/書き込み可能。|
||[mapOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#mapoptions)|コレクション内の各項目について、次の操作を行います。地域マップグラフのオプションをカプセル化します。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriescollectionloadoptions#parentlabelstrategy)|コレクション内の各アイテムについて: ツリーマップグラフの系列の親ラベル戦略領域を設定または返します。 読み取り/書き込み可能。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showconnectorlines)|コレクション内の各アイテムについて: フォールドグラフでコネクタの線を表示するかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showleaderlines)|コレクション内の各アイテムについて: 系列の各データラベルに対して、引き出し線を表示するかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#splitvalue)|コレクション内の各項目について: 円グラフまたは補助縦棒グラフ付き円グラフの2つのセクションを区切るしきい値を設定します。 読み取り/書き込み可能。|
||[xErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#xerrorbars)|コレクション内の各アイテムについて: グラフ系列の誤差 bar オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#yerrorbars)|コレクション内の各アイテムについて: グラフ系列の誤差 bar オブジェクトを表します。|
|[Chart系列データ](/javascript/api/excel/excel.chartseriesdata)|[binOptions](/javascript/api/excel/excel.chartseriesdata#binoptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。 読み取り専用です。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesdata#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。 読み取り専用です。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesdata#bubblescale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesdata#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesdata#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesdata#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesdata#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesdata#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesdata#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesdata#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesdata#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesdata#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesdata#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseriesdata#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[mapOptions](/javascript/api/excel/excel.chartseriesdata#mapoptions)|リージョン マップ グラフのオプションをカプセル化します。 読み取り専用です。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesdata#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesdata#showconnectorlines)|ウォーターフォール図で、コネクタが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesdata#showleaderlines)|系列の各データ ラベルの引き出し線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseriesdata#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesdata#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesdata#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
|[Chart系列 Loadoptions](/javascript/api/excel/excel.chartseriesloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriesloadoptions#binoptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesloadoptions#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesloadoptions#bubblescale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesloadoptions#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseriesloadoptions#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[mapOptions](/javascript/api/excel/excel.chartseriesloadoptions#mapoptions)|リージョン マップ グラフのオプションをカプセル化します。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesloadoptions#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesloadoptions#showconnectorlines)|ウォーターフォール図で、コネクタが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesloadoptions#showleaderlines)|系列の各データ ラベルの引き出し線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseriesloadoptions#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[binOptions](/javascript/api/excel/excel.chartseriesupdatedata#binoptions)|ヒストグラム図とパレート図のビンのオプションをカプセル化します。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesupdatedata#boxwhiskeroptions)|箱ひげ図グラフのオプションをカプセル化します。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesupdatedata#bubblescale)|既定のサイズのパーセンテージを表す 0 (ゼロ) から 300 までの整数値とすることができます。 このプロパティは、バブルチャートにのみ使用できます。 読み取り/書き込み可能。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumcolor)|リージョン マップ グラフ系列の最大値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumtype)|リージョン マップ グラフ系列の最大値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumvalue)|リージョン マップ グラフ系列の最大値を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointcolor)|リージョン マップ グラフ系列の中間値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointtype)|リージョン マップ グラフ系列の中間値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointvalue)|リージョン マップ グラフ系列の中間値を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumcolor)|リージョン マップ グラフ系列の最小値の色を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumtype)|リージョン マップ グラフ系列の最小値の種類を設定または返します。 読み取り/書き込み可能。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumvalue)|リージョン マップ グラフ系列の最小値を設定または返します。 読み取り/書き込み可能。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesupdatedata#gradientstyle)|リージョン マップ グラフのグラデーション スタイルを設定または返します。 読み取り/書き込み可能。|
||[invertColor](/javascript/api/excel/excel.chartseriesupdatedata#invertcolor)|系列の負のデータ ポイントに対して塗りつぶしの色を設定または返します。 読み取り/書き込み可能。|
||[mapOptions](/javascript/api/excel/excel.chartseriesupdatedata#mapoptions)|リージョン マップ グラフのオプションをカプセル化します。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesupdatedata#parentlabelstrategy)|ツリーマップ グラフの系列上位ラベル方法領域を設定または返します。 読み取り/書き込み可能。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesupdatedata#showconnectorlines)|ウォーターフォール図で、コネクタが表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesupdatedata#showleaderlines)|系列の各データ ラベルの引き出し線が表示されるかどうかを指定します。 読み取り/書き込み可能。|
||[splitValue](/javascript/api/excel/excel.chartseriesupdatedata#splitvalue)|補助のグラフ (円または縦棒) が付いた円グラフを 2 つの部分に区切るしきい値を設定または返します。 読み取り/書き込み可能。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#xerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#yerrorbars)|グラフ系列の誤差範囲オブジェクトを表します。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#linknumberformat)|(セル内で変更されたときにラベルの数値形式が変わるように) 数値形式がセルにリンクされているかどうかを表すブール値です。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[pivotOptions](/javascript/api/excel/excel.chartupdatedata#pivotoptions)|ピボット グラフのオプションをカプセル化します。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|`columnIndex` プロパティを表します。|
|[ColumnPropertiesLoadOptions](/javascript/api/excel/excel.columnpropertiesloadoptions)|[columnHidden](/javascript/api/excel/excel.columnpropertiesloadoptions#columnhidden)|`columnHidden`プロパティに読み込むかどうかを指定します。|
||[columnIndex](/javascript/api/excel/excel.columnpropertiesloadoptions#columnindex)|`columnIndex`プロパティに読み込むかどうかを指定します。|
||[columnWidth](/javascript/api/excel/excel.columnpropertiesloadoptions#columnwidth)||
||[format: Excel. CellPropertiesFormatLoadOptions & {
            columnWidth?](/javascript/api/excel/excel.columnpropertiesloadoptions # 形式)|`format`プロパティに読み込むかどうかを指定します。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|1 つまたは複数の長方形範囲で構成され、条件付き書式が適用された RangeAreas を返します。 読み取り専用です。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは ItemNotFound エラーがスローされます。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|1 つまたは複数の長方形範囲で構成され、無効なセル値を含む RangeAreas を返します。 すべてのセル値が有効な場合、この関数からは null が返されます。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|リッチな値にリッチなフィルターを適用する場合、フィルターによって使用されるプロパティです。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|図形 ID を返します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|幾何学的図形の Shape オブジェクトを返します。 読み取り専用です。|
|[Geometricshape データ](/javascript/api/excel/excel.geometricshapedata)|[id](/javascript/api/excel/excel.geometricshapedata#id)|図形 ID を返します。 読み取り専用です。|
|[Geometricshape Loadoptions](/javascript/api/excel/excel.geometricshapeloadoptions)|[$all](/javascript/api/excel/excel.geometricshapeloadoptions#$all)||
||[id](/javascript/api/excel/excel.geometricshapeloadoptions#id)|図形 ID を返します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.geometricshapeloadoptions#shape)|幾何学的図形の Shape オブジェクトを返します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|図形グループの図形の数を返します。 読み取り専用です。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|コレクション内の位置に基づいて図形を取得します。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[$all](/javascript/api/excel/excel.groupshapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttextdescription)|コレクション内の各アイテムについて: Shape オブジェクトの別の説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttexttitle)|コレクション内の各アイテムについて: Shape オブジェクトの代替タイトルテキストを取得または設定します。|
||[connectionSiteCount](/javascript/api/excel/excel.groupshapecollectionloadoptions#connectionsitecount)|コレクション内の各アイテムについて: この図形の接続サイト数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.groupshapecollectionloadoptions#fill)|コレクション内の各アイテムについて: この図形の塗りつぶしの書式設定を返します。|
||[geometricShape](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshape)|コレクション内の各アイテムについて: 図形に関連付けられたジオメトリック図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[geometricShapeType](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshapetype)|コレクション内の各アイテムについて: この幾何学図形の幾何学的な図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[group](/javascript/api/excel/excel.groupshapecollectionloadoptions#group)|コレクション内の各アイテムについて: 図形に関連付けられている図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[height](/javascript/api/excel/excel.groupshapecollectionloadoptions#height)|コレクション内の各項目について: 図形の高さをポイント単位で表します。|
||[id](/javascript/api/excel/excel.groupshapecollectionloadoptions#id)|コレクション内の各アイテムについて: 図形の識別子を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.groupshapecollectionloadoptions#image)|コレクション内の各アイテムについて: 図形に関連付けられているイメージを返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[left](/javascript/api/excel/excel.groupshapecollectionloadoptions#left)|コレクション内の各項目について、図形の左側からワークシートの左端までの距離をポイント単位で指定します。|
||[level](/javascript/api/excel/excel.groupshapecollectionloadoptions#level)|コレクション内の各項目について: 指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.groupshapecollectionloadoptions#line)|コレクション内の各アイテムについて: 図形に関連付けられている線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.groupshapecollectionloadoptions#lineformat)|コレクション内の各アイテムについて: この図形の線の書式設定を返します。|
||[lockAspectRatio](/javascript/api/excel/excel.groupshapecollectionloadoptions#lockaspectratio)|コレクション内の各アイテムについて: この図形の縦横比がロックされているかどうかを指定します。|
||[name](/javascript/api/excel/excel.groupshapecollectionloadoptions#name)|コレクション内の各アイテムについて: 図形の名前を表します。|
||[parentGroup](/javascript/api/excel/excel.groupshapecollectionloadoptions#parentgroup)|コレクション内の各アイテムについて: この図形の親グループを表します。|
||[rotation](/javascript/api/excel/excel.groupshapecollectionloadoptions#rotation)|コレクション内の各項目の場合: 図形の回転角度を度で表します。|
||[textFrame](/javascript/api/excel/excel.groupshapecollectionloadoptions#textframe)|コレクション内の各アイテムについて: この図形のテキストフレームオブジェクトを返します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.groupshapecollectionloadoptions#top)|コレクション内の各項目について、図形の上端からワークシートの上端までの距離をポイント単位で指定します。|
||[type](/javascript/api/excel/excel.groupshapecollectionloadoptions#type)|コレクション内の各アイテムについて: この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.groupshapecollectionloadoptions#visible)|コレクション内の各アイテムについて: この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.groupshapecollectionloadoptions#width)|コレクション内の各項目について: 図形の幅をポイント単位で表します。|
||[zOrderPosition](/javascript/api/excel/excel.groupshapecollectionloadoptions#zorderposition)|コレクション内の各アイテムについて: z オーダーで指定した図形の位置を返します。0は順序スタックの最下部を表します。 読み取り専用です。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
||[set (プロパティ: Excel. HeaderFooter)](/javascript/api/excel/excel.headerfooter#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: HeaderFooterUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.headerfooter#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[Headerフッターデータ](/javascript/api/excel/excel.headerfooterdata)|[centerFooter](/javascript/api/excel/excel.headerfooterdata#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooterdata#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooterdata#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooterdata#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooterdata#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooterdata#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[set (properties: Excel. Headerフッターグループ)](/javascript/api/excel/excel.headerfootergroup#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: HeaderFooterGroupUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.headerfootergroup#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Headerフッター Groupdata](/javascript/api/excel/excel.headerfootergroupdata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupdata#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroupdata#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroupdata#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroupdata#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroupdata#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupdata#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupdata#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Headerheadergrou道路 Adoptions](/javascript/api/excel/excel.headerfootergrouploadoptions)|[$all](/javascript/api/excel/excel.headerfootergrouploadoptions#$all)||
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergrouploadoptions#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergrouploadoptions#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergrouploadoptions#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergrouploadoptions#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergrouploadoptions#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[HeaderFooterGroupUpdateData](/javascript/api/excel/excel.headerfootergroupupdatedata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupupdatedata#defaultforallpages)|偶数/奇数または最初のページが指定されていない場合にすべてのページに使用される汎用ヘッダー/フッター。|
||[evenPages](/javascript/api/excel/excel.headerfootergroupupdatedata#evenpages)|偶数ページに使用するヘッダー/フッター。奇数ページには奇数のヘッダー/フッターを指定する必要があります。|
||[firstPage](/javascript/api/excel/excel.headerfootergroupupdatedata#firstpage)|最初のページに使用するヘッダー/フッター。その他すべてのページには汎用または偶数/奇数のヘッダー/フッターが使用されます。|
||[oddPages](/javascript/api/excel/excel.headerfootergroupupdatedata#oddpages)|奇数ページに使用するヘッダー/フッター。偶数ページには偶数のヘッダー/フッターを指定する必要があります。|
||[state](/javascript/api/excel/excel.headerfootergroupupdatedata#state)|設定されているヘッダー/フッターの状態を取得または設定します。 詳細については、Excel.HeaderFooterState をご覧ください。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetmargins)|ワークシートのページ レイアウト オプションに設定されているページ余白に合わせてヘッダー/フッターの位置が調整されているかどうかを示すフラグを取得または設定します。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetscale)|ワークシートのページ レイアウト オプションに設定されているページ パーセンテージ スケールによってヘッダー/フッターが調整されているかどうかを示すフラグを取得または設定します。|
|[Headerフッター Loadoptions](/javascript/api/excel/excel.headerfooterloadoptions)|[$all](/javascript/api/excel/excel.headerfooterloadoptions#$all)||
||[centerFooter](/javascript/api/excel/excel.headerfooterloadoptions#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooterloadoptions#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooterloadoptions#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooterloadoptions#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooterloadoptions#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooterloadoptions#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
|[HeaderFooterUpdateData](/javascript/api/excel/excel.headerfooterupdatedata)|[centerFooter](/javascript/api/excel/excel.headerfooterupdatedata#centerfooter)|ワークシートの中央フッターを取得または設定します。|
||[centerHeader](/javascript/api/excel/excel.headerfooterupdatedata#centerheader)|ワークシートの中央ヘッダーを取得または設定します。|
||[leftFooter](/javascript/api/excel/excel.headerfooterupdatedata#leftfooter)|ワークシートの左フッターを取得または設定します。|
||[leftHeader](/javascript/api/excel/excel.headerfooterupdatedata#leftheader)|ワークシートの左ヘッダーを取得または設定します。|
||[rightFooter](/javascript/api/excel/excel.headerfooterupdatedata#rightfooter)|ワークシートの右フッターを取得または設定します。|
||[rightHeader](/javascript/api/excel/excel.headerfooterupdatedata#rightheader)|ワークシートの右ヘッダーを取得または設定します。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|画像の形式を返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.image#id)|画像オブジェクトの図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.image#shape)|画像に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
|[ImageData](/javascript/api/excel/excel.imagedata)|[format](/javascript/api/excel/excel.imagedata#format)|画像の形式を返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.imagedata#id)|画像オブジェクトの図形 ID を表します。 読み取り専用です。|
|[ImageLoadOptions](/javascript/api/excel/excel.imageloadoptions)|[$all](/javascript/api/excel/excel.imageloadoptions#$all)||
||[format](/javascript/api/excel/excel.imageloadoptions#format)|画像の形式を返します。 読み取り専用です。|
||[id](/javascript/api/excel/excel.imageloadoptions#id)|画像オブジェクトの図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.imageloadoptions#shape)|画像に関連付けられた Shape オブジェクトを返します。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
||[set (properties: IterativeCalculation)](/javascript/api/excel/excel.iterativecalculation#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: IterativeCalculationUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.iterativecalculation#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[IterativeCalculationData](/javascript/api/excel/excel.iterativecalculationdata)|[enabled](/javascript/api/excel/excel.iterativecalculationdata#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationdata#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationdata#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
|[IterativeCalculationLoadOptions](/javascript/api/excel/excel.iterativecalculationloadoptions)|[$all](/javascript/api/excel/excel.iterativecalculationloadoptions#$all)||
||[enabled](/javascript/api/excel/excel.iterativecalculationloadoptions#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationloadoptions#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationloadoptions#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
|[IterativeCalculationUpdateData](/javascript/api/excel/excel.iterativecalculationupdatedata)|[enabled](/javascript/api/excel/excel.iterativecalculationupdatedata#enabled)|Excel で反復計算を使用して循環参照を解決する場合、true となります。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationupdatedata#maxchange)|循環参照は Excel の反復計算によって解決されます。その反復計算間の変化の最大値を設定または返します。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationupdatedata#maxiteration)|Excel で循環参照の解決に使用できる、最大反復回数を設定または返します。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|指定されたコネクタの始点を指定された図形に接続します。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|指定されたコネクタの終点を指定された図形に接続します。|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|線のコネクタの種類を表します。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|指定されたコネクタの始点を図形から切り離します。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|指定されたコネクタの終点を図形から切り離します。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|指定された線の始点が接続されている図形を表します。 読み取り専用です。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|コネクタの始点が接続されている結合点を表します。 読み取り専用です。 線の始点がどの図形にも接続されていない場合は、null を返します。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|指定された線の終点が接続されている図形を表します。 読み取り専用です。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|コネクタの終点が接続されている結合点を表します。 読み取り専用です。 線の終点がどの図形にも接続されていない場合は、null を返します。|
||[id](/javascript/api/excel/excel.line#id)|図形 ID を表します。 読み取り専用です。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定された線の始点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定された線の終点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.line#shape)|線に関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
||[set (properties: Excel. Line)](/javascript/api/excel/excel.line#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: LineUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.line#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[LineData](/javascript/api/excel/excel.linedata)|[beginArrowheadLength](/javascript/api/excel/excel.linedata#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.linedata#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.linedata#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[beginConnectedSite](/javascript/api/excel/excel.linedata#beginconnectedsite)|コネクタの始点が接続されている結合点を表します。 読み取り専用です。 線の始点がどの図形にも接続されていない場合は、null を返します。|
||[connectorType](/javascript/api/excel/excel.linedata#connectortype)|線のコネクタの種類を表します。|
||[endArrowheadLength](/javascript/api/excel/excel.linedata#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.linedata#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.linedata#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
||[endConnectedSite](/javascript/api/excel/excel.linedata#endconnectedsite)|コネクタの終点が接続されている結合点を表します。 読み取り専用です。 線の終点がどの図形にも接続されていない場合は、null を返します。|
||[id](/javascript/api/excel/excel.linedata#id)|図形 ID を表します。 読み取り専用です。|
||[isBeginConnected](/javascript/api/excel/excel.linedata#isbeginconnected)|指定された線の始点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[isEndConnected](/javascript/api/excel/excel.linedata#isendconnected)|指定された線の終点が図形に接続されているかどうかを指定します。 読み取り専用です。|
|[LineLoadOptions](/javascript/api/excel/excel.lineloadoptions)|[$all](/javascript/api/excel/excel.lineloadoptions#$all)||
||[beginArrowheadLength](/javascript/api/excel/excel.lineloadoptions#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[beginConnectedShape](/javascript/api/excel/excel.lineloadoptions#beginconnectedshape)|指定された線の始点が接続されている図形を表します。|
||[beginConnectedSite](/javascript/api/excel/excel.lineloadoptions#beginconnectedsite)|コネクタの始点が接続されている結合点を表します。 読み取り専用です。 線の始点がどの図形にも接続されていない場合は、null を返します。|
||[connectorType](/javascript/api/excel/excel.lineloadoptions#connectortype)|線のコネクタの種類を表します。|
||[endArrowheadLength](/javascript/api/excel/excel.lineloadoptions#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
||[endConnectedShape](/javascript/api/excel/excel.lineloadoptions#endconnectedshape)|指定された線の終点が接続されている図形を表します。|
||[endConnectedSite](/javascript/api/excel/excel.lineloadoptions#endconnectedsite)|コネクタの終点が接続されている結合点を表します。 読み取り専用です。 線の終点がどの図形にも接続されていない場合は、null を返します。|
||[id](/javascript/api/excel/excel.lineloadoptions#id)|図形 ID を表します。 読み取り専用です。|
||[isBeginConnected](/javascript/api/excel/excel.lineloadoptions#isbeginconnected)|指定された線の始点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[isEndConnected](/javascript/api/excel/excel.lineloadoptions#isendconnected)|指定された線の終点が図形に接続されているかどうかを指定します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.lineloadoptions#shape)|線に関連付けられた Shape オブジェクトを返します。|
|[LineUpdateData](/javascript/api/excel/excel.lineupdatedata)|[beginArrowheadLength](/javascript/api/excel/excel.lineupdatedata#beginarrowheadlength)|指定された線の始点の矢印の長さを表します。|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#beginarrowheadstyle)|指定された線の始点の矢印のスタイルを表します。|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#beginarrowheadwidth)|指定された線の始点の矢印の幅を表します。|
||[connectorType](/javascript/api/excel/excel.lineupdatedata#connectortype)|線のコネクタの種類を表します。|
||[endArrowheadLength](/javascript/api/excel/excel.lineupdatedata#endarrowheadlength)|指定された線の終点の矢印の長さを表します。|
||[endArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#endarrowheadstyle)|指定された線の終点の矢印のスタイルを表します。|
||[endArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#endarrowheadwidth)|指定された線の終点の矢印の幅を表します。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|改ページ オブジェクトを削除します。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|改ページの後の最初のセルを取得します。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|改ページの行インデックスを表します。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|指定された範囲の左上セルの前に改ページを追加します。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|コレクション内の改ページの数を取得します。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|インデックス経由で改ページ オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|コレクション内の手動改ページをすべてリセットします。|
|[PageBreakCollectionLoadOptions](/javascript/api/excel/excel.pagebreakcollectionloadoptions)|[$all](/javascript/api/excel/excel.pagebreakcollectionloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#columnindex)|コレクション内の各アイテムについて: 改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#rowindex)|コレクション内の各アイテムについて: 改ページの行インデックスを表します。|
|[PageBreakData](/javascript/api/excel/excel.pagebreakdata)|[columnIndex](/javascript/api/excel/excel.pagebreakdata#columnindex)|改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreakdata#rowindex)|改ページの行インデックスを表します。|
|[PageBreakLoadOptions](/javascript/api/excel/excel.pagebreakloadoptions)|[$all](/javascript/api/excel/excel.pagebreakloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakloadoptions#columnindex)|改ページの列インデックスを表します。|
||[rowIndex](/javascript/api/excel/excel.pagebreakloadoptions#rowindex)|改ページの行インデックスを表します。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、ItemNotFound エラーがスローされます。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|ワークシートの印刷範囲を表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。 印刷範囲がない場合、null オブジェクトが返されます。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|タイトル列を表す範囲オブジェクトを取得します。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|タイトル列を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|タイトル行を表す範囲オブジェクトを取得します。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|タイトル行を表す範囲オブジェクトを取得します。 設定されていない場合、null オブジェクトが返されます。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[set (properties: PageLayout)](/javascript/api/excel/excel.pagelayout#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PageLayoutUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pagelayout#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|ワークシートの印刷範囲を設定します。|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|ワークシートのページ余白を単位で設定します。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|ワークシートのページ余白を単位で設定します。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|セルを含む列を、印刷時、ワークシートの各ページの左で繰り返すように設定します。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|セルを含む行を、印刷時、ワークシートの各ページの上で繰り返すように設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutData](/javascript/api/excel/excel.pagelayoutdata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutdata#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutdata#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutdata#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutdata#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayoutdata#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutdata#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutdata#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutdata#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutdata#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutdata#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayoutdata#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayoutdata#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayoutdata#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayoutdata#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutdata#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutdata#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayoutdata#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutdata#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayoutdata#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayoutdata#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutLoadOptions](/javascript/api/excel/excel.pagelayoutloadoptions)|[$all](/javascript/api/excel/excel.pagelayoutloadoptions#$all)||
||[blackAndWhite](/javascript/api/excel/excel.pagelayoutloadoptions#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutloadoptions#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutloadoptions#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutloadoptions#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayoutloadoptions#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutloadoptions#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutloadoptions#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutloadoptions#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayoutloadoptions#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayoutloadoptions#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayoutloadoptions#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayoutloadoptions#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutloadoptions#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutloadoptions#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayoutloadoptions#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutloadoptions#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayoutloadoptions#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayoutloadoptions#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|印刷時に使用するように指定された単位でページ レイアウトの下余白を表します。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|印刷時に使用するように指定された単位でページ レイアウトのフッター余白を表します。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|印刷時に使用するように指定された単位でページ レイアウトのヘッダー余白を表します。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|印刷時に使用するように指定された単位でページ レイアウトの左余白を表します。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|印刷時に使用するように指定された単位でページ レイアウトの右余白を表します。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|印刷時に使用するように指定された単位でページ レイアウトの上余白を表します。|
|[PageLayoutUpdateData](/javascript/api/excel/excel.pagelayoutupdatedata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutupdatedata#blackandwhite)|ワークシートの白黒印刷オプションを取得または設定します。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutupdatedata#bottommargin)|ポイント単位印刷に使用するワークシートの下部ページ余白を取得または設定します。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutupdatedata#centerhorizontally)|ワークシートの [ページ中央] の [水平] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を水平に設定するかどうかが決定されます。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutupdatedata#centervertically)|ワークシートの [ページ中央] の [垂直] フラグを取得または設定します。 このフラグによって、印刷時、ワークシートのページ中央を垂直に設定するかどうかが決定されます。|
||[draftMode](/javascript/api/excel/excel.pagelayoutupdatedata#draftmode)|ワークシートの下書きモード オプションを取得または設定します。 true の場合、グラフィックスなしでシートが印刷されます。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutupdatedata#firstpagenumber)|印刷するワークシートの最初のページ番号を取得または設定します。 null 値は "自動" ページ番号を表します。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#footermargin)|印刷時に使用するワークシートのフッター余白 (ポイント数) を取得または設定します。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#headermargin)|印刷時に使用するワークシートのヘッダー余白 (ポイント数) を取得または設定します。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutupdatedata#headersfooters)|ワークシートのヘッダーとフッターの構成。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutupdatedata#leftmargin)|印刷時に使用するワークシートの左余白 (ポイント数) を取得または設定します。|
||[orientation](/javascript/api/excel/excel.pagelayoutupdatedata#orientation)|ワークシートのページの向きを取得または設定します。|
||[paperSize](/javascript/api/excel/excel.pagelayoutupdatedata#papersize)|ワークシートのページの用紙サイズを取得または設定します。|
||[printComments](/javascript/api/excel/excel.pagelayoutupdatedata#printcomments)|印刷時、ワークシートのコメントを表示するかどうかを取得または設定します。|
||[printErrors](/javascript/api/excel/excel.pagelayoutupdatedata#printerrors)|ワークシートの印刷エラー オプションを取得または設定します。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutupdatedata#printgridlines)|ワークシートの印刷目盛線フラグを取得または設定します。 このフラグによって、目盛線を印刷するかどうかが決定されます。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutupdatedata#printheadings)|ワークシートの見出し印刷フラグを取得または設定します。 このフラグによって、見出しを印刷するかどうかが決定されます。|
||[printOrder](/javascript/api/excel/excel.pagelayoutupdatedata#printorder)|ワークシートのページ印刷順序オプションを取得または設定します。 これによって、印刷されるページ番号の処理に使用する順序が指定されます。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutupdatedata#rightmargin)|印刷時に使用するワークシートの右余白 (ポイント数) を取得または設定します。|
||[topMargin](/javascript/api/excel/excel.pagelayoutupdatedata#topmargin)|印刷時に使用するワークシートの上余白 (ポイント数) を取得または設定します。|
||[zoom](/javascript/api/excel/excel.pagelayoutupdatedata#zoom)|ワークシートの拡大印刷オプションを取得または設定します。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|横方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|印刷ページのスケール値は 10 から 400 までです。 縦または横方向にページを合わせるように指定されている場合、この値には null を指定できます。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|縦方向に合わせるページ数。 パーセンテージ スケールが使用される場合、この値には null を指定できます。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues (sortBy: "昇順" \| "降順", valuehierarchy: DataPivotHierarchy, pivotItemScope?: Array<PivotItem \|文字列>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。 この範囲によって、並べ替えに使用する特定の値が定義されます。|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|与えられた範囲で、指定された値に基づいて PivotField を並べ替えます。 この範囲によって、並べ替えに使用する特定の値が定義されます。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定されるかどうかを指定します。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|PivotTable 内で指定された範囲の値を計算するために使用される DataHierarchy を取得します。|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|PivotTable 内で指定された範囲の値を構成する PivotItems を軸から取得します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存されるかどうかを指定します。|
||[setAutoSortOnCell (cell: Range \| String, sortBy: "昇順" \| "降順")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|必要なすべての条件とコンテキストを自動的に選択するため、指定したセルを使用して自動的に並べ替えを実行するようピボットテーブルを設定します。 これは、UI から自動並べ替えを適用するのと同じ動作です。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|必要なすべての条件とコンテキストを自動的に選択するため、指定したセルを使用して自動的に並べ替えを実行するようピボットテーブルを設定します。 これは、UI から自動並べ替えを適用するのと同じ動作です。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutdata#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定されるかどうかを指定します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutdata#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存されるかどうかを指定します。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[autoFormat](/javascript/api/excel/excel.pivotlayoutloadoptions#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定されるかどうかを指定します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutloadoptions#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存されるかどうかを指定します。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutupdatedata#autoformat)|更新時またはフィールドの削除時に書式が自動的に設定されるかどうかを指定します。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutupdatedata#preserveformatting)|ピボット、並べ替え、ページ フィールド項目の変更などの操作によってレポートが更新または再計算されたとき、書式設定が保存されるかどうかを指定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|ピボットテーブルでユーザーがデータの値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|ピボット テーブルで並べ替えを実行する際にユーザー設定リストを使用するかどうかを指定します。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottablecollectionloadoptions#enabledatavalueediting)|コレクション内の各アイテムについて: ピボットテーブルで、ユーザーがデータ本文の値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottablecollectionloadoptions#usecustomsortlists)|コレクション内の各アイテムについて: 並べ替えの際に、ピボットテーブルでカスタムリストを使用するかどうかを指定します。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottabledata#enabledatavalueediting)|ピボットテーブルでユーザーがデータの値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottabledata#usecustomsortlists)|ピボット テーブルで並べ替えを実行する際にユーザー設定リストを使用するかどうかを指定します。|
|[ピボットのオプション](/javascript/api/excel/excel.pivottableloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableloadoptions#enabledatavalueediting)|ピボットテーブルでユーザーがデータの値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableloadoptions#usecustomsortlists)|ピボット テーブルで並べ替えを実行する際にユーザー設定リストを使用するかどうかを指定します。|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableupdatedata#enabledatavalueediting)|ピボットテーブルでユーザーがデータの値を編集できるようにするかどうかを指定します。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableupdatedata#usecustomsortlists)|ピボット テーブルで並べ替えを実行する際にユーザー設定リストを使用するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|現在の範囲から対象の範囲までの範囲に値を設定します。|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|現在の範囲から対象の範囲までの範囲に値を設定します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|データ型を含む範囲セルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|ワークシート内で範囲セルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の範囲にセル データまたは書式設定をコピーします。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の範囲にセル データまたは書式設定をコピーします。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|指定された条件に基づいて指定された文字列を見つけます。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|現在の範囲に対してフラッシュ フィルを実行します。フラッシュ フィルでは、パターンを感知して自動的にデータが設定されるので、範囲は単一列範囲で、かつパターンを検出できるように周囲にデータが存在する必要があります。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|2D 配列を返します。各セルのフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|一次元配列を返します。各列のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された列内の列間で一貫性のないプロパティについては、null が返されます。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|一次元配列を返します。各行のフォント、塗りつぶし、罫線、配置などのプロパティ データをカプセル化します。  指定された行内の列間で一貫性のないプロパティについては、null が返されます。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の長方形範囲で構成される RangeAreas オブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の範囲で構成される、RangeAreas オブジェクトを取得します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表し、1 つまたは複数の範囲を構成する RangeAreas オブジェクトを取得します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|範囲と重なるテーブルの集まりを範囲限定で取得します。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|各セルのデータ型の状態を表します。 読み取り専用です。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|列によって指定される範囲から重複する値を削除します。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|現在の範囲内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|セル プロパティの 2D 配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|列プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|次の再計算が発生したときに再計算する範囲を設定します。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|行プロパティの一次元配列に基づいて範囲を更新します。フォント、塗りつぶし、罫線、配置などをカプセル化します。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|RangeAreas のすべてのセルを計算します。|
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|この RangeAreas オブジェクトを構成する各領域で値、フォーマット、塗りつぶし、罫線などを消去します。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|この RangeAreas オブジェクトを構成する各領域で値、フォーマット、塗りつぶし、罫線などを消去します。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|RangeAreas 内でデータ型を含むすべてのセルをテキストに変換します。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|RangeAreas 内のすべてのセルをリンク付きデータ型に変換します。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の RangeAreas にセル データまたは書式設定をコピーします。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|ソース範囲または RangeAreas から現在の RangeAreas にセル データまたは書式設定をコピーします。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|RangeAreas の列全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11, H2" を表す場合、列 "B:E, H:H" を表す RangeAreas が返されます)。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|RangeAreas の行全体を表す RangeAreas オブジェクトを返します (たとえば、現在の RangeAreas がセル "B4:E11" を表す場合、行 "4:11" を表す RangeAreas が返されます)。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、ItemNotFound エラーがスローされます。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|指定した範囲または RangeAreas の交差を表す RangeAreas オブジェクトを返します。 交差が見つからない場合、null オブジェクトが返されます。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|特定の行と列のオフセットによってシフトされる RangeAreas オブジェクトを返します。 返される RangeAreas のディメンションは元のオブジェクトと一致します。 結果の RangeAreas がワークシート グリッドの境界線の外にはみ出る場合、エラーがスローされます。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、エラーがスローされます。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、エラーがスローされます。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、null オブジェクトを返します。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|指定された型と値に一致するすべてのセルを表す RangeAreas オブジェクトを返します。 条件に一致する特別なセルが見つからない場合、null オブジェクトを返します。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|この RangeAreas オブジェクトの範囲と重なるテーブルの集まりを範囲限定で返します。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|RangeAreas オブジェクトの個別の長方形範囲の全使用済み領域を構成する使用済み RangeAreas を返します。|
||[address](/javascript/api/excel/excel.rangeareas#address)|A1 スタイルで RageAreas 参照を返します。 address 値にはセルの各長方形ブロックのワークシート名が含まれます ("Sheet1!A1:B4, Sheet1!D1:D4" など)。 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|ユーザー ロケールで RageAreas 参照を返します。 読み取り専用です。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|この RangeAreas オブジェクトを構成する長方形範囲の数を返します。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|この RangeAreas オブジェクトを構成する長方形範囲の集まりを返します。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|RangeAreas オブジェクトのセル数を返します。すべての個別長方形範囲のセル数が合計されます。 セル数が 2^31-1 (2,147,483,647) を超える場合、-1 を返します。 読み取り専用です。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|この RangeAreas オブジェクトのセルと交差する ConditionalFormats の集まりを返します。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareas#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。 読み取り専用です。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|この RangeAreas オブジェクトの全範囲が列全体を表すかどうかを示します ("A:C, Q:Z" など)。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|この RangeAreas オブジェクトの全範囲が行全体を表すかどうかを示します ("1:3, 5:7" など)。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|現在の RangeAreas のワークシートを返します。 読み取り専用です。|
||[set (properties: Excel. RangeAreas)](/javascript/api/excel/excel.rangeareas#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeAreasUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangeareas#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|次の再計算が発生したときに再計算する RangeAreas を設定します。|
||[style](/javascript/api/excel/excel.rangeareas#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|ドキュメントの環境変更に基づいて自動的に調整する目的でオブジェクトを追跡します。 これは context.trackedObjects.add(thisObject) 呼び出しの省略形です。 ".sync" 呼び出し間で、かつ ".run" バッチの連続実行の外でこのオブジェクトを使用しているとき、オブジェクトであるプロパティを設定したか、あるメソッドを呼び出したときに "InvalidObjectPath" エラーが表示される場合、オブジェクトを最初に作成したときに、追跡対象オブジェクトの集まりにそのオブジェクトを追加しておく必要がありました。|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|前に追跡されていた場合、このオブジェクトに関連付けられているメモリを解放します。 これは context.trackedObjects.remove(thisObject) 呼び出しの省略形です。 追跡対象オブジェクトが多いとホスト アプリケーションの動作が遅くなります。追加したオブジェクトが不要になったら、必ずそれを解放してください。 メモリ リリースを有効にするには、"context.sync()" を先に呼び出す必要があります。|
|[RangeAreasData](/javascript/api/excel/excel.rangeareasdata)|[address](/javascript/api/excel/excel.rangeareasdata#address)|A1 スタイルで RageAreas 参照を返します。 address 値にはセルの各長方形ブロックのワークシート名が含まれます ("Sheet1!A1:B4, Sheet1!D1:D4" など)。 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeareasdata#addresslocal)|ユーザー ロケールで RageAreas 参照を返します。 読み取り専用です。|
||[areaCount](/javascript/api/excel/excel.rangeareasdata#areacount)|この RangeAreas オブジェクトを構成する長方形範囲の数を返します。|
||[areas](/javascript/api/excel/excel.rangeareasdata#areas)|この RangeAreas オブジェクトを構成する長方形範囲の集まりを返します。|
||[cellCount](/javascript/api/excel/excel.rangeareasdata#cellcount)|RangeAreas オブジェクトのセル数を返します。すべての個別長方形範囲のセル数が合計されます。 セル数が 2^31-1 (2,147,483,647) を超える場合、-1 を返します。 読み取り専用です。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareasdata#conditionalformats)|この RangeAreas オブジェクトのセルと交差する ConditionalFormats の集まりを返します。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangeareasdata#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareasdata#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。 読み取り専用です。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasdata#isentirecolumn)|この RangeAreas オブジェクトの全範囲が列全体を表すかどうかを示します ("A:C, Q:Z" など)。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeareasdata#isentirerow)|この RangeAreas オブジェクトの全範囲が行全体を表すかどうかを示します ("1:3, 5:7" など)。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangeareasdata#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
|[RangeAreasLoadOptions](/javascript/api/excel/excel.rangeareasloadoptions)|[$all](/javascript/api/excel/excel.rangeareasloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeareasloadoptions#address)|A1 スタイルで RageAreas 参照を返します。 address 値にはセルの各長方形ブロックのワークシート名が含まれます ("Sheet1!A1:B4, Sheet1!D1:D4" など)。 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangeareasloadoptions#addresslocal)|ユーザー ロケールで RageAreas 参照を返します。 読み取り専用です。|
||[areaCount](/javascript/api/excel/excel.rangeareasloadoptions#areacount)|この RangeAreas オブジェクトを構成する長方形範囲の数を返します。|
||[cellCount](/javascript/api/excel/excel.rangeareasloadoptions#cellcount)|RangeAreas オブジェクトのセル数を返します。すべての個別長方形範囲のセル数が合計されます。 セル数が 2^31-1 (2,147,483,647) を超える場合、-1 を返します。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangeareasloadoptions#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareasloadoptions#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasloadoptions#isentirecolumn)|この RangeAreas オブジェクトの全範囲が列全体を表すかどうかを示します ("A:C, Q:Z" など)。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangeareasloadoptions#isentirerow)|この RangeAreas オブジェクトの全範囲が行全体を表すかどうかを示します ("1:3, 5:7" など)。 読み取り専用です。|
||[style](/javascript/api/excel/excel.rangeareasloadoptions#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
||[worksheet](/javascript/api/excel/excel.rangeareasloadoptions#worksheet)|現在の RangeAreas のワークシートを返します。|
|[RangeAreasUpdateData](/javascript/api/excel/excel.rangeareasupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeareasupdatedata#datavalidation)|RangeAreas の全範囲に対して dataValidation オブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangeareasupdatedata#format)|rangeFormat オブジェクトを返します。RangeAreas オブジェクトの全範囲を対象にフォント、塗りつぶし、罫線、配置などのプロパティをカプセル化します。|
||[style](/javascript/api/excel/excel.rangeareasupdatedata#style)|この RangeAreas オブジェクトの全範囲のスタイルを表します。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionloadoptions#tintandshade)|コレクション内の各アイテムについて: 範囲の境界線の色を明るくまたは暗くする倍精度浮動小数点型 (double) の値を設定します。値は、-1 (最も暗い) から 1 (最も明るい) までの範囲で、元の色は0です。|
|[RangeBorderCollectionUpdateData](/javascript/api/excel/excel.rangebordercollectionupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionupdatedata#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[tintAndShade](/javascript/api/excel/excel.rangeborderdata#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangeborderloadoptions#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangeborderupdatedata#tintandshade)|範囲の境界線の色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|RangeCollection 内の範囲数を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|RangeCollection 内のその位置に基づいて範囲オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[$all](/javascript/api/excel/excel.rangecollectionloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangecollectionloadoptions#address)|コレクション内の各アイテムについて: A1 形式の範囲参照を表します。 Address 値にはシート参照が含まれます (例: "Sheet1!A1: B4 ") 読み取り専用です。|
||[addressLocal](/javascript/api/excel/excel.rangecollectionloadoptions#addresslocal)|コレクション内の各アイテムについて: ユーザーの言語で指定された範囲の範囲参照を表します。 読み取り専用です。|
||[cellCount](/javascript/api/excel/excel.rangecollectionloadoptions#cellcount)|コレクション内の各項目について: 範囲内のセルの数。 セルの数が 2^31-1 (2,147,483,647) を超えると、この API は -1 を返します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangecollectionloadoptions#columncount)|コレクション内の各アイテムについて: 範囲内の列の合計数を表します。 読み取り専用です。|
||[columnHidden](/javascript/api/excel/excel.rangecollectionloadoptions#columnhidden)|コレクション内の各アイテムについて: 現在の範囲のすべての列が非表示になっているかどうかを表します。|
||[columnIndex](/javascript/api/excel/excel.rangecollectionloadoptions#columnindex)|コレクション内の各アイテムについて: 範囲内の最初のセルの列番号を表します。 0 を起点とする番号になります。 読み取り専用です。|
||[dataValidation](/javascript/api/excel/excel.rangecollectionloadoptions#datavalidation)|コレクション内の各アイテムについて: データバリデーションオブジェクトを返します。|
||[format](/javascript/api/excel/excel.rangecollectionloadoptions#format)|コレクション内の各アイテムについて: 範囲のフォント、塗りつぶし、罫線、配置、およびその他のプロパティをカプセル化する format オブジェクトを返します。|
||[formulas](/javascript/api/excel/excel.rangecollectionloadoptions#formulas)|コレクション内の各アイテムについて: A1 形式の表記で数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangecollectionloadoptions#formulaslocal)|コレクション内の各項目について、: ユーザーの言語と書式設定ロケールで、A1 形式の表記法の数式を表します。  たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangecollectionloadoptions#formulasr1c1)|コレクション内の各項目について、: R1C1 形式の表記法で数式を表します。|
||[hidden](/javascript/api/excel/excel.rangecollectionloadoptions#hidden)|コレクション内の各アイテムについて: 現在の範囲のすべてのセルが非表示になっているかどうかを表します。 読み取り専用です。|
||[hyperlink](/javascript/api/excel/excel.rangecollectionloadoptions#hyperlink)|コレクション内の各アイテムについて: 現在の範囲のハイパーリンクを表します。|
||[isEntireColumn](/javascript/api/excel/excel.rangecollectionloadoptions#isentirecolumn)|コレクション内の各アイテムについて: 現在の範囲が列全体であるかどうかを表します。 読み取り専用です。|
||[isEntireRow](/javascript/api/excel/excel.rangecollectionloadoptions#isentirerow)|コレクション内の各アイテムについて: 現在の範囲が行全体であるかどうかを表します。 読み取り専用です。|
||[linkedDataTypeState](/javascript/api/excel/excel.rangecollectionloadoptions#linkeddatatypestate)|コレクション内の各アイテムについて: 各セルのデータ型の状態を表します。 読み取り専用です。|
||[numberFormat](/javascript/api/excel/excel.rangecollectionloadoptions#numberformat)|コレクション内の各アイテムについて: 指定された範囲の Excel の数値書式コードを表します。|
||[numberFormatLocal](/javascript/api/excel/excel.rangecollectionloadoptions#numberformatlocal)|コレクション内の各アイテムについて: 指定された範囲の Excel の番号書式コードを、ユーザーの言語の文字列で表します。|
||[rowCount](/javascript/api/excel/excel.rangecollectionloadoptions#rowcount)|コレクション内の各アイテムについて: 範囲内の行の合計数を返します。 読み取り専用です。|
||[rowHidden](/javascript/api/excel/excel.rangecollectionloadoptions#rowhidden)|コレクション内の各アイテムについて: 現在の範囲のすべての行が非表示になっているかどうかを表します。|
||[rowIndex](/javascript/api/excel/excel.rangecollectionloadoptions#rowindex)|コレクション内の各項目について: 範囲内の最初のセルの行番号を返します。 0 を起点とする番号になります。 読み取り専用。|
||[style](/javascript/api/excel/excel.rangecollectionloadoptions#style)|コレクション内の各アイテムについて: 現在の範囲のスタイルを表します。|
||[text](/javascript/api/excel/excel.rangecollectionloadoptions#text)|コレクション内の各項目について: 指定された範囲のテキスト値。 テキスト値は、セルの幅には依存しません。 Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。 読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangecollectionloadoptions#valuetypes)|コレクション内の各アイテムについて: 各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangecollectionloadoptions#values)|コレクション内の各項目について: 指定された範囲の生の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
||[worksheet](/javascript/api/excel/excel.rangecollectionloadoptions#worksheet)|コレクション内の各アイテムについて: 現在の範囲を含むワークシート。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[linkedDataTypeState](/javascript/api/excel/excel.rangedata#linkeddatatypestate)|各セルのデータ型の状態を表します。 読み取り専用です。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[pattern](/javascript/api/excel/excel.rangefilldata#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefilldata#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefilldata#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefilldata#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[pattern](/javascript/api/excel/excel.rangefillloadoptions#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefillloadoptions#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillloadoptions#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefillloadoptions#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[pattern](/javascript/api/excel/excel.rangefillupdatedata#pattern)|範囲のパターンを取得または設定します。 詳細については、Excel.FillPattern をご覧ください。 LinearGradient と RectangularGradient はサポートされていません。|
||[patternColor](/javascript/api/excel/excel.rangefillupdatedata#patterncolor)|Range パターンの色を表す HTML カラー コードを設定します。形式は #RRGGBB (例: "FFA500") か名前付き HTML 色 (例: "オレンジ") です。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillupdatedata#patterntintandshade)|範囲の塗りつぶしのパターン色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
||[tintAndShade](/javascript/api/excel/excel.rangefillupdatedata#tintandshade)|範囲の塗りつぶしの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[strikethrough](/javascript/api/excel/excel.rangefontdata#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefontdata#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefontdata#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefontdata#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[strikethrough](/javascript/api/excel/excel.rangefontloadoptions#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefontloadoptions#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefontloadoptions#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefontloadoptions#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[strikethrough](/javascript/api/excel/excel.rangefontupdatedata#strikethrough)|フォントの取り消し線の状態を表します。 null 値は、範囲全体に同じ取り消し線設定がないことを示します。|
||[subscript](/javascript/api/excel/excel.rangefontupdatedata#subscript)|フォントの下付きの状態を表します。|
||[superscript](/javascript/api/excel/excel.rangefontupdatedata#superscript)|フォントの上付きの状態を表します。|
||[tintAndShade](/javascript/api/excel/excel.rangefontupdatedata#tintandshade)|範囲のフォントの色を明るくするか、暗くする double 値を設定または返します。値は -1 が最も暗く、1 が最も明るくなります。元の色は 0 です。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[autoIndent](/javascript/api/excel/excel.rangeformatdata#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformatdata#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformatdata#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatdata#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[autoIndent](/javascript/api/excel/excel.rangeformatloadoptions#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformatloadoptions#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformatloadoptions#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatloadoptions#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[autoIndent](/javascript/api/excel/excel.rangeformatupdatedata#autoindent)|テキスト配置が均等割り付けに設定されている場合、テキストを自動的にインデントするかどうかを指定します。|
||[indentLevel](/javascript/api/excel/excel.rangeformatupdatedata#indentlevel)|インデント レベルを示す 0 から 250 までの整数。|
||[readingOrder](/javascript/api/excel/excel.rangeformatupdatedata#readingorder)|範囲に適用される読み上げ順序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatupdatedata#shrinktofit)|使用可能な列幅に収まるように自動的に文字列が縮小されるかどうかを示します。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[linkedDataTypeState](/javascript/api/excel/excel.rangeloadoptions#linkeddatatypestate)|各セルのデータ型の状態を表します。 読み取り専用です。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[RemoveDuplicatesResultData](/javascript/api/excel/excel.removeduplicatesresultdata)|[removed](/javascript/api/excel/excel.removeduplicatesresultdata#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultdata#uniqueremaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[RemoveDuplicatesResultLoadOptions](/javascript/api/excel/excel.removeduplicatesresultloadoptions)|[$all](/javascript/api/excel/excel.removeduplicatesresultloadoptions#$all)||
||[removed](/javascript/api/excel/excel.removeduplicatesresultloadoptions#removed)|操作によって削除された重複行の数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultloadoptions#uniqueremaining)|結果として生じた範囲に存在する残りの一意の行の数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|一致方法として完全一致か部分一致を指定します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|`address` プロパティを表します。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|`addressLocal` プロパティを表します。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|`rowIndex` プロパティを表します。|
|[RowPropertiesLoadOptions](/javascript/api/excel/excel.rowpropertiesloadoptions)|[format: Excel. CellPropertiesFormatLoadOptions & {
            rowHeight?](/javascript/api/excel/excel.rowpropertiesloadoptions # 形式)|`format`プロパティに読み込むかどうかを指定します。|
||[rowHeight](/javascript/api/excel/excel.rowpropertiesloadoptions#rowheight)||
||[rowHidden](/javascript/api/excel/excel.rowpropertiesloadoptions#rowhidden)|`rowHidden`プロパティに読み込むかどうかを指定します。|
||[rowIndex](/javascript/api/excel/excel.rowpropertiesloadoptions#rowindex)|`rowIndex`プロパティに読み込むかどうかを指定します。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 完全一致は、セルの内容全体に一致します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|検索の方向を指定します。 既定値は前方向です。 Excel.SearchDirection をご覧ください。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|`format` プロパティを表します。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|`hyperlink` プロパティを表します。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|`style` プロパティを表します。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|`columnHidden` プロパティを表します。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel. CellPropertiesFormat & {
            columnWidth?](/javascript/api/excel/excel.settablecolumnproperties # 形式)|`format` プロパティを表します。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel. CellPropertiesFormat & {
            rowHeight?](/javascript/api/excel/excel.settablerowproperties # 形式)|`format` プロパティを表します。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|`rowHidden` プロパティを表します。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|ワークシートから図形を削除します。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。 DPI は 96 です。 サポートされている形式は、`Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG`、`Excel.PictureFormat.GIF` だけです。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|図形を画像に変換し、base 64 でエンコードされた文字列として画像を返します。 DPI は 96 です。 サポートされている形式は、`Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG`、`Excel.PictureFormat.GIF` だけです。|
||[height](/javascript/api/excel/excel.shape#height)|図形の高さをポイント数で表します。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|指定したポイント数だけ水平方向に図形を移動します。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|z 軸を中心に、指定された度数だけ、図形を時計回りに回転します。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|指定したポイント数だけ垂直方向に図形を移動します。|
||[left](/javascript/api/excel/excel.shape#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shape#name)|図形の名前を表します。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|この図形の結合点の数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shape#fill)|この図形の塗りつぶしの書式設定を返します。 読み取り専用です。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|図形に関連付けられた幾何学的図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[group](/javascript/api/excel/excel.shape#group)|図形に関連付けられた図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[id](/javascript/api/excel/excel.shape#id)|図形 ID を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.shape#image)|図形に関連付けられた画像を返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[level](/javascript/api/excel/excel.shape#level)|指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.shape#line)|図形に関連付けられた線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|この図形の線の書式設定を返します。 読み取り専用です。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|図形がアクティブになったときに発生します。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|図形が非アクティブになると発生します。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|この図形の親グループを表します。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|この図形のテキスト フレーム オブジェクトを返します。 読み取り専用です。|
||[type](/javascript/api/excel/excel.shape#type)|この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。 読み取り専用です。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|図形の回転を角度で表します。|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の高さを変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の高さに対して拡大または縮小されます。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の高さを変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 図以外の図形の場合は、常に現在の高さに対して拡大または縮小されます。|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の幅を変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 図以外の図形の場合は、常に現在の幅に対して拡大または縮小されます。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|指定した係数分だけ図形の幅を変更します。 画像の場合は、図形を元のサイズに対して拡大または縮小するのか、現在のサイズに対して拡大または縮小するのかを指定できます。 画像以外の図形の場合は、常に現在の幅に対して拡大または縮小されます。|
||[set (properties: Excel. Shape)](/javascript/api/excel/excel.shape#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ShapeUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.shape#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|指定された図形をコレクションの z オーダーで上または下に移動します。他の図形の手前または奥に移動します。|
||[top](/javascript/api/excel/excel.shape#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[visible](/javascript/api/excel/excel.shape#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shape#width)|図形の幅 (ポイント数) を表します。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|アクティブ化された図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|図形がアクティブにされたワークシートの ID を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|幾何学的図形をワークシートに追加します。 新しい図形を表す Shape オブジェクトを返します。|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|幾何学的図形をワークシートに追加します。 新しい図形を表す Shape オブジェクトを返します。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|このコレクションのワークシート内の図形のサブセットをグループ化します。 図形の新しいグループを表す Shape オブジェクトを返します。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|base64 エンコード文字列から画像を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|ワークシートに行を追加します。 新しい行を表す Shape オブジェクトを返します。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|ワークシートに行を追加します。 新しい行を表す Shape オブジェクトを返します。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|指定されたテキストを含むテキスト ボックスをワークシートに追加します。 新しいテキスト ボックスを表す Shape オブジェクトを返します。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|ワークシートの図形数を返します。 読み取り専用です。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|名前または ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|コレクション内の位置を使用して図形を取得します。|
||[items](/javascript/api/excel/excel.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[$all](/javascript/api/excel/excel.shapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapecollectionloadoptions#alttextdescription)|コレクション内の各アイテムについて: Shape オブジェクトの別の説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shapecollectionloadoptions#alttexttitle)|コレクション内の各アイテムについて: Shape オブジェクトの代替タイトルテキストを取得または設定します。|
||[connectionSiteCount](/javascript/api/excel/excel.shapecollectionloadoptions#connectionsitecount)|コレクション内の各アイテムについて: この図形の接続サイト数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shapecollectionloadoptions#fill)|コレクション内の各アイテムについて: この図形の塗りつぶしの書式設定を返します。|
||[geometricShape](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshape)|コレクション内の各アイテムについて: 図形に関連付けられたジオメトリック図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[geometricShapeType](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshapetype)|コレクション内の各アイテムについて: この幾何学図形の幾何学的な図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[group](/javascript/api/excel/excel.shapecollectionloadoptions#group)|コレクション内の各アイテムについて: 図形に関連付けられている図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[height](/javascript/api/excel/excel.shapecollectionloadoptions#height)|コレクション内の各項目について: 図形の高さをポイント単位で表します。|
||[id](/javascript/api/excel/excel.shapecollectionloadoptions#id)|コレクション内の各アイテムについて: 図形の識別子を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.shapecollectionloadoptions#image)|コレクション内の各アイテムについて: 図形に関連付けられているイメージを返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[left](/javascript/api/excel/excel.shapecollectionloadoptions#left)|コレクション内の各項目について、図形の左側からワークシートの左端までの距離をポイント単位で指定します。|
||[level](/javascript/api/excel/excel.shapecollectionloadoptions#level)|コレクション内の各項目について: 指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.shapecollectionloadoptions#line)|コレクション内の各アイテムについて: 図形に関連付けられている線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.shapecollectionloadoptions#lineformat)|コレクション内の各アイテムについて: この図形の線の書式設定を返します。|
||[lockAspectRatio](/javascript/api/excel/excel.shapecollectionloadoptions#lockaspectratio)|コレクション内の各アイテムについて: この図形の縦横比がロックされているかどうかを指定します。|
||[name](/javascript/api/excel/excel.shapecollectionloadoptions#name)|コレクション内の各アイテムについて: 図形の名前を表します。|
||[parentGroup](/javascript/api/excel/excel.shapecollectionloadoptions#parentgroup)|コレクション内の各アイテムについて: この図形の親グループを表します。|
||[rotation](/javascript/api/excel/excel.shapecollectionloadoptions#rotation)|コレクション内の各項目の場合: 図形の回転角度を度で表します。|
||[textFrame](/javascript/api/excel/excel.shapecollectionloadoptions#textframe)|コレクション内の各アイテムについて: この図形のテキストフレームオブジェクトを返します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.shapecollectionloadoptions#top)|コレクション内の各項目について、図形の上端からワークシートの上端までの距離をポイント単位で指定します。|
||[type](/javascript/api/excel/excel.shapecollectionloadoptions#type)|コレクション内の各アイテムについて: この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.shapecollectionloadoptions#visible)|コレクション内の各アイテムについて: この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shapecollectionloadoptions#width)|コレクション内の各項目について: 図形の幅をポイント単位で表します。|
||[zOrderPosition](/javascript/api/excel/excel.shapecollectionloadoptions#zorderposition)|コレクション内の各アイテムについて: z オーダーで指定した図形の位置を返します。0は順序スタックの最下部を表します。 読み取り専用です。|
|[図形データ](/javascript/api/excel/excel.shapedata)|[altTextDescription](/javascript/api/excel/excel.shapedata#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shapedata#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[connectionSiteCount](/javascript/api/excel/excel.shapedata#connectionsitecount)|この図形の結合点の数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shapedata#fill)|この図形の塗りつぶしの書式設定を返します。 読み取り専用です。|
||[geometricShapeType](/javascript/api/excel/excel.shapedata#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[height](/javascript/api/excel/excel.shapedata#height)|図形の高さをポイント数で表します。|
||[id](/javascript/api/excel/excel.shapedata#id)|図形 ID を表します。 読み取り専用です。|
||[left](/javascript/api/excel/excel.shapedata#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[level](/javascript/api/excel/excel.shapedata#level)|指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[lineFormat](/javascript/api/excel/excel.shapedata#lineformat)|この図形の線の書式設定を返します。 読み取り専用です。|
||[lockAspectRatio](/javascript/api/excel/excel.shapedata#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shapedata#name)|図形の名前を表します。|
||[rotation](/javascript/api/excel/excel.shapedata#rotation)|図形の回転を角度で表します。|
||[top](/javascript/api/excel/excel.shapedata#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[type](/javascript/api/excel/excel.shapedata#type)|この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.shapedata#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shapedata#width)|図形の幅 (ポイント数) を表します。|
||[zOrderPosition](/javascript/api/excel/excel.shapedata#zorderposition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。 読み取り専用です。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|非アクティブにされた図形の ID を取得します。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|図形が非アクティブにされたワークシートの ID を取得します。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[type](/javascript/api/excel/excel.shapefill#type)|図形の塗りつぶしの種類を返します。 読み取り専用です。 詳細については、Excel.ShapeFillType をご覧ください。|
||[set (プロパティ: エクセル塗りつぶし)](/javascript/api/excel/excel.shapefill#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ShapeFillUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.shapefill#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|図形の塗りつぶしの書式設定を均一な色に設定します。 これにより、塗りつぶしの種類が "Solid" に変更されます。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
|[図形の Filldata](/javascript/api/excel/excel.shapefilldata)|[foregroundColor](/javascript/api/excel/excel.shapefilldata#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[transparency](/javascript/api/excel/excel.shapefilldata#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
||[type](/javascript/api/excel/excel.shapefilldata#type)|図形の塗りつぶしの種類を返します。 読み取り専用です。 詳細については、Excel.ShapeFillType をご覧ください。|
|[図形の Fillloadoptions](/javascript/api/excel/excel.shapefillloadoptions)|[$all](/javascript/api/excel/excel.shapefillloadoptions#$all)||
||[foregroundColor](/javascript/api/excel/excel.shapefillloadoptions#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[transparency](/javascript/api/excel/excel.shapefillloadoptions#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
||[type](/javascript/api/excel/excel.shapefillloadoptions#type)|図形の塗りつぶしの種類を返します。 読み取り専用です。 詳細については、Excel.ShapeFillType をご覧ください。|
|[ShapeFillUpdateData](/javascript/api/excel/excel.shapefillupdatedata)|[foregroundColor](/javascript/api/excel/excel.shapefillupdatedata#foregroundcolor)|図形塗りつぶしの前面色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[transparency](/javascript/api/excel/excel.shapefillupdatedata#transparency)|塗りつぶしの透明度の割合を示す 0.0 (不透明) から 1.0 (透明) までの値を取得または設定します。 図形の種類が透明度をサポートしていない場合、または図形の塗りつぶしが一定ではない場合 (グラデーション塗りつぶしの種類など) は、null を返します。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefont#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefont#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[set (properties: エクセルフォント)](/javascript/api/excel/excel.shapefont#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ShapeFontUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.shapefont#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[size](/javascript/api/excel/excel.shapefont#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeFontData](/javascript/api/excel/excel.shapefontdata)|[bold](/javascript/api/excel/excel.shapefontdata#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefontdata#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefontdata#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefontdata#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[size](/javascript/api/excel/excel.shapefontdata#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefontdata#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeFontLoadOptions](/javascript/api/excel/excel.shapefontloadoptions)|[$all](/javascript/api/excel/excel.shapefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.shapefontloadoptions#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefontloadoptions#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefontloadoptions#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefontloadoptions#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[size](/javascript/api/excel/excel.shapefontloadoptions#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefontloadoptions#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeFontUpdateData](/javascript/api/excel/excel.shapefontupdatedata)|[bold](/javascript/api/excel/excel.shapefontupdatedata#bold)|フォントの太字の状態を表します。 TextRange に太字テキストと太字ではないテキストの両方が含まれている場合、null を返します。|
||[color](/javascript/api/excel/excel.shapefontupdatedata#color)|テキストの色を表す HTML カラー コードです (例: "#FF0000" は赤を表します)。 TextRange に色の異なるテキストが含まれている場合、null を返します。|
||[italic](/javascript/api/excel/excel.shapefontupdatedata#italic)|フォントの斜体の状態を表します。 TextRange に斜体テキストと斜体ではないテキストの両方が含まれている場合、null を返します。|
||[name](/javascript/api/excel/excel.shapefontupdatedata#name)|フォント名 (例: "Calibri") を表します。 テキストが複雑なスクリプトか東アジアの言語の場合、それに対応するフォント名です。それ以外の場合は、ラテン フォントの名前です。|
||[size](/javascript/api/excel/excel.shapefontupdatedata#size)|フォント サイズをポイント単位で表します (11 など)。 TextRange にサイズの異なるテキストが含まれている場合、null を返します。|
||[underline](/javascript/api/excel/excel.shapefontupdatedata#underline)|フォントに適用する下線の種類。 TextRange に下線スタイルの異なるテキストが含まれている場合、null を返します。 詳細については、Excel.ShapeFontUnderlineStyle をご覧ください。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|グループに関連付けられた Shape オブジェクトを返します。 読み取り専用です。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Shape オブジェクトのコレクションを返します。 読み取り専用です。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|指定した図形グループに含まれるグループ化された図形のグループを解除します。|
|[図形 Groupdata](/javascript/api/excel/excel.shapegroupdata)|[id](/javascript/api/excel/excel.shapegroupdata#id)|図形 ID を表します。 読み取り専用です。|
||[shapes](/javascript/api/excel/excel.shapegroupdata#shapes)|Shape オブジェクトのコレクションを返します。 読み取り専用です。|
|[図形/道路/経路オプション](/javascript/api/excel/excel.shapegrouploadoptions)|[$all](/javascript/api/excel/excel.shapegrouploadoptions#$all)||
||[id](/javascript/api/excel/excel.shapegrouploadoptions#id)|図形 ID を表します。 読み取り専用です。|
||[shape](/javascript/api/excel/excel.shapegrouploadoptions#shape)|グループに関連付けられた Shape オブジェクトを返します。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[set (properties: ShapeLineFormat)](/javascript/api/excel/excel.shapelineformat#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: ShapeLineFormatUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.shapelineformat#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[ShapeLineFormatData](/javascript/api/excel/excel.shapelineformatdata)|[color](/javascript/api/excel/excel.shapelineformatdata#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformatdata#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[style](/javascript/api/excel/excel.shapelineformatdata#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformatdata#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformatdata#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformatdata#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[ShapeLineFormatLoadOptions](/javascript/api/excel/excel.shapelineformatloadoptions)|[$all](/javascript/api/excel/excel.shapelineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.shapelineformatloadoptions#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformatloadoptions#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[style](/javascript/api/excel/excel.shapelineformatloadoptions#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformatloadoptions#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformatloadoptions#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformatloadoptions#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[ShapeLineFormatUpdateData](/javascript/api/excel/excel.shapelineformatupdatedata)|[color](/javascript/api/excel/excel.shapelineformatupdatedata#color)|線の色を #RRGGBB 形式の HTML カラー フォーマットで表すか ("FFA500" など)、名前付き HTML カラーで表します ("オレンジ" など)。|
||[dashStyle](/javascript/api/excel/excel.shapelineformatupdatedata#dashstyle)|図形の線スタイルを表します。 線が非表示の場合、または破線のスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[style](/javascript/api/excel/excel.shapelineformatupdatedata#style)|図形の線スタイルを表します。 線が非表示の場合、またはスタイルが一定ではない場合は、null を返します。 詳細については、Excel.ShapeLineStyle をご覧ください。|
||[transparency](/javascript/api/excel/excel.shapelineformatupdatedata#transparency)|指定された線の透明度を示す 0.0 (不透明) から 1.0 (透明) までの値を表します。 図形の透明度が一定ではない場合は、null を返します。|
||[visible](/javascript/api/excel/excel.shapelineformatupdatedata#visible)|図形要素の線の書式設定が表示されるかどうかを表します。 図形の可視性が一定ではない場合は、null を返します。|
||[weight](/javascript/api/excel/excel.shapelineformatupdatedata#weight)|線の太さ (ポイント数) を表します。 線が非表示の場合、または線の太さが一定ではない場合は、null を返します。|
|[図形 Loadoptions](/javascript/api/excel/excel.shapeloadoptions)|[$all](/javascript/api/excel/excel.shapeloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapeloadoptions#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shapeloadoptions#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[connectionSiteCount](/javascript/api/excel/excel.shapeloadoptions#connectionsitecount)|この図形の結合点の数を返します。 読み取り専用です。|
||[fill](/javascript/api/excel/excel.shapeloadoptions#fill)|この図形の塗りつぶしの書式設定を返します。|
||[geometricShape](/javascript/api/excel/excel.shapeloadoptions#geometricshape)|図形に関連付けられた幾何学的図形を返します。 図形の種類が "GeometricShape" ではない場合は、エラーがスローされます。|
||[geometricShapeType](/javascript/api/excel/excel.shapeloadoptions#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[group](/javascript/api/excel/excel.shapeloadoptions#group)|図形に関連付けられた図形グループを返します。 図形の種類が "GroupShape" ではない場合は、エラーがスローされます。|
||[height](/javascript/api/excel/excel.shapeloadoptions#height)|図形の高さをポイント数で表します。|
||[id](/javascript/api/excel/excel.shapeloadoptions#id)|図形 ID を表します。 読み取り専用です。|
||[image](/javascript/api/excel/excel.shapeloadoptions#image)|図形に関連付けられた画像を返します。 図形の種類が "Image" ではない場合は、エラーがスローされます。|
||[left](/javascript/api/excel/excel.shapeloadoptions#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[level](/javascript/api/excel/excel.shapeloadoptions#level)|指定された図形のレベルを表します。 たとえば、レベル 0 は図形がどのグループの一部でもないことを意味し、レベル 1 は図形が最上位グループの一部であることを意味し、レベル 2 は図形が最上位レベルのサブグループの一部であることを意味します。|
||[line](/javascript/api/excel/excel.shapeloadoptions#line)|図形に関連付けられた線を返します。 図形の種類が "Line" ではない場合は、エラーがスローされます。|
||[lineFormat](/javascript/api/excel/excel.shapeloadoptions#lineformat)|この図形の線の書式設定を返します。|
||[lockAspectRatio](/javascript/api/excel/excel.shapeloadoptions#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shapeloadoptions#name)|図形の名前を表します。|
||[parentGroup](/javascript/api/excel/excel.shapeloadoptions#parentgroup)|この図形の親グループを表します。|
||[rotation](/javascript/api/excel/excel.shapeloadoptions#rotation)|図形の回転を角度で表します。|
||[textFrame](/javascript/api/excel/excel.shapeloadoptions#textframe)|この図形のテキスト フレーム オブジェクトを返します。 読み取り専用です。|
||[top](/javascript/api/excel/excel.shapeloadoptions#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[type](/javascript/api/excel/excel.shapeloadoptions#type)|この図形の種類を返します。 詳細については、Excel.ShapeType をご覧ください。 読み取り専用です。|
||[visible](/javascript/api/excel/excel.shapeloadoptions#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shapeloadoptions#width)|図形の幅 (ポイント数) を表します。|
||[zOrderPosition](/javascript/api/excel/excel.shapeloadoptions#zorderposition)|指定された図形の z オーダーでの位置を返します。0 はオーダー スタックの一番下を表します。 読み取り専用です。|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[altTextDescription](/javascript/api/excel/excel.shapeupdatedata#alttextdescription)|Shape オブジェクトの代替説明テキストを取得または設定します。|
||[altTextTitle](/javascript/api/excel/excel.shapeupdatedata#alttexttitle)|Shape オブジェクトの代替タイトル テキストを取得または設定します。|
||[fill](/javascript/api/excel/excel.shapeupdatedata#fill)|この図形の塗りつぶしの書式設定を返します。|
||[geometricShapeType](/javascript/api/excel/excel.shapeupdatedata#geometricshapetype)|この幾何学的図形の種類を表します。 詳細については、Excel.GeometricShapeType をご覧ください。 図形の種類が "GeometricShape" ではない場合は、null を返します。|
||[height](/javascript/api/excel/excel.shapeupdatedata#height)|図形の高さをポイント数で表します。|
||[left](/javascript/api/excel/excel.shapeupdatedata#left)|図形の左側からワークシートの左側までの距離 (ポイント数) です。|
||[lineFormat](/javascript/api/excel/excel.shapeupdatedata#lineformat)|この図形の線の書式設定を返します。|
||[lockAspectRatio](/javascript/api/excel/excel.shapeupdatedata#lockaspectratio)|この図形の縦横比をロックするかどうかを指定します。|
||[name](/javascript/api/excel/excel.shapeupdatedata#name)|図形の名前を表します。|
||[rotation](/javascript/api/excel/excel.shapeupdatedata#rotation)|図形の回転を角度で表します。|
||[top](/javascript/api/excel/excel.shapeupdatedata#top)|図形の上端からワークシートの上までのポイント単位の距離です。|
||[visible](/javascript/api/excel/excel.shapeupdatedata#visible)|この図形の可視性を表します。|
||[width](/javascript/api/excel/excel.shapeupdatedata#width)|図形の幅 (ポイント数) を表します。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|並べ替え基準となるリッチな値のターゲット プロパティ名である下位フィールドを表します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|コレクション内のスタイルの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|コレクション内の位置に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|テーブルの AutoFilter オブジェクトを表します。 読み取り専用です。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|追加されたテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|テーブルが追加されたワークシートの ID を取得します。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|変更の詳細に関する情報を表します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|新しいテーブルがブックに追加されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|指定されたテーブルがブックで削除されたときに発生します。|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.tablecollectionloadoptions#autofilter)|コレクション内の各アイテムについて: テーブルのオートフィルターオブジェクトを表します。|
|[TableData](/javascript/api/excel/excel.tabledata)|[autoFilter](/javascript/api/excel/excel.tabledata#autofilter)|テーブルの AutoFilter オブジェクトを表します。 読み取り専用です。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|イベントのソースを指定します。 詳細については、Excel.EventSource をご覧ください。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|削除されたテーブルの ID を指定します。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|削除されたテーブルの名前を指定します。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|イベントの種類を指定します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|テーブルが削除されたワークシートの ID を指定します。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[autoFilter](/javascript/api/excel/excel.tableloadoptions#autofilter)|テーブルの AutoFilter オブジェクトを表します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|コレクションに含まれるテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|コレクション内の最初のテーブルを取得します。 一番上の左のテーブルがコレクション内で最初のテーブルになるように、コレクションのテーブルが上から下へ、左から右への順で並べ替えられます。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|名前または ID を使用してテーブルを取得します。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Charts Copedcollectionloadoptions](/javascript/api/excel/excel.tablescopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablescopedcollectionloadoptions#$all)||
||[autoFilter](/javascript/api/excel/excel.tablescopedcollectionloadoptions#autofilter)|コレクション内の各アイテムについて: テーブルのオートフィルターオブジェクトを表します。|
||[列](/javascript/api/excel/excel.tablescopedcollectionloadoptions#columns)|コレクション内の各アイテムについて: テーブル内のすべての列のコレクションを表します。|
||[highlightFirstColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightfirstcolumn)|コレクション内の各アイテムについて: 最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightlastcolumn)|コレクション内の各アイテムについて: 最後の列に特別な書式設定が含まれているかどうかを示します。|
||[id](/javascript/api/excel/excel.tablescopedcollectionloadoptions#id)|コレクション内の各アイテムについて: 指定されたブック内のテーブルを一意に識別する値を返します。 識別子の値は、テーブルの名前が変更された場合も変わりません。 読み取り専用です。|
||[legacyId](/javascript/api/excel/excel.tablescopedcollectionloadoptions#legacyid)|コレクション内の各アイテムについて: 数値 id を返します。|
||[name](/javascript/api/excel/excel.tablescopedcollectionloadoptions#name)|コレクション内の各項目について: テーブルの名前。|
||[rows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#rows)|コレクション内の各アイテムについて: テーブル内のすべての行のコレクションを表します。|
||[showBandedColumns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedcolumns)|コレクション内の各アイテムについて: 列に縞模様の書式が表示されているかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedrows)|コレクション内の各アイテムについて: 行に縞模様の書式が設定されているかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showfilterbutton)|コレクション内の各アイテムについて: フィルターボタンを各列ヘッダーの上部に表示するかどうかを示します。 これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
||[showHeaders](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showheaders)|コレクション内の各アイテムについて: ヘッダー行が表示されるかどうかを示します。 この値によって、ヘッダー行の表示または削除を設定できます。|
||[showTotals](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showtotals)|コレクション内の各アイテムについて: [合計] 行が表示されるかどうかを示します。 この値によって、集計行の表示または削除を設定できます。|
||[並べ替え](/javascript/api/excel/excel.tablescopedcollectionloadoptions#sort)|コレクション内の各アイテムについて: テーブルの並べ替えを表します。|
||[style](/javascript/api/excel/excel.tablescopedcollectionloadoptions#style)|コレクション内の各項目について、表のスタイルを表す定数値。 使用可能な値は次のとおりです。 TableStyleLight1 スルー TableStyleLight21、TableStyleMedium1 スルー TableStyleMedium28、TableStyleStyleDark1 スルー TableStyleStyleDark11。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[worksheet](/javascript/api/excel/excel.tablescopedcollectionloadoptions#worksheet)|コレクション内の各アイテムについて: 現在のテーブルを含むワークシート。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|テキスト フレーム内のテキストをすべて削除します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|テキスト フレームにテキストが含まれるかどうかを指定します。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。 詳細については、Excel.TextRange を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[set (properties: Excel. TextFrame)](/javascript/api/excel/excel.textframe#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TextFrameUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.textframe#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[Textframe データ](/javascript/api/excel/excel.textframedata)|[autoSizeSetting](/javascript/api/excel/excel.textframedata#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframedata#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[hasText](/javascript/api/excel/excel.textframedata#hastext)|テキスト フレームにテキストが含まれるかどうかを指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframedata#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframedata#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframedata#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframedata#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframedata#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframedata#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[topMargin](/javascript/api/excel/excel.textframedata#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframedata#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframedata#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[Textframe Loadoptions](/javascript/api/excel/excel.textframeloadoptions)|[$all](/javascript/api/excel/excel.textframeloadoptions#$all)||
||[autoSizeSetting](/javascript/api/excel/excel.textframeloadoptions#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframeloadoptions#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[hasText](/javascript/api/excel/excel.textframeloadoptions#hastext)|テキスト フレームにテキストが含まれるかどうかを指定します。|
||[horizontalAlignment](/javascript/api/excel/excel.textframeloadoptions#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframeloadoptions#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframeloadoptions#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframeloadoptions#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframeloadoptions#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframeloadoptions#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[textRange](/javascript/api/excel/excel.textframeloadoptions#textrange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。 詳細については、Excel.TextRange を参照してください。|
||[topMargin](/javascript/api/excel/excel.textframeloadoptions#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframeloadoptions#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframeloadoptions#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[TextFrameUpdateData](/javascript/api/excel/excel.textframeupdatedata)|[autoSizeSetting](/javascript/api/excel/excel.textframeupdatedata#autosizesetting)|テキスト フレームの自動サイズ変更設定を取得または設定します。 テキストをテキスト フレームに自動的に合わせる、テキスト フレームをテキストに自動的に合わせる、自動サイズ変更を行わない、のいずれかにテキスト フレームを設定できます。|
||[bottomMargin](/javascript/api/excel/excel.textframeupdatedata#bottommargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[horizontalAlignment](/javascript/api/excel/excel.textframeupdatedata#horizontalalignment)|テキスト フレームの水平方向の配置を表します。 詳細については、Excel.ShapeTextHorizontalAlignment を参照してください。|
||[horizontalOverflow](/javascript/api/excel/excel.textframeupdatedata#horizontaloverflow)|テキスト フレームの水平方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextHorizontalOverflow を参照してください。|
||[leftMargin](/javascript/api/excel/excel.textframeupdatedata#leftmargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[orientation](/javascript/api/excel/excel.textframeupdatedata#orientation)|テキスト フレームのテキストの向きを表します。 詳細については、Excel.ShapeTextOrientation を参照してください。|
||[readingOrder](/javascript/api/excel/excel.textframeupdatedata#readingorder)|テキスト フレームの読む方向を表します (左から右または右から左)。 詳細については、Excel.ShapeTextReadingOrder を参照してください。|
||[rightMargin](/javascript/api/excel/excel.textframeupdatedata#rightmargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[topMargin](/javascript/api/excel/excel.textframeupdatedata#topmargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/excel/excel.textframeupdatedata#verticalalignment)|テキスト フレームの垂直方向の配置を表します。 詳細については、Excel.ShapeTextVerticalAlignment を参照してください。|
||[verticalOverflow](/javascript/api/excel/excel.textframeupdatedata#verticaloverflow)|テキスト フレームの垂直方向のオーバーフローの動作を表します。 詳細については、Excel.ShapeTextVerticalOverflow を参照してください。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|指定された範囲の部分文字列に対する TextRange オブジェクトを返します。|
||[font](/javascript/api/excel/excel.textrange#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。 読み取り専用です。|
||[set (プロパティ: Excel. TextRange)](/javascript/api/excel/excel.textrange#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TextRangeUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.textrange#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[text](/javascript/api/excel/excel.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[TextRangeData](/javascript/api/excel/excel.textrangedata)|[font](/javascript/api/excel/excel.textrangedata#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.textrangedata#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[TextRangeLoadOptions](/javascript/api/excel/excel.textrangeloadoptions)|[$all](/javascript/api/excel/excel.textrangeloadoptions#$all)||
||[font](/javascript/api/excel/excel.textrangeloadoptions#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。|
||[text](/javascript/api/excel/excel.textrangeloadoptions#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[TextRangeUpdateData](/javascript/api/excel/excel.textrangeupdatedata)|[font](/javascript/api/excel/excel.textrangeupdatedata#font)|テキスト範囲のフォント属性を表す ShapeFont オブジェクトを返します。|
||[text](/javascript/api/excel/excel.textrangeupdatedata#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、このステートメントを呼び出すと、例外がスローされます。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|ブックで現在アクティブになっているグラフを取得します。 アクティブになっているグラフがない場合、null オブジェクトを返します。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|ブックが複数のユーザーによって編集されている場合 (共同編集)、true となります。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|ブックから現在選択されている 1 つまたは複数の範囲を取得します。 getSelectedRange() の場合と同様に、このメソッドは、選択されているすべての範囲を表す RangeAreas オブジェクトを返します。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|ブックが自動保存モードかどうかを指定します。 読み取り専用です。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Excel 計算エンジンのバージョンとして数字を返します。 読み取り専用です。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|ブックで autoSave の設定が変更されると発生します。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|ブックがローカル環境またはオンライン環境に保存されたかどうかを指定します。 読み取り専用です。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[autoSave](/javascript/api/excel/excel.workbookdata#autosave)|ブックが自動保存モードかどうかを指定します。 読み取り専用です。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookdata#calculationengineversion)|Excel 計算エンジンのバージョンとして数字を返します。 読み取り専用です。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookdata#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[isDirty](/javascript/api/excel/excel.workbookdata#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[previouslySaved](/javascript/api/excel/excel.workbookdata#previouslysaved)|ブックがローカル環境またはオンライン環境に保存されたかどうかを指定します。 読み取り専用です。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookdata#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[autoSave](/javascript/api/excel/excel.workbookloadoptions#autosave)|ブックが自動保存モードかどうかを指定します。 読み取り専用です。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookloadoptions#calculationengineversion)|Excel 計算エンジンのバージョンとして数字を返します。 読み取り専用です。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookloadoptions#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[isDirty](/javascript/api/excel/excel.workbookloadoptions#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[previouslySaved](/javascript/api/excel/excel.workbookloadoptions#previouslysaved)|ブックがローカル環境またはオンライン環境に保存されたかどうかを指定します。 読み取り専用です。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookloadoptions#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[chartDataPointTrack](/javascript/api/excel/excel.workbookupdatedata#chartdatapointtrack)|関連付けられている実際のデータ ポイントをブックの全グラフが追跡している場合、true となります。|
||[isDirty](/javascript/api/excel/excel.workbookupdatedata#isdirty)|ブックが最後に保存された後に変更が行われたかどうかを指定します。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookupdatedata#useprecisionasdisplayed)|ブックを表示桁数でのみ計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|指定された条件に基づいて指定された文字列の発生箇所をすべて見つけ、1 つまたは複数の長方形範囲を構成する RangeAreas オブジェクトとして返します。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|アドレスまたは名前で指定され、1 つまたは複数の長方形範囲ブロックを表す RangeAreas オブジェクトを取得します。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|ワークシートの AutoFilter オブジェクトを表します。 読み取り専用です。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|ワークシートの水平改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|フォーマットが特定のワークシートで変更されたときに発生します。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|ワークシート上のすべての Shape オブジェクトをまとめて返します。 読み取り専用です。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|ワークシートの垂直改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|現在のワークシート内で、指定された条件に基づき、指定された文字列を検索し、置換します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|変更の詳細に関する情報を表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|ブックのワークシートが変更されたときに発生します。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|ブック内のワークシートの書式が変更されたときに発生します。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|ワークシートで選択範囲を変更したときに発生します。|
|[ワークシート Collectionloadoptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetcollectionloadoptions#autofilter)|コレクション内の各アイテムについて: ワークシートのオートフィルターオブジェクトを表します。|
||[enableCalculation](/javascript/api/excel/excel.worksheetcollectionloadoptions#enablecalculation)|コレクション内の各アイテムについて: ワークシートの enableCalculation プロパティを取得または設定します。|
||[pageLayout](/javascript/api/excel/excel.worksheetcollectionloadoptions#pagelayout)|コレクション内の各アイテムについて: ワークシートの PageLayout オブジェクトを取得します。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[autoFilter](/javascript/api/excel/excel.worksheetdata#autofilter)|ワークシートの AutoFilter オブジェクトを表します。 読み取り専用です。|
||[enableCalculation](/javascript/api/excel/excel.worksheetdata#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheetdata#horizontalpagebreaks)|ワークシートの水平改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
||[pageLayout](/javascript/api/excel/excel.worksheetdata#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|
||[shapes](/javascript/api/excel/excel.worksheetdata#shapes)|ワークシート上のすべての Shape オブジェクトをまとめて返します。 読み取り専用です。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheetdata#verticalpagebreaks)|ワークシートの垂直改ページをまとめて取得します。 このコレクションには、手動の改ページのみが含まれます。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|特定のワークシートで変更されたエリアを表す範囲を取得します。 null オブジェクトを返すこともあります。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[ワークシート Loadoptions](/javascript/api/excel/excel.worksheetloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetloadoptions#autofilter)|ワークシートの AutoFilter オブジェクトを表します。|
||[enableCalculation](/javascript/api/excel/excel.worksheetloadoptions#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[pageLayout](/javascript/api/excel/excel.worksheetloadoptions#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|一致方法として完全一致か部分一致を指定します。 完全一致は、セルの内容全体に一致します。 既定値は false (部分一致) です。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|照合の際に大文字と小文字を区別するかどうかを指定します。 既定値は false (区別しない) です。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[enableCalculation](/javascript/api/excel/excel.worksheetupdatedata#enablecalculation)|ワークシートの enableCalculation プロパティを取得または設定します。|
||[pageLayout](/javascript/api/excel/excel.worksheetupdatedata#pagelayout)|ワークシートの PageLayout オブジェクトを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
