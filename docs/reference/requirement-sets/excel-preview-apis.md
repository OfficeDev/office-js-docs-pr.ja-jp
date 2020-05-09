---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 64c102e2d530d23789597d29edb97796d3f584fd
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170829"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| 日付と時刻の[カルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 日付と時刻の書式に関するその他のカルチャ設定へのアクセスを提供します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [ブックの挿入](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| ピボットフィルター | ピボットテーブルのフィールドに、値に基づくフィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|範囲 spilling | アドインで[動的配列](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)の結果に関連付けられた範囲を検索できます。 | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「[すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフの系列の1つの次元から値を取得します。 指定できるのは、指定された次元と、グラフ系列に対するデータのマッピング方法によって異なります。|
|[コメント](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツタイプを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|返信のコンテンツの種類。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示するためのカルチャに適した形式を定義します。 これは、現在のシステムのカルチャ設定に基づいています。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付の区切り文字として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[短い日付パターン](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotdatefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|`Between`フィルター条件の範囲の下限を指定します。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|`Between`フィルター条件の範囲の上限を指定します。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|、 `Equals` `Before`、 `After`、および`Between`フィルター条件の場合、比較を日単位で行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filter: PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|フィールドの現在の PivotFilters を1つまたは複数設定し、フィールドに適用します。|
||[clearAllFilters ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|すべてのフィールドフィルターのすべての条件をクリアします。 これにより、そのフィールドのアクティブなフィルター処理がすべて削除されます。|
||[clearFilter (filterType: PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|指定した種類のフィールドのフィルターから、すべての既存の条件を削除します (現在適用されている場合)。|
||[getFilters ()](/javascript/api/excel/excel.pivotfield#getfilters--)|フィールドに現在適用されているすべてのフィルターを取得します。|
||[isFiltered (filterType?: PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|フィールドに適用されているフィルターがあるかどうかを確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|ピボットフィールドの現在適用されている日付フィルター。 何も適用されていない場合は、Null を返します。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|ピボットフィールドの現在適用されているラベルフィルター。 何も適用されていない場合は、Null を返します。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|ピボットフィールドの現在適用されている手動フィルター。 何も適用されていない場合は、Null を返します。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|ピボットフィールドの現在適用されている値フィルター。 何も適用されていない場合は、Null を返します。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|フィルター条件間の範囲の下限。|
||[副](/javascript/api/excel/excel.pivotlabelfilter#substring)|`BeginsWith`、 `EndsWith`、および`Contains`フィルター条件で使用される部分文字列。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|フィルター条件の間の範囲の上限を指定します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルです。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理するために選択されたアイテムのリスト。 これらは、選択されたフィールドの既存のアイテムおよび有効なアイテムである必要があります。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[Allow多重 Filtersperfield](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定の PivotField に対して複数の PivotFilters を適用できるかどうかを指定します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|`Between`フィルター条件の範囲の下限を指定します。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターを上位/下位 N 個のアイテム、上位/下位 n%、上位/下位 N 個の合計にするかどうかを指定します。|
||[基準](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上位/下位フィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" 個のしきい値。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|`Between`フィルター条件の範囲の上限を指定します。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理の対象となるフィールドで選択されている "value" の名前です。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[番号 Formatcategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|スライサーに適用されるスタイルです。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|表に適用されるスタイルです。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの id を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの id を取得します。|
|[ブック](/javascript/api/excel/excel.workbook)|[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールドリストウィンドウをブックレベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシートレベルのカスタムプロパティのコレクションを取得します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[ワークシート Customproperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。 カスタムプロパティのキーは大文字と小文字を区別しません。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|指定したキーに対応する新しいカスタムプロパティを追加します。 これにより、既存のカスタムプロパティがそのキーで上書きされます。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタムプロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
