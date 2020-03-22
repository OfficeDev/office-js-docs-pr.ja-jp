---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 03/19/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fda0721bd5d7cbec6349c4800a97132d61a26ab9
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891202"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [カルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings-preview) | ブックのカルチャシステム設定 (数値の書式設定など) を取得します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [ブックの挿入](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| ピボットフィルター | ピボットテーブルのフィールドに、値に基づくフィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | ブックを保存して閉じます。  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「[すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|現在のシステムのカルチャ設定に基づく情報を提供します。 これには、カルチャ名、数値形式、およびその他のカルチャに依存する設定が含まれます。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|数値の小数点の記号として使用される文字列を取得します。 これは、Excel のローカル設定に基づいています。|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。 これは、Excel のローカル設定に基づいています。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Microsoft Excel のシステム区切り記号を有効にするかどうかを指定します。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|グラフ軸のタイトルに対して、テキストの方向を指定する角度を表します。 この値は、-90 ~ 90 の整数、または垂直方向のテキストの整数の180のいずれかである必要があります。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフの系列の1つの次元から値を取得します。 指定できるのは、指定された次元と、グラフ系列に対するデータのマッピング方法によって異なります。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツタイプを取得します。|
||[解析](/javascript/api/excel/excel.comment#resolved)|コメントスレッドの状態を取得または設定します。 値 "true" は、スレッドが解決されることを意味します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|応答のコンテンツタイプを取得します。|
||[解析](/javascript/api/excel/excel.commentreply#resolved)|返信の状態を取得または設定します。 値 "true" は、応答が解決された状態であることを意味します。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示するためのカルチャに適した形式を定義します。 これは、現在のシステムのカルチャ設定に基づいています。|
||[name](/javascript/api/excel/excel.cultureinfo#name)|カルチャ名を languagecode2-country/regioncode2 の形式で取得します (例: "zh-cn-cn" または "en-us")。 これは、現在のシステム設定に基づいています。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|数字を表示するためのカルチャに適した形式を定義します。 これは、現在のシステムのカルチャ設定に基づいています。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付の区切り文字として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[短い日付パターン](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|数値の小数点の記号として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[番号 Groupseparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|数値の小数点の左側にある数字のグループを区切るために使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を示します。|
||[排他](/javascript/api/excel/excel.pivotdatefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|`Between`フィルター条件の範囲の下限を指定します。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|`Between`フィルター条件の範囲の上限を指定します。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|、 `Equals` `Before`、 `After`、および`Between`フィルター条件の場合、比較を日単位で行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filter: PivotValueFilter \| pivotvaluefilter \| PivotManualFilter \| pivotvaluefilter \| PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|フィールドの現在の PivotFilters を1つまたは複数設定し、フィールドに適用します。|
||[clearAllFilters ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|すべてのフィールドフィルターのすべての条件をクリアします。 これにより、そのフィールドのアクティブなフィルター処理がすべて削除されます。|
||[clearFilter (filterType: PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|指定した種類のフィールドのフィルターから、すべての既存の条件を削除します (現在適用されている場合)。|
||[getFilters ()](/javascript/api/excel/excel.pivotfield#getfilters--)|フィールドに現在適用されているすべてのフィルターを取得します。|
||[isFiltered (filterType?: PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|フィールドに適用されているフィルターがあるかどうかを確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|ピボットフィールドの現在適用されている日付フィルター。 何も適用されていない場合は、Null を返します。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|ピボットフィールドの現在適用されているラベルフィルター。 何も適用されていない場合は、Null を返します。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|ピボットフィールドの現在適用されている手動フィルター。 何も適用されていない場合は、Null を返します。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|ピボットフィールドの現在適用されている値フィルター。 何も適用されていない場合は、Null を返します。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を示します。|
||[排他](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|フィルター条件間の範囲の下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|`BeginsWith`、 `EndsWith`、および`Contains`フィルター条件で使用される部分文字列。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|フィルター条件の間の範囲の上限を指定します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルです。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理するために選択されたアイテムのリスト。 これらは、選択されたフィールドの既存のアイテムおよび有効なアイテムである必要があります。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[Allow多重 Filtersperfield](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定の PivotField に対して複数の PivotFilters を適用できるかどうかを指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|コレクション内の最初のピボットテーブルを取得します。 コレクション内のピボットテーブルは、上から下、左から右に並べ替えられます。この場合、左上のテーブルはコレクションの最初のピボットテーブルになります。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|名前を使用してピボットテーブルを取得します。 PivotTable が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を示します。|
||[排他](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを*除外*します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|`Between`フィルター条件の範囲の下限を指定します。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターが上位/下位 N 個のアイテム、上位/下位 n%、上位/下位 N の合計であるかどうかを示します。|
||[基準](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上位/下位フィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" 個のしきい値。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|`Between`フィルター条件の範囲の上限を指定します。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理の対象となるフィールドで選択されている "value" の名前です。|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables テーブル (fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|範囲に重なっているピボットテーブルのスコープ設定されたコレクションを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 読み取り専用です。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[番号 Formatcategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。 読み取り専用です。|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシートレベルのカスタムプロパティのコレクションを取得します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|計算を完了した範囲のアドレス。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[ワークシート Customproperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタムプロパティの値を取得します。 読み取り専用です。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタムプロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|イベントがトリガーされた方法を表す変更の種類を取得します。 詳細は「`Excel.RowHiddenChangeType`」をご覧ください。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
