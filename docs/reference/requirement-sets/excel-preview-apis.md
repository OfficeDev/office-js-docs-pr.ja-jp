---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 06/29/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d1701ad393b96e33f0007bfcb5609c93c13608a2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430766"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| 日付と時刻の [カルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 日付と時刻の書式に関するその他のカルチャ設定へのアクセスを提供します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [ブックの挿入](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| ピボットフィルター | ピボットテーブルのフィールドに、値に基づくフィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|範囲 spilling | アドインで [動的配列](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) の結果に関連付けられた範囲を検索できます。 | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「 [すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフの系列の1つの次元から値を取得します。 指定できるのは、指定された次元と、グラフ系列に対するデータのマッピング方法によって異なります。|
|[コメント](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツタイプを取得します。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[コメントの詳細](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|`CommentDetail`関連付けられている返信のコメント id と id を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|イベントのソースを指定します。 詳細は「`Excel.EventSource`」をご覧ください。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|イベントの種類を取得します。 詳細は「`Excel.EventType`」をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|変更されたイベントの発生方法を表す変更の種類を取得します。|
||[コメントの詳細](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|`CommentDetail`関連する返信のコメント id と id を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|イベントのソースを指定します。 詳細は「`Excel.EventSource`」をご覧ください。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|イベントの種類を取得します。 詳細は「`Excel.EventType`」をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|コメントが追加されるときに発生します。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|返信が削除されたときを含め、コメントのコレクション内のコメントまたは返信が変更されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|コメントのコレクションでコメントが削除されると発生します。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[コメントの詳細](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|`CommentDetail`関連付けられている返信のコメント id と id を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|イベントのソースを指定します。 詳細は「`Excel.EventSource`」をご覧ください。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|イベントの種類を取得します。 詳細は「`Excel.EventType`」をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[コメントの詳細](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|コメントの ID を表します。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|コメントに関連付けられている返信の Id を表します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|返信のコンテンツの種類。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示するためのカルチャに適した形式を定義します。 これは、現在のシステムのカルチャ設定に基づいています。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付の区切り文字として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式指定文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[短い日付パターン](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。 これは、現在のシステム設定に基づいています。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。 これは、現在のシステム設定に基づいています。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシートビューをアクティブにします。 これは、Excel UI の [切り替え先] を使用するのと同じです。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシートビューを削除します。|
||[重複 (名前?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシートビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シートビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定した名前の新しいシートビューを作成します。|
||[enterTemporary ()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シートビューを作成してアクティブにします。|
||[exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシートビューを終了します。|
||[getActive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシートビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシートビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシートビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスによってシートビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotdatefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|フィルター条件の範囲の下限を指定し `Between` ます。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|フィルター条件の範囲の上限を指定し `Between` ます。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|`Equals`、、 `Before` `After` 、および `Between` フィルター条件の場合、比較を日単位で行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filter: PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|1つまたは複数のフィールドの現在の PivotFilters を設定し、フィールドに適用します。|
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
||[排他](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|フィルター条件間の範囲の下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|`BeginsWith`、、 `EndsWith` およびフィルター条件で使用される部分文字列 `Contains` 。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|フィルター条件の間の範囲の上限を指定します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルです。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理するために選択されたアイテムのリスト。 これらは、選択されたフィールドの既存のアイテムおよび有効なアイテムである必要があります。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[Allow多重 Filtersperfield](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定の PivotField に対して複数の PivotFilters を適用できるかどうかを指定します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比較演算子は、他の値を比較する静的な値です。 比較の種類は、条件によって定義されます。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。 既定では false (条件に一致するアイテムを含むフィルター)。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|フィルター条件の範囲の下限を指定し `Between` ます。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターを上位/下位 N 個のアイテム、上位/下位 n%、上位/下位 N 個の合計にするかどうかを指定します。|
||[基準](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上位/下位フィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" 個のしきい値。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|フィルター条件の範囲の上限を指定し `Between` ます。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理の対象となるフィールドで選択されている "value" の名前です。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents 元 ()](/javascript/api/excel/excel.range#getdirectprecedents--)|`WorkbookRangeAreas`同じワークシートまたは複数のワークシート内のセルのすべての直接の参照元を含む範囲を表すオブジェクト型 (object) の値を取得します。|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|この範囲内の結合領域を表す RangeAreas オブジェクトを返します。 この範囲内のマージされた領域の数が512を超える場合、API は結果を返すことに失敗します。|
||[getPrecedents 元 ()](/javascript/api/excel/excel.range#getprecedents--)|`WorkbookRangeAreas`同じワークシートまたは複数のワークシート内のセルのすべての参照元を含む範囲を表すオブジェクト型 (object) の値を取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[番号 Formatcategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|このコレクション内の RangeAreas オブジェクトの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|コレクション内の位置に基づいて RangeAreas オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
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
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|`RangeAreas`ワークシート id またはコレクション内の名前に基づいてオブジェクトを返します。|
||[getRangeAreasOrNullObjectBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|`RangeAreas`ワークシートの名前またはコレクション内の id に基づいてオブジェクトを返します。 ワークシートが存在しない場合は null オブジェクトを返します。|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|A1 形式のアドレスの配列を返します。 Address 値には、セルの各長方形ブロックのワークシート名が格納されます (例: "Sheet1!A1: B4、Sheet1!D1: D4 ")。 読み取り専用です。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|RangeAreasCollection オブジェクトを取得します。コレクション内の各 RangeAreas は、1つのワークシート内の1つまたは複数の四角形の範囲を表します。|
||[域](/javascript/api/excel/excel.workbookrangeareas#ranges)|このオブジェクトを構成する範囲のコレクションを返します。|
|[ワークシート](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシートレベルのカスタムプロパティのコレクションを取得します。|
||[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートにあるシートビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[ワークシート Customproperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。 カスタムプロパティのキーは大文字と小文字を区別しません。 キーは255文字に制限されています (大きい値を指定すると、"InvalidArgument" エラーがスローされます)。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|指定したキーに対応する新しいカスタムプロパティを追加します。 これにより、既存のカスタムプロパティがそのキーで上書きされます。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタムプロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
