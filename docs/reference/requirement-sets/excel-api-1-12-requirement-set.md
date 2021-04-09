---
title: Excel JavaScript API 要件セット 1.12
description: ExcelApi 1.12 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d66f5797d41c8c07f66fcc8069cd4687cd8d8118
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652217"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Excel JavaScript API 1.12 の新機能

ExcelApi 1.12 では、動的配列を追跡し、数式の直接の前例を見つけ出す API を追加することで、範囲の数式のサポートが増加しました。 ピボットテーブル フィルターの API コントロールも追加しました。 コメント、カルチャ設定、カスタム プロパティの機能領域でも改善が行われた。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメント イベント](../../excel/excel-add-ins-comments.md#comment-events) | コメント コレクションに追加、変更、削除のイベントを追加します。| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 日付と時刻 [のカルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 日付と時刻の書式設定に関する追加の文化設定にアクセスできます。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [アプリケーション](/javascript/api/excel/excel.application) |
| [直接の前例](../../excel/excel-add-ins-ranges-precedents.md) | セルの数式の評価に使用される範囲を返します。| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| ピボット フィルター | ピボットテーブルのフィールドに値駆動型フィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-) [、PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [範囲の流出](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | 動的配列の結果に関連付けられた範囲をアドイン [が検索](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) できます。 | [Range](/javascript/api/excel/excel.range) |
| [ワークシート レベルのカスタム プロパティ](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | ブック レベルのスコープに加えて、ワークシート レベルのカスタム プロパティのスコープを設定できます。 | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット 1.12 の API を示します。 Excel JavaScript API 要件セット 1.12 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)以前の Excel API」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|グラフ軸タイトルのテキストの向きを指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフ系列の 1 つのディメンションから値を取得します。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツ タイプを取得します。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|関連する返信のコメント ID と Id を含む CommentDetail 配列を取得します。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|変更されたイベントのトリガー方法を表す変更の種類を取得します。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|関連する返信のコメント ID と Id を含む CommentDetail 配列を取得します。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|コメントが追加された場合に発生します。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|コメント コレクション内のコメントまたは返信が変更された場合 (返信が削除される場合を含む) に発生します。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|コメント コレクション内のコメントが削除された場合に発生します。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|関連する返信のコメント ID と Id を含む CommentDetail 配列を取得します。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|イベントが発生したワークシートの ID を取得します。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|コメントの ID を表します。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|コメントに属する関連する返信の ID を表します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|返信のコンテンツ タイプ。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示する文化的に適切な形式を定義します。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付区切り記号として使用される文字列を取得します。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付値の書式文字列を取得します。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式文字列を取得します。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[コンパレータ](/javascript/api/excel/excel.pivotdatefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotdatefilter#exclusive)|true の場合、フィルター *は条件を* 満たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|フィルター条件の範囲の下限 `Between` 。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|フィルター条件の範囲の上限 `Between` 。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|`Equals`、、 `Before` `After` およびフィルター条件の場合は、比較を丸 1 日 `Between` として行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|1 つ以上のフィールドの現在のピボットフィルターを設定し、フィールドに適用します。|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|フィールドのすべてのフィルターからすべての条件をクリアします。|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|指定した種類のフィールドのフィルターからすべての既存の条件をクリアします (現在適用されている場合)。|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|フィールドに現在適用されているフィルターを取得します。|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|フィールドに適用されたフィルターが何かあるか確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|PivotField の現在適用されている日付フィルター。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|PivotField の現在適用されているラベル フィルター。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|PivotField の現在適用されている手動フィルター。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|PivotField の現在適用されている値フィルター。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[コンパレータ](/javascript/api/excel/excel.pivotlabelfilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|true の場合、フィルター *は条件を* 満たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Between フィルター条件の範囲の下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|、、およびフィルター条件 `BeginsWith` `EndsWith` に使用される `Contains` サブ文字列。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Between フィルター条件の範囲の上限。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理する選択したアイテムの一覧。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定のピボットフィールドに複数のピボットフィルターを適用できる場合を指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|コレクション内の最初のピボットテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[コンパレータ](/javascript/api/excel/excel.pivotvaluefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|true の場合、フィルター *は条件を* 満たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|フィルター条件の範囲の下限 `Between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターが上位/下位の N 項目、上/下の N パーセント、または上/下の N 合計のフィルターの値を指定します。|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Top/Bottom フィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" しきい値数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|フィルター条件の範囲の上限 `Between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理するフィールドで選択した "value" の名前。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|同じワークシートまたは複数のワークシート内のセルのすべての直接の前例を含む範囲を表す WorkbookRangeAreas オブジェクトを返します。|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|範囲と重なるピボットテーブルのスコープ付きコレクションを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルを配列数式として保存する場合を表します。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|このコレクション内の RangeAreas オブジェクトの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|コレクション内の位置に基づいて RangeAreas オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|コレクション内の `RangeAreas` ワークシート ID または名前に基づいてオブジェクトを返します。|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|コレクション内の `RangeAreas` ワークシート名または id に基づいてオブジェクトを返します。|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|A1 スタイルのアドレスの配列を返します。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|オブジェクトを返 `RangeAreasCollection` します。|
||[範囲](/javascript/api/excel/excel.workbookrangeareas#ranges)|オブジェクト内のこのオブジェクトを構成する範囲を返 `RangeCollection` します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシート レベルのカスタム プロパティのコレクションを取得します。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|指定されたキーにマップする新しいカスタム プロパティを追加します。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
