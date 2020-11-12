---
title: Excel JavaScript API 要件セット1.12
description: ExcelApi 1.12 の要件セットに関する詳細。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ac0085bc504d224bcf56e4cff1f22bbe696bbac8
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996264"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Excel JavaScript API 1.12 の新機能

ExcelApi 1.12 は、動的配列を追跡するための Api を追加し、式の直接の参照元を検索することによって、範囲内の数式のサポートが向上しました。 また、ピボットテーブルフィルターの API コントロールも追加しました。 機能強化は、コメント、カルチャ設定、およびカスタムプロパティ機能領域にも適用されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメントイベント](../../excel/excel-add-ins-comments.md#comment-events) | コメントのコレクションに追加、変更、および削除するイベントを追加します。| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 日付と時刻の [カルチャ設定](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 日付と時刻の書式に関するその他のカルチャ設定へのアクセスを提供します。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [直接の参照元](../../excel/excel-add-ins-ranges-advanced.md#get-formula-precedents) | セルの数式の評価に使用される範囲を返します。| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| ピボットフィルター | ピボットテーブルのフィールドに、値に基づくフィルターを適用します。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [範囲 spilling](../../excel/excel-add-ins-ranges-advanced.md#handle-dynamic-arrays-and-spilling) | アドインで [動的配列](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) の結果に関連付けられた範囲を検索できます。 | [Range](/javascript/api/excel/excel.range) |
| [ワークシートレベルのカスタムプロパティ](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | ブックレベルにスコープが設定されているだけでなく、カスタムプロパティをワークシートレベルでスコープ設定することができます。 | [ワークシート Customproperty](/javascript/api/excel/excel.worksheetcustomproperty)、 [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.12 の Api を示します。 Excel JavaScript API 要件セット1.12 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.12 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|グラフ軸のタイトルに対して、テキストの方向を指定する角度を指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフの系列の1つの次元から値を取得します。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|コメントのコンテンツタイプを取得します。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[コメントの詳細](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|コメントの Id とそれに関連する返信の Id が含まれているコメントの詳細配列を取得します。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|イベントが発生したワークシートの Id を取得します。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|変更されたイベントの発生方法を表す変更の種類を取得します。|
||[コメントの詳細](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|コメントの Id とそれに関連する返信の Id が含まれているコメントの詳細配列を取得します。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|イベントが発生したワークシートの Id を取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|コメントが追加されるときに発生します。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|返信が削除されたときを含め、コメントのコレクション内のコメントまたは返信が変更されたときに発生します。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|コメントのコレクションでコメントが削除されると発生します。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[コメントの詳細](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|コメントの Id とそれに関連する返信の Id が含まれているコメントの詳細配列を取得します。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|イベントが発生したワークシートの Id を取得します。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|コメントの id を表します。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|コメントに関連付けられている返信の id を表します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|返信のコンテンツの種類。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|日付と時刻を表示するためのカルチャに適した形式を定義します。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|日付の区切り文字として使用される文字列を取得します。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|長い日付の値の書式指定文字列を取得します。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|長い時間の値の書式指定文字列を取得します。|
||[短い日付パターン](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|短い日付の値の書式文字列を取得します。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|時刻の区切り記号として使用される文字列を取得します。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|比較演算子は、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotdatefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|フィルター条件の範囲の下限を指定し `Between` ます。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|フィルター条件の範囲の上限を指定し `Between` ます。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|`Equals`、、 `Before` `After` 、および `Between` フィルター条件の場合、比較を日単位で行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filter: PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|1つまたは複数のフィールドの現在の PivotFilters を設定し、フィールドに適用します。|
||[clearAllFilters ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|すべてのフィールドフィルターのすべての条件をクリアします。|
||[clearFilter (filterType: PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|指定した種類のフィールドのフィルターから、すべての既存の条件を削除します (現在適用されている場合)。|
||[getFilters ()](/javascript/api/excel/excel.pivotfield#getfilters--)|フィールドに現在適用されているすべてのフィルターを取得します。|
||[isFiltered (filterType?: PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|フィールドに適用されているフィルターがあるかどうかを確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|ピボットフィールドの現在適用されている日付フィルター。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|ピボットフィールドの現在適用されているラベルフィルター。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|ピボットフィールドの現在適用されている手動フィルター。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|ピボットフィールドの現在適用されている値フィルター。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比較演算子は、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|フィルター条件間の範囲の下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|`BeginsWith`、、 `EndsWith` およびフィルター条件で使用される部分文字列 `Contains` 。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|フィルター条件の間の範囲の上限を指定します。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|手動でフィルター処理するために選択されたアイテムのリスト。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[Allow多重 Filtersperfield](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|ピボットテーブルで、テーブル内の特定の PivotField に対して複数の PivotFilters を適用できるかどうかを指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|コレクション内の最初のピボットテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比較演算子は、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|True の場合、フィルターは条件に一致するアイテムを *除外* します。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|フィルター条件の範囲の下限を指定し `Between` ます。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|フィルターを上位/下位 N 個のアイテム、上位/下位 n%、上位/下位 N 個の合計にするかどうかを指定します。|
||[基準](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上位/下位フィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" 個のしきい値。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|フィルター条件の範囲の上限を指定し `Between` ます。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理の対象となるフィールドで選択されている "value" の名前です。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents 元 ()](/javascript/api/excel/excel.range#getdirectprecedents--)|同じワークシートまたは複数のワークシート内のセルのすべての直接の参照元を含む範囲を表す WorkbookRangeAreas オブジェクトを返します。|
||[getPivotTables テーブル (fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|範囲に重なっているピボットテーブルのスコープ設定されたコレクションを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[番号 Formatcategories](/javascript/api/excel/excel.range#numberformatcategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|このコレクション内の RangeAreas オブジェクトの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|コレクション内の位置に基づいて RangeAreas オブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|`RangeAreas`ワークシート id またはコレクション内の名前に基づいてオブジェクトを返します。|
||[getRangeAreasOrNullObjectBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|`RangeAreas`ワークシートの名前またはコレクション内の id に基づいてオブジェクトを返します。|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|A1 形式のアドレスの配列を返します。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|オブジェクトを返し `RangeAreasCollection` ます。|
||[域](/javascript/api/excel/excel.workbookrangeareas#ranges)|オブジェクト内のこのオブジェクトを構成する範囲を返し `RangeCollection` ます。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|ワークシートレベルのカスタムプロパティのコレクションを取得します。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|指定したキーに対応する新しいカスタムプロパティを追加します。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|このワークシートのカスタムプロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
