---
title: ExcelJavaScript API 要件セット 1.12
description: ExcelApi 1.12 要件セットの詳細。
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 10587b84ba476b91cdd56d8472e551348b3a718b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938732"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>JavaScript API 1.12 Excel新機能

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

次の表に、JavaScript API 要件セット 1.12 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.12 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textOrientation)|グラフ軸タイトルのテキストの向きを指定します。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getDimensionValues_dimension_)|グラフ系列の 1 つのディメンションから値を取得します。|
|[コメント](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contentType)|コメントのコンテンツ タイプを取得します。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentDetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetId)|イベントが発生したワークシートの ID を取得します。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changeType)|変更されたイベントのトリガー方法を表す変更の種類を取得します。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentDetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetId)|イベントが発生したワークシートの ID を取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onAdded)|コメントが追加された場合に発生します。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onChanged)|コメント コレクション内のコメントまたは返信が変更された場合 (返信が削除される場合を含む) に発生します。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#onDeleted)|コメント コレクション内のコメントが削除された場合に発生します。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentDetails)|関連する `CommentDetail` 返信のコメント ID と ID を含む配列を取得します。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|イベントのソースを指定します。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetId)|イベントが発生したワークシートの ID を取得します。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentId)|コメントの ID を表します。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyIds)|コメントに属する関連する返信の ID を表します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contentType)|返信のコンテンツ タイプ。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeFormat)|日付と時刻を表示する文化的に適切な形式を定義します。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateSeparator)|日付区切り記号として使用される文字列を取得します。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longDatePattern)|長い日付値の書式文字列を取得します。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longTimePattern)|長い時間の値の書式文字列を取得します。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortDatePattern)|短い日付の値の書式文字列を取得します。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeSeparator)|時刻の区切り記号として使用される文字列を取得します。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[コンパレータ](/javascript/api/excel/excel.pivotdatefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotdatefilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerBound)|フィルター条件の範囲の下限 `between` 。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperBound)|フィルター条件の範囲の上限 `between` 。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholeDays)|`equals`、、 `before` `after` およびフィルター条件の場合は、比較を丸 1 日 `between` として行う必要があるかどうかを示します。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyFilter_filter_)|1 つ以上のフィールドの現在のピボットフィルターを設定し、フィールドに適用します。|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearAllFilters__)|フィールドのすべてのフィルターからすべての条件をクリアします。|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearFilter_filterType_)|指定した種類のフィールドのフィルターからすべての既存の条件をクリアします (現在適用されている場合)。|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getFilters__)|フィールドに現在適用されているフィルターを取得します。|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isFiltered_filterType_)|フィールドに適用されたフィルターが何かあるか確認します。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#dateFilter)|PivotField の現在適用されている日付フィルター。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelFilter)|PivotField の現在適用されているラベル フィルター。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualFilter)|PivotField の現在適用されている手動フィルター。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valueFilter)|PivotField の現在適用されている値フィルター。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[コンパレータ](/javascript/api/excel/excel.pivotlabelfilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerBound)|フィルター条件の範囲の下限 `between` 。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|、、およびフィルター条件 `beginsWith` `endsWith` に使用される `contains` サブ文字列。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperBound)|フィルター条件の範囲の上限 `between` 。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selectedItems)|手動でフィルター処理する選択したアイテムの一覧。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowMultipleFiltersPerField)|ピボットテーブルで、テーブル内の特定のピボットフィールドに複数のピボットフィルターを適用できる場合を指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getCount__)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getFirst__)|コレクション内の最初のピボットテーブルを取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getItem_key_)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getItemOrNullObject_name_)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[コンパレータ](/javascript/api/excel/excel.pivotvaluefilter#comparator)|コンパレータは、他の値を比較する静的な値です。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|必要なフィルター条件を定義するフィルターの条件を指定します。|
||[排他的](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|場合 `true` 、フィルター *は条件を満* たすアイテムを除外します。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerBound)|フィルター条件の範囲の下限 `between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectionType)|フィルターが上位/下位の N 項目、上/下の N パーセント、または上/下の N 合計のフィルターの値を指定します。|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|上/下のフィルター条件に対してフィルター処理するアイテム、パーセント、または合計の "N" しきい値数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperBound)|フィルター条件の範囲の上限 `between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|フィルター処理するフィールドで選択した "value" の名前。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getDirectPrecedents__)|同じワークシートまたは複数のワークシート内のセルのすべての直接の前例を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getPivotTables_fullyContained_)|範囲と重なるピボットテーブルのスコープ付きコレクションを取得します。|
||[getSpillParent()](/javascript/api/excel/excel.range#getSpillParent__)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getSpillParentOrNullObject__)|セルが流出するアンカー セルを含む範囲オブジェクトを取得します。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getSpillingToRange__)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getSpillingToRangeOrNullObject__)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。|
||[hasSpill](/javascript/api/excel/excel.range#hasSpill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberFormatCategories)|各セルの数値形式のカテゴリを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedAsArray)|すべてのセルが配列数式として保存される場合を表します。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getCount__)|このコレクション内の `RangeAreas` オブジェクトの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getItemAt_index_)|コレクション内の `RangeAreas` 位置に基づいてオブジェクトを返します。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasBySheet_key_)|コレクション内の `RangeAreas` ワークシート ID または名前に基づいてオブジェクトを返します。|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasOrNullObjectBySheet_key_)|コレクション内の `RangeAreas` ワークシート名または ID に基づいてオブジェクトを返します。|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|A1 スタイルのアドレスの配列を返します。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|オブジェクトを返 `RangeAreasCollection` します。|
||[範囲](/javascript/api/excel/excel.workbookrangeareas#ranges)|オブジェクト内のこのオブジェクトを構成する範囲を返 `RangeCollection` します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customProperties)|ワークシート レベルのカスタム プロパティのコレクションを取得します。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete__)|カスタム プロパティを削除します。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|カスタム プロパティのキーを取得します。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|カスタム プロパティの値を取得または設定します。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add_key__value_)|指定されたキーにマップする新しいカスタム プロパティを追加します。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getCount__)|このワークシートのカスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItem_key_)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItemOrNullObject_key_)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
