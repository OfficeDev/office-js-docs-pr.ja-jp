---
title: Excel JavaScript API 要件セット 1.14
description: ExcelApi 1.14 要件セットの詳細。
ms.date: 12/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 97c221bb8530d32fbfe42ae82519c942c6499928
ms.sourcegitcommit: e4d7791cefb29498a8bffce626a6218cee06abd9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/03/2021
ms.locfileid: "61285000"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>JavaScript API 1.14 Excel新機能

ExcelApi 1.14 には、グラフのデータ テーブル機能を制御するオブジェクト、数式のすべての先行セルを検索するメソッド、ワークシートの保護状態の変更を追跡するワークシート保護イベントが追加されました。 また、、 、、など、オブジェクトの複数のメソッドを追加 [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) `CommentCollection` `ShapeCollection` し、エラー `StyleCollection` 処理を改善しました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [グラフ データ テーブル](../../excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | グラフ上のデータ テーブルの外観、書式設定、および表示を制御します。 | [Chart](/javascript/api/excel/excel.chart)、 [ChartDataTable](/javascript/api/excel/excel.chartdatatable)、 [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [数式の前例](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | 数式のすべての前のセルを返します。 | [Range](/javascript/api/excel/excel.range) |
| クエリ | 名前、更新日、クエリ数など、Power Query 属性を取得します。 | [Query](/javascript/api/excel/excel.query)、 [QueryCollection](/javascript/api/excel/excel.querycollection)|
| ワークシート保護イベント | ワークシートの保護状態に対する変更と、それらの変更のソースを追跡します。 | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.14 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.14 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|オートフィルターの列フィルター条件をクリアします。|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|セルまたはセルが削除された場合に残りのセルが移動する方向 (上または左など) を表します。|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|新しいセルまたはセルを挿入するときに既存のセルが移動する方向 (下方向や右方向など) を表します。|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|グラフのデータ テーブルを取得します。|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|グラフのデータ テーブルを取得します。|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|塗りつぶし、フォント、罫線の形式を含むグラフ データ テーブルの形式を表します。|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|データ テーブルの水平方向の罫線を表示するかどうかを指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|データ テーブルの凡例キーを表示するかどうかを指定します。|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|データ テーブルの輪郭線を表示するかどうかを指定します。|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|データ テーブルの垂直罫線を表示するかどうかを指定します。|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|グラフのデータ テーブルを表示するかどうかを指定します。|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[罫線](/javascript/api/excel/excel.chartdatatableformat#border)|グラフ データ テーブルの罫線の形式 (色、線のスタイル、太さ) を表します。|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|ID に基づいてコレクションからコメントを取得します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|その ID で識別されるコメント返信を返します。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|ID で識別される条件付き書式を返します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|名前または ID を使用して図形を取得します。|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|クエリが最後に更新された場合のクエリ エラー メッセージを取得します。|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|オブジェクトの種類に読み込まれたクエリを取得します。|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|データ モデルに読み込まれたクエリを指定します。|
||[name](/javascript/api/excel/excel.query#name)|クエリの名前を取得します。|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|クエリが最後に更新された日時を取得します。|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|クエリが最後に更新されたときに読み込まれた行の数を取得します。|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|ブック内のクエリの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|コレクションの名前に基づいてクエリを取得します。|
||[items](/javascript/api/excel/excel.querycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|同じワークシートまたは複数のワークシート内のセルのすべての前例を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|名前または ID を使用して図形を取得します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|名前に基づいてスタイルを取得します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|名前または ID でテーブルを取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[クエリ](/javascript/api/excel/excel.workbook#queries)|ブックの一部である Power Query クエリのコレクションを返します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|ワークシートの保護状態が変更された場合に発生します。|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Open ファイルの XML で読み取り可能なこのワークシートを表すOfficeします。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|セルまたはセルを削除または挿入するときに、ワークシート内のセルが移動する方向への変更を表します。|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|イベントのトリガー ソースを表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|ワークシートの保護状態が変更された場合に発生します。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|ワークシートの現在の保護状態を取得します。|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|保護状態が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
