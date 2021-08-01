---
title: Excel JavaScript プレビュー API
description: JavaScript API のExcel詳細。
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5de8ee52aea357c8dce4d2027556e5e8a5b1a4ac
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671717"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

次の表に、API の簡潔な概要を示しますが、後続の API リスト [テーブル](#api-list) には詳細な一覧が示されています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| グラフ データ テーブル | グラフ上のデータ テーブルの外観、書式設定、および表示を制御します。 | [Chart](/javascript/api/excel/excel.chart)、 [ChartDataTable](/javascript/api/excel/excel.chartdatatable)、 [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| ドキュメント タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| ID | 表示名や電子メール アドレスなど、ユーザー ID を管理します。 | [Identity](/javascript/api/excel/excel.identity)、 [IdentityCollection](/javascript/api/excel/excel.identitycollection)、 [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| リンクされたデータ型 | 外部ソースからデータに接続されたデータExcelサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| リンクされたブック | ブック間のリンクを管理します。ブックリンクの更新と破損のサポートを含む。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)、 [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| テーブルのスタイル | フォント、罫線、塗りつぶしの色、および表のスタイルの他の側面のコントロールを提供します。 | [Table](/javascript/api/excel/excel.table)、 [PivotTable](/javascript/api/excel/excel.pivottable)、 [Slicer](/javascript/api/excel/excel.slicer) |
| クエリ | 名前、更新日、クエリ数のようなクエリ属性を取得します。 | [Query](/javascript/api/excel/excel.query)、 [QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中Excel JavaScript API の一覧を示します。 すべての JavaScript API (プレビュー API Excel以前にリリースされた API を含む) の完全な一覧については[、JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)API Excel参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|セルまたはセルが削除された場合に残りのセルが移動する方向 (上または左など) を表します。|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|新しいセルまたはセルを挿入するときに既存のセルが移動する方向 (下方向や右方向など) を表します。|
|[グラフ](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|グラフのデータ テーブルを取得します。|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|グラフのデータ テーブルを取得します。|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|塗りつぶし、フォント、罫線の形式を含むグラフ データ テーブルの形式を表します。|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|データ テーブルの水平方向の罫線を表示するかどうかを指定します。|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|データ テーブルの凡例キーを表示するかどうかを指定します。|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|データ テーブルの輪郭線を表示するかどうかを指定します。|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|データ テーブルの垂直罫線を表示するかどうかを指定します。|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|グラフのデータ テーブルを表示するかどうかを指定します。|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|グラフ データ テーブルの罫線の形式 (色、線のスタイル、太さ) を表します。|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|背景の書式設定情報を含む、オブジェクトの塗りつぶしの書式を表します。|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|現在のオブジェクトのフォント属性 (フォント名、フォント サイズ、色など) を表します。|
|[コメント](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|コメントに添付されたタスクを、割り当て先として指定されたユーザーに割り当てる。|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|このコメントに関連付けられているタスクを取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|ID に基づいてコレクションからコメントを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|その ID で識別されるコメント返信を返します。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|ID で識別される条件付き書式を返します。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|タスクの完了率を指定します。|
||[優先度](/javascript/api/excel/excel.documenttask#priority)|タスクの優先度を指定します。|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|タスクの割り当て人のコレクションを返します。|
||[変更点](/javascript/api/excel/excel.documenttask#changes)|タスクの変更レコードを取得します。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|タスクに関連付けられたコメントを取得します。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|タスクを完了した最新のユーザーを取得します。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|タスクが完了した日時を取得します。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|タスクを作成したユーザーを取得します。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|タスクが作成された日時を取得します。|
||[id](/javascript/api/excel/excel.documenttask#id)|タスクの ID を取得します。|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|タスクの開始日と期日を変更します。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|タスクを開始する日付と時刻を取得または設定します。期限が設定されます。|
||[title](/javascript/api/excel/excel.documenttask#title)|タスクのタイトルを指定します。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[割り当て先](/javascript/api/excel/excel.documenttaskchange#assignee)|変更レコードの種類のタスクに割り当てられたユーザー、または変更レコードの種類のタスクから割り当てられていないユーザー `assign` `unassign` を表します。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|タスクを作成または変更したユーザーを表します。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|タスクの変更をアンカー `Comment` する ID `CommentReply` を表します。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|タスク変更レコードの作成日時を表します。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|タスクの期日と時刻を UTC タイム ゾーンで表します。|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|タスク変更レコードの ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|タスクの完了率を表します。|
||[優先度](/javascript/api/excel/excel.documenttaskchange#priority)|タスクの優先度を表します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|タスクの開始日時を UTC タイム ゾーンで表します。|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|タスクのタイトルを表します。|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|タスク変更レコードのアクションの種類を表します。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|変更レコードの `DocumentTaskChange.id` 種類に対して元に戻されたプロパティ `undo` を表します。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|タスクのコレクション内の変更レコードの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|コレクション内のインデックスを使用してタスク変更レコードを取得します。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|コレクション内のタスクの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|ID を使用してタスクを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|コレクション内のインデックスによってタスクを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|ID を使用してタスクを取得します。|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|タスクが期限の日時を取得します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|タスクを開始する日付と時刻を取得します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|名前または ID を使用して図形を取得します。|
|[ID](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.identity#email)|ユーザーの電子メール アドレスを表します。|
||[id](/javascript/api/excel/excel.identity#id)|ユーザーの一意の ID を表します。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add_assignee_)|コレクションにユーザー ID を追加します。|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|コレクションからすべてのユーザー ID を削除します。|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|コレクション内のアイテムの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|コレクション内のインデックスを使用してドキュメント ユーザー ID を取得します。|
||[items](/javascript/api/excel/excel.identitycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove_assignee_)|コレクションからユーザー ID を削除します。|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.identityentity#email)|ユーザーの電子メール アドレスを表します。|
||[id](/javascript/api/excel/excel.identityentity#id)|ユーザーの一意の ID を表します。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|リンクされたデータ型のデータ プロバイダーの名前。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|リンクされたデータ型が最後に更新されたときにブックが開か以降のローカルタイム ゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|リンクされたデータ型が "定期的" に設定されている場合に更新される頻度 (秒 `refreshMode` )。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|リンクされたデータ型のデータを取得するメカニズム。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|リンクされたデータ型の一意の ID。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|リンクされたデータ型でサポートされているすべての更新モードを持つ配列を返します。|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|新しいリンクされたデータ型の一意の ID。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|サービス ID 別にリンクされたデータ型を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|コレクション内のインデックスによってリンクされたデータ型を取得します。|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|ID によってリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|リンクされたブックを指すリンクを壊す要求を行います。|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|リンクされたブックを指す元の URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|リンクされたブックから取得したデータを更新する要求を行います。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|リンクされたブックへのすべてのリンクを壊します。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|リンクされたブックに関する情報を URL で取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|リンクされたブックに関する情報を URL で取得します。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|すべてのブック リンクを更新する要求を行います。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|ブック リンクの更新モードを表します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|名前を使用してシート ビューを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|ピボットテーブルに適用されるスタイル。|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|コレクション内の最初のピボットテーブルを取得します。|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|クエリが最後に更新された場合のクエリ エラー メッセージを取得します。|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|オブジェクトの種類を '読み込まれた' クエリを取得します。|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|データ モデルに読み込まれたクエリを指定します。|
||[name](/javascript/api/excel/excel.query#name)|クエリの名前を取得します。|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|クエリが最後に更新された日時を取得します。|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|クエリが最後に更新されたときに読み込まれた行の数を取得します。|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|ブック内のクエリの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|コレクションの名前に基づいてクエリを取得します。|
||[items](/javascript/api/excel/excel.querycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|同じワークシートまたは複数のワークシート内のセルのすべての従属セルを含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|同じワークシートまたは複数のワークシート内のセルのすべての前例を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|更新モードが変更されたオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が成功したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|更新要求が完了したオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求から生成された警告を含む配列。|
|[図形](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|図形の表示名を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|名前または ID を使用して図形を取得します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|スライサーに適用されるスタイル。|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|スライサーに適用されるスタイルを設定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|名前に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|特定のテーブルにフィルターが適用されると発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|テーブルに適用されるスタイル。|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|テーブルに適用されるスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|ブックまたはワークシート内の任意のテーブルにフィルターが適用されると発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|フィルターが適用されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|テーブルを含むワークシートの ID を取得します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|テーブルから複数の行を削除します。|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|指定したインデックスから、指定した数の行をテーブルから削除します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|名前または ID でテーブルを取得します。|
|[ブック](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|ブックの一部であるリンクされたデータ型のコレクションを返します。|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|リンクされたブックのコレクションを返します。|
||[クエリ](/javascript/api/excel/excel.workbook#queries)|ブックの一部である Power Query クエリのコレクションを返します。|
||[タスク](/javascript/api/excel/excel.workbook#tasks)|ブックに存在するタスクのコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|ピボットテーブルのフィールド 一覧ウィンドウをブック レベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|特定のワークシートにフィルターが適用されると発生します。|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|ワークシートの保護状態が変更された場合に発生します。|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Open ファイルの XML で読み取り可能なこのワークシートを表すOfficeします。|
||[タスク](/javascript/api/excel/excel.worksheet#tasks)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|セルまたはセルを削除または挿入するときに、ワークシート内のセルが移動する方向への変更を表します。|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|イベントのトリガー ソースを表します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|ワークシートの保護状態が変更された場合に発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|フィルターが適用されるワークシートの ID を取得します。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|ワークシートの現在の保護状態を取得します。|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|保護状態が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
