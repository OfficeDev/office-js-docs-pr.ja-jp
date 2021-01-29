---
title: Excel JavaScript プレビュー API
description: 予定されている Excel JavaScript API の詳細。
ms.date: 01/26/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 10057123cc159af0c00a6b6e6345d8f6ab316822
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043898"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| リンクされたデータ型 | 外部ソースから Excel に接続されるデータ型のサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [タスク](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript API を示します。 すべての Excel JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧については、 [すべての Excel JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|このコメントに関連付けられているタスクを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|このコメントに関連付けられているタスクを取得します。|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|変更された数式を含むセルのアドレス。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|変更前の数式を表します。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|リンクされたデータ型のデータ プロバイダーの名前。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|リンクされたデータ型が最後に更新されたブックを開いた後のローカル タイム ゾーンの日時。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|リンクされたデータ型が "Periodic" に設定されている場合に更新される頻度 ( `refreshMode` 秒)。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|リンクされたデータ型のデータを取得するメカニズム。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|リンクされたデータ型の一意の ID。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|リンクされたデータ型でサポートされているすべての更新モードを含む配列を返します。|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新しいリンク されたデータ型の一意の ID。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|サービス ID によってリンクされたデータ型を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|コレクション内のインデックスによってリンクされたデータ型を取得します。|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|ID によってリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定された名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスでシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|ピボットテーブルの代替テキスト タイトル。|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|各項目の後に空白行を表示するかどうかを設定します。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|次の場合、ピボットテーブルの空のセルに自動的に入力されるテキスト `fillEmptyCells == true` です。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|ピボットテーブル内の空のセルに次の値を設定するかどうかを指定します `emptyCellText` 。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルを指定します。|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|ピボットテーブル内のすべてのフィールドで [すべてのアイテム ラベルを繰り返す] 設定を設定します。|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用するスタイルを設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|ピボットテーブルにフィールド ヘッダー (フィールド キャプションとフィルター ドロップダウン) を表示するかどうかを指定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|ブックを開くとピボットテーブルを更新するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|同じワークシートまたは複数のワークシート内のセルのすべての参照元を含む範囲を表す `WorkbookRangeAreas` オブジェクトを返します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|更新モードが変更されたオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が成功したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|更新要求が完了したオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求から生成された警告を含む配列。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|スライサーに適用されるスタイル。|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用するスタイルを設定します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Table に適用されるスタイルを指定します。|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|表に適用するスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの ID を取得します。|
|[タスク](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|タスクに割り当て先を追加します。|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|指定された変更をタスクに適用します。|
||[assignees](/javascript/api/excel/excel.task#assignees)|タスクが割り当てられているユーザーを取得します。|
||[comment](/javascript/api/excel/excel.task#comment)|タスクに関連付けられているコメントを取得します。|
||[dueDate](/javascript/api/excel/excel.task#duedate)|タスクの期限の日時を取得します。|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|タスクの履歴レコードを取得します。|
||[id](/javascript/api/excel/excel.task#id)|タスクの ID を取得します。|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|タスクの完了率を取得します。|
||[priority](/javascript/api/excel/excel.task#priority)|タスクの優先度を取得します。|
||[startDate](/javascript/api/excel/excel.task#startdate)|タスクを開始する日時を取得します。|
||[title](/javascript/api/excel/excel.task#title)|タスクのタイトルを取得します。|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|タスクからすべての割り当て者を削除します。|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|タスクから割り当て先を削除します。|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|タスクの完了を変更します。|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|タスクの優先度を変更します。|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|タスクの開始日と期限を変更します。|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|タスクのタイトルを変更します。|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|タスクの新しい期限を UTC タイム ゾーンで設定します。|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|タスクに割り当てるユーザーの電子メール アドレスを設定します。|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|タスクの割り当てを解除するユーザーの電子メール アドレスを設定します。|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|タスクの新しい完了率を設定します。|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|タスクの新しい優先度を設定します。|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|変更によって以前のすべての割り当て先をタスクから削除する必要がある場合に設定します。|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|タスクの新しい開始日を UTC タイム ゾーンで設定します。|
||[title](/javascript/api/excel/excel.taskchanges#title)|タスクの新しいタイトルを設定します。|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|コレクション内のタスクの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|ID を使用してタスクを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|コレクション内のインデックスでタスクを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|ID を使用してタスクを取得します。|
||[items](/javascript/api/excel/excel.taskcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|タスクがアンカーされているオブジェクトの ID を表します (コメントに関連付けられたタスクの commentId など)。|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|"割り当て" 履歴レコードの種類のタスクに割り当てられているユーザー、または "割り当て解除" 履歴レコードの種類のタスクから割り当てを解除するユーザーを表します。|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|タスクを作成または変更したユーザーを表します。|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|タスクの期限を表します。|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|タスク履歴レコードの作成日を表します。|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|履歴レコードの ID。|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|タスクの完了率を表します。|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|タスクの優先度を表します。|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|タスクの開始日を表します。|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|タスクのタイトルを表します。|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|タスク履歴レコードの種類を表します。|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|"元に戻TaskHistoryRecord.id" 履歴レコードの種類に対して取り消された新しいプロパティを表します。|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|タスクのコレクション内の履歴レコードの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|コレクション内のインデックスを使用して、タスク履歴レコードを取得します。|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ユーザー](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.user#email)|ユーザーの電子メール アドレスを表します。|
||[uid](/javascript/api/excel/excel.user#uid)|ユーザーの一意の ID を表します。|
|[ブック](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|ブックの一部であるリンクされたデータ型のコレクションを返します。|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|ブックに存在するタスクのコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールド リスト ウィンドウをブック レベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートに存在するシート ビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|このワークシートで 1 つ以上の数式が変更された場合に発生します。|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|このコレクションのワークシートで 1 つ以上の数式が変更された場合に発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの ID を取得します。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|すべての変更された数式に関する詳細を含む FormulaChangedEventDetail オブジェクトの配列を取得します。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|数式が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
