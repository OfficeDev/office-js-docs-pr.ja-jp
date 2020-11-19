---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細。
ms.date: 11/17/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 083741d35d3e881c2e46b186c4e93591bf7f4834
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131767"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| リンクされたデータ型 | 外部ソースから Excel に接続されたデータ型のサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 指定したシートビュー | ユーザー単位のワークシートビューをプログラムによって制御します。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [タスク](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストについては、「 [すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[割り当てタスク (電子メール: 文字列)](/javascript/api/excel/excel.comment#assigntask-email-)|コメントに関連付けられているタスクを、指定されたユーザーに割り当てられた唯一の担当者として割り当てます。|
||[getTask ()](/javascript/api/excel/excel.comment#gettask--)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|このコメントに関連付けられているタスクを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[割り当てタスク (電子メール: 文字列)](/javascript/api/excel/excel.commentreply#assigntask-email-)|コメントに関連付けられているタスクを、指定されたユーザーに割り当てられた唯一の担当者として割り当てます。|
||[getTask ()](/javascript/api/excel/excel.commentreply#gettask--)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|このコメントに関連付けられているタスクを取得します。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[プロバイダー](/javascript/api/excel/excel.linkeddatatype#dataprovider)|リンクされたデータ型のデータプロバイダーの名前を指定します。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|リンクされたデータ型が最後に更新されたときに、ブックが開かれてからのローカルタイムゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前を指定します。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|リンクされたデータ型が `refreshMode` "定期的" に設定されている場合に更新される頻度 (秒単位)。|
||[示し](/javascript/api/excel/excel.linkeddatatype#refreshmode)|リンクされたデータ型のデータを取得するメカニズムを指定します。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|リンクされたデータ型の一意の id。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|リンクされたデータ型によってサポートされるすべての更新モードを含む配列を返します。|
||[requestRefresh ()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode (refreshMode: LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新しいリンクされたデータ型の一意の id。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|リンクされたデータ型をサービス id で取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|コレクション内のインデックスによって、リンクされたデータ型を取得します。|
||[getItemOrNullObject (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|ID でリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシートビューをアクティブにします。|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|ピボットテーブルの代替テキストタイトル。|
||[各アイテムを表示する (display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|各アイテムの後に空白行を表示するかどうかを設定します。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|ピボットテーブル内の空のセルに自動的に入力されるテキスト `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|ピボットテーブルの空のセルにを設定するかどうかを指定し `emptyCellText` ます。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイルです。|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|ピボットテーブルのすべてのフィールドで [すべてのアイテムのラベルを繰り返す] 設定を設定します。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|ピボットテーブルにフィールドヘッダーを表示するかどうかを指定します (フィールドのタイトルとフィルターのドロップダウン)。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|ブックを開くときにピボットテーブルを更新するかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents 元 ()](/javascript/api/excel/excel.range#getprecedents--)|`WorkbookRangeAreas`同じワークシートまたは複数のワークシート内のセルのすべての参照元を含む範囲を表すオブジェクト型 (object) の値を取得します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[示し](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|更新モードが変更されたオブジェクトの一意の id です。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が正常に終了したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|更新要求が完了したオブジェクトの一意の id。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[注意](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求によって生成された警告を含む配列。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|スライサーに適用されるスタイルです。|
||[setStyle (style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|表に適用されるスタイルです。|
||[setStyle (style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|表に適用するスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの id を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの id を取得します。|
|[タスク](/javascript/api/excel/excel.task)|[addAssignee (email: string)](/javascript/api/excel/excel.task#addassignee-email-)|タスクに実施者を追加します。|
||[applyChanges (taskChanges: Excel の変更)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|指定した変更をタスクに適用します。|
||[実施](/javascript/api/excel/excel.task#assignees)|タスクが割り当てられているユーザーを取得します。|
||[comment](/javascript/api/excel/excel.task#comment)|タスクに関連付けられているコメントを取得します。|
||[dueDate](/javascript/api/excel/excel.task#duedate)|タスクの期限の日付と時刻を取得します。|
||[履歴レコード](/javascript/api/excel/excel.task#historyrecords)|タスクの履歴レコードを取得します。|
||[id](/javascript/api/excel/excel.task#id)|タスクの id を取得します。|
||[達成](/javascript/api/excel/excel.task#percentcomplete)|タスクの達成率を取得します。|
||[priority](/javascript/api/excel/excel.task#priority)|タスクの優先度を取得します。|
||[startDate](/javascript/api/excel/excel.task#startdate)|タスクが開始する日付と時刻を取得します。|
||[title](/javascript/api/excel/excel.task#title)|タスクのタイトルを取得します。|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|タスクからすべてのタスク実施者を削除します。|
||[removeAssignee (email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|タスクから担当者を削除します。|
||[setPercentComplete 率 (達成率: 数値)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|タスクの完了を変更します。|
||[setPriority (priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|タスクの優先度を変更します。|
||[setStartDateAndDueDate (startDate: Date、dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|タスクの開始日と期限を変更します。|
||[setTitle (title: string)](/javascript/api/excel/excel.task#settitle-title-)|タスクのタイトルを変更します。|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|タスクの新しい期限を UTC タイムゾーンで設定します。|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|タスクに割り当てるユーザーの電子メールアドレスを設定します。|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|タスクの割り当てを解除するユーザーの電子メールアドレスを設定します。|
||[達成](/javascript/api/excel/excel.taskchanges#percentcomplete)|タスクの新しい達成率を設定します。|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|タスクの新しい優先度を設定します。|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|変更によって、タスクから以前のすべての担当者を削除する必要があるかどうかを設定します。|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|タスクの新しい開始日を UTC タイムゾーンで設定します。|
||[title](/javascript/api/excel/excel.taskchanges#title)|タスクの新しいタイトルを設定します。|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|コレクション内のタスクの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Id を使用してタスクを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|コレクション内のインデックスによってタスクを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Id を使用してタスクを取得します。|
||[items](/javascript/api/excel/excel.taskcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Taskhistory レコード](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|タスクが固定されているオブジェクトの ID を表します (たとえば、コメントに添付されたタスクの commentId)。|
||[担当者](/javascript/api/excel/excel.taskhistoryrecord#assignee)|[割り当て] 履歴レコードの種類に対してタスクに割り当てられているユーザー、またはタスクの割り当てを解除するユーザーを表します。履歴レコードの種類を "割り当て解除" します。|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|タスクを作成または変更したユーザーを表します。|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|タスクの締め切り日を表します。|
||[履歴レコードの指定日](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|タスク履歴レコードの作成日を表します。|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|履歴レコードの ID。|
||[達成](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|タスクの達成率を表します。|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|タスクの優先度を表します。|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|タスクの開始日を表します。|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|タスクのタイトルを表します。|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|タスク履歴レコードの種類を表します。|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|"元に戻す" 履歴レコードの種類では、元に戻された TaskHistoryRecord.id プロパティを表します。|
|[Taskhistory Recordcollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|タスクのコレクション内の履歴レコードの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|コレクション内のインデックスを使用して、タスク履歴レコードを取得します。|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ユーザー](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.user#email)|ユーザーの電子メール アドレスを表します。|
||[uid](/javascript/api/excel/excel.user#uid)|ユーザーの一意の ID を表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes 型](/javascript/api/excel/excel.workbook#linkeddatatypes)|ブックの一部である、リンクされたデータ型のコレクションを返します。|
||[タスク](/javascript/api/excel/excel.workbook#tasks)|ブック内に存在するタスクのコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールドリストウィンドウをブックレベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートにあるシートビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[タスク](/javascript/api/excel/excel.worksheet#tasks)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
