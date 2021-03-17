---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript API の詳細。
ms.date: 03/10/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b4a2db19ce04d316cf106dcd97f2d71f0f009e55
ms.sourcegitcommit: 929dcf2f415b94f42330a9035ed11a5cedad88f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/16/2021
ms.locfileid: "50830980"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| ドキュメント タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| 数式の変更イベント | 変更の原因となるイベントのソースと種類を含む、数式の変更を追跡します。 | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| リンクされたデータ型 | 外部ソースから Excel に接続されているデータ型のサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| ピボットテーブル ピボットレイアウト | Alt テキストと空のセル管理の新しいサポートを含む、PivotLayout クラスの拡張。 | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |
| テーブルのスタイル | フォント、罫線、塗りつぶしの色、および表のスタイルの他の側面のコントロールを提供します。 | [Table](/javascript/api/excel/excel.table)、 [PivotTable](/javascript/api/excel/excel.pivottable)、 [Slicer](/javascript/api/excel/excel.slicer) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript API の一覧を示します。 すべての Excel JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧については、 [すべての Excel JavaScript API を参照してください](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|オートフィルターのフィルター条件がクリアされます。|
|[コメント](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|コメントに添付されたタスクを、割り当て先として指定されたユーザーに割り当てる。|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|このコメントに関連付けられているタスクを取得します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|ID に基づいてコレクションからコメントを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|その ID で識別されるコメント返信を返します。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|ID で識別される条件付き書式を返します。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|タスクの完了率を指定します。|
||[priority](/javascript/api/excel/excel.documenttask#priority)|タスクの優先度を指定します。|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|タスクの割り当て人のコレクションを返します。|
||[変更点](/javascript/api/excel/excel.documenttask#changes)|タスクの変更レコードを取得します。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|タスクに関連付けられたコメントを取得します。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|タスクを完了した最新のユーザーを取得します。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|タスクが完了した日時を取得します。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|タスクを作成したユーザーを取得します。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|タスクが作成された日時を取得します。|
||[id](/javascript/api/excel/excel.documenttask#id)|タスクの ID を取得します。|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|タスクの開始日と期日を変更します。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|タスクを開始する日付と時刻を取得または設定します。期限が設定されます。|
||[title](/javascript/api/excel/excel.documenttask#title)|タスクのタイトルを指定します。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[割り当て先](/javascript/api/excel/excel.documenttaskchange#assignee)|変更レコードの種類のタスクに割り当てられたユーザー、または変更レコードの種類のタスクから割り当てられていないユーザー `assign` `unassign` を表します。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|タスクを作成または変更したユーザーを表します。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|タスクの変更をアンカー `Comment` する ID `CommentReply` を表します。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|タスク変更レコードの作成日時を表します。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|タスクの期日と時刻を UTC タイム ゾーンで表します。|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|タスク変更レコードの ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|タスクの完了率を表します。|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|タスクの優先度を表します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|タスクの開始日時を UTC タイム ゾーンで表します。|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|タスクのタイトルを表します。|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|タスク変更レコードのアクションの種類を表します。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|変更レコードの `DocumentTaskChange.id` 種類に対して元に戻されたプロパティ `undo` を表します。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|タスクのコレクション内の変更レコードの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|コレクション内のインデックスを使用してタスク変更レコードを取得します。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|コレクション内のタスクの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|ID を使用してタスクを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|コレクション内のインデックスによってタスクを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|ID を使用してタスクを取得します。|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|タスクが期限の日時を取得します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|タスクを開始する日付と時刻を取得します。|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|変更された数式を含むセルのアドレス。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|変更前の数式を表します。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|名前または ID を使用して図形を取得します。|
|[ID](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.identity#email)|ユーザーの電子メール アドレスを表します。|
||[id](/javascript/api/excel/excel.identity#id)|ユーザーの一意の ID を表します。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|コレクションにユーザー ID を追加します。|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|コレクションからすべてのユーザー ID を削除します。|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|コレクション内のアイテムの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|コレクション内のインデックスを使用してドキュメント ユーザー ID を取得します。|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|コレクションからユーザー ID を削除します。|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|新しいワークシートの現在のブック内の挿入位置。|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|パラメーターに対して参照されている現在のブック内の `WorksheetPositionType` ワークシート。|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|挿入する個々のワークシートの名前。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|リンクされたデータ型のデータ プロバイダーの名前。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|リンクされたデータ型が最後に更新されたときにブックが開か以降のローカルタイム ゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|リンクされたデータ型が "定期的" に設定されている場合に更新される頻度 (秒 `refreshMode` )。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|リンクされたデータ型のデータを取得するメカニズム。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|リンクされたデータ型の一意の ID。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|リンクされたデータ型でサポートされているすべての更新モードを持つ配列を返します。|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新しいリンクされたデータ型の一意の ID。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|サービス ID 別にリンクされたデータ型を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|コレクション内のインデックスによってリンクされたデータ型を取得します。|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|ID によってリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定した名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスによってシート ビューを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|名前を使用してシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|ピボットテーブルの代替テキスト タイトル。|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|各項目の後に空白行を表示するかどうかを設定します。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|ピボットテーブル内の空のセルに自動的に入力されるテキスト `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|ピボットテーブルの空のセルに、 を設定するかどうかを指定します `emptyCellText` 。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|ピボットテーブルに適用されるスタイル。|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|ピボットテーブルのすべてのフィールドで[すべてのアイテム ラベルを繰り返す] 設定を設定します。|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|ピボットテーブルにフィールド ヘッダー (フィールド キャプションとフィルター ドロップダウン) を表示するかどうかを指定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|ブックが開くとピボットテーブルが更新されるかどうかを指定します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|コレクション内の最初のピボットテーブルを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|同じワークシートまたは複数のワークシート内のセルのすべての従属セルを含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|同じワークシートまたは複数のワークシート内のセルのすべての直接依存を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|指定された方向に基づいて、現在の範囲と範囲の端までの範囲オブジェクトを返します。|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|この範囲内の結合領域を表す RangeAreas オブジェクトを返します。|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|同じワークシートまたは複数のワークシート内のセルのすべての前例を含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|指定された方向に対応するデータ領域のエッジ セルである範囲オブジェクトを返します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|更新モードが変更されたオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が成功したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|更新要求が完了したオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求から生成された警告を含む配列。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|名前または ID を使用して図形を取得します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|スライサーに適用されるスタイル。|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|名前に基づいてスタイルを取得します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|特定のテーブルにフィルターが適用されると発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|テーブルに適用されるスタイル。|
||[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|テーブルのサイズを新しい範囲に変更します。|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|テーブルに適用されるスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシート内の任意のテーブルにフィルターが適用されると発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの ID を取得します。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|名前または ID でテーブルを取得します。|
|[ブック](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|指定したワークシートをソース ブックから現在のブックに挿入します。|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|ブックの一部であるリンクされたデータ型のコレクションを返します。|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|ブックがアクティブ化されると発生します。|
||[タスク](/javascript/api/excel/excel.workbook#tasks)|ブックに存在するタスクのコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールド 一覧ウィンドウをブック レベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|イベントの種類を取得します。|
|[ワークシート](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートに存在するシート ビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|特定のワークシートにフィルターが適用されると発生します。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|このワークシートで 1 つ以上の数式が変更された場合に発生します。|
||[タスク](/javascript/api/excel/excel.worksheet#tasks)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|このコレクションのワークシートで 1 つ以上の数式が変更された場合に発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されるワークシートの ID を取得します。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|変更された数式 `FormulaChangedEventDetail` の詳細を含むオブジェクトの配列を取得します。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|数式が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
