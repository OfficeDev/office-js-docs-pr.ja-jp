---
title: Excel JavaScript プレビュー API
description: JavaScript API のExcel詳細。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 32a2f5d355086c51cbf165dd7ed335e96c96647a
ms.sourcegitcommit: ddb1d85186fd6e77d732159430d20eb7395b9a33
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/10/2021
ms.locfileid: "61406642"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

次の表に、API の簡潔な概要を示しますが、後続の API リスト [テーブル](#api-list) には詳細な一覧が示されています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [データ型](../../excel/excel-data-types-overview.md) | 書式付き番号と web Excelのサポートを含む、既存のデータ型の拡張。 | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue), [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [データ型のエラー](../../excel/excel-data-types-concepts.md#improved-error-support) | 拡張データ型をサポートするエラー オブジェクト。 | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue , NumErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [refErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue) [](/javascript/api/excel/excel.spillerrorcellvalue)|
| ドキュメント タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| ID | 表示名や電子メール アドレスなど、ユーザー ID を管理します。 | [Identity](/javascript/api/excel/excel.identity)、 [IdentityCollection](/javascript/api/excel/excel.identitycollection)、 [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| リンクされたデータ型 | 外部ソースからデータに接続されたデータExcelサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| テーブルのスタイル | フォント、罫線、塗りつぶしの色、および表のスタイルの他の側面のコントロールを提供します。 | [Table](/javascript/api/excel/excel.table)、 [PivotTable](/javascript/api/excel/excel.pivottable)、 [Slicer](/javascript/api/excel/excel.slicer) |
| ワークシートの保護 | 承認されていないユーザーがワークシート内で指定した範囲に変更を加えなかねない。 | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中Excel JavaScript API の一覧を示します。 すべての JavaScript API (プレビュー API Excel以前にリリースされた API を含む) の完全な一覧については[、JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)API Excel参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|オブジェクトに関連付けられている範囲を指定します。|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|からこのオブジェクトを削除します `AllowEditRangeCollection` 。|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|is is password `AllowEditRange` protected を指定します。|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|特定のセッションのユーザーの特定 `AllowEditRange` のオブジェクトに対するワークシートの保護を一時停止します。|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|に関連付けられているパスワードを変更 `AllowEditRange` します。|
||[title](/javascript/api/excel/excel.alloweditrange#title)|オブジェクトのタイトルを指定します。|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|コレクションに `AllowEditRange` オブジェクトを追加します。|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|コレクション内のオブジェクト `AllowEditRange` の数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|タイトルによって `AllowEditRange` オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|コレクション内の `AllowEditRange` インデックスによってオブジェクトを返します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|タイトルによって `AllowEditRange` オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|特定のセッションでユーザーに対して指定されたパスワードを持つコレクション内のすべてのオブジェクトに対するワークシート保護 `AllowEditRange` を一時停止します。|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#password)|に関連付けられている `AllowEditRange` パスワード。|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[要素](/javascript/api/excel/excel.arraycellvalue#elements)|配列の要素を表します。|
||[type](/javascript/api/excel/excel.arraycellvalue#type)|このセル値の種類を表します。|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|の種類を表します `BlockedErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|このセル値の種類を表します。|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|このセル値の種類を表します。|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|の種類を表します `BusyErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|このセル値の種類を表します。|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|の種類を表します `CalcErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|このセル値の種類を表します。|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[レイアウト](/javascript/api/excel/excel.cardlayoutlistsection#layout)|このセクションのレイアウトの種類を表します。|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#property)|カード レイアウトによって参照されるプロパティの名前。|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[折りたたむ](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsed)|カードのこのセクションが最初に折りたたまれるかどうかを表します。|
||[折りたたみ可能](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsible)|カードのこのセクションが折りたたみ可能かどうかを表します。|
||[プロパティ](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#properties)|このセクションのプロパティの名前を表します。|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#title)|カードのこのセクションのタイトルを表します。|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#mainImage)|カードのメイン イメージとして使用するプロパティを指定します。|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#sections)|カードのセクションを表します。|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#subTitle)|カードのサブタイトルを含むプロパティの仕様を表します。|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#title)|カードのタイトル、またはカードのタイトルを含むプロパティの仕様を表します。|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[レイアウト](/javascript/api/excel/excel.cardlayouttablesection#layout)|このセクションのレイアウトの種類を表します。|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|このプロパティの使用方法を説明するライセンスまたはソースの URL を表します。|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|このプロパティを管理するライセンスの名前を表します。|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|ソースの URL を表します `CellValue` 。|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|のソースの名前を表します `CellValue` 。|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[属性](/javascript/api/excel/excel.cellvaluepropertymetadata#attribution)|このプロパティを使用するソース要件とライセンス要件を説明する属性情報を表します。|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excludeFrom)|このプロパティが除外される機能を表します。|
||[サブラベル](/javascript/api/excel/excel.cellvaluepropertymetadata#sublabel)|カード ビューに表示されるこのプロパティのサブラベルを表します。|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#autoComplete)|True は、プロパティがオートコンプリートによって表示されるプロパティから除外されます。|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#calcCompare)|True は、プロパティが再計算時にセルの値を比較するために使用されるプロパティから除外されます。|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#cardView)|True は、プロパティがカード ビューで表示されるプロパティから除外されます。|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#dotNotation)|True は、プロパティが FIELDVALUE 関数を介してアクセスできるプロパティから除外されます。|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|ロゴが指定されていない場合にカード ビューで使用されるプロバイダーの説明プロパティを表します。|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|カード ビューでロゴとして使用される画像をダウンロードするために使用される URL を表します。|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|ユーザーがカード ビューのロゴ要素をクリックした場合のナビゲーション ターゲットの URL を表します。|
|[コメント](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|コメントに添付されたタスクを、割り当て先として指定されたユーザーに割り当てる。|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|このコメントに関連付けられているタスクを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|このコメント返信のスレッドに関連付けられているタスクを取得します。|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|の種類を表します `ConnectErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|このセル値の種類を表します。|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|このセル値の種類を表します。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#assignees)|タスクの割り当て人のコレクションを返します。|
||[変更点](/javascript/api/excel/excel.documenttask#changes)|タスクの変更レコードを取得します。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|タスクに関連付けられたコメントを取得します。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|タスクを完了した最新のユーザーを取得します。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|タスクが完了した日時を取得します。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|タスクを作成したユーザーを取得します。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|タスクが作成された日時を取得します。|
||[id](/javascript/api/excel/excel.documenttask#id)|タスクの ID を取得します。|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|タスクの完了率を指定します。|
||[優先度](/javascript/api/excel/excel.documenttask#priority)|タスクの優先度を指定します。|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|このセル値の種類を表します。|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|このセル値の種類を表します。|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[レイアウト](/javascript/api/excel/excel.entitycardlayout#layout)|このレイアウトの種類を表します。|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#cardLayout)|カード ビューでこのエンティティのレイアウトを表します。|
||[properties: { [key: string]](/javascript/api/excel/excel.entitycellvalue#properties)|このエンティティのプロパティとそのメタデータを表します。|
||[text](/javascript/api/excel/excel.entitycellvalue#text)|この値を持つセルがレンダリングされる場合に表示されるテキストを表します。|
||[type](/javascript/api/excel/excel.entitycellvalue#type)|このセル値の種類を表します。|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|の種類を表します `FieldErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|このセル値の種類を表します。|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|この値の表示に使用される数値書式指定文字列を返します。|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|このセル値の種類を表します。|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|このセル値の種類を表します。|
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
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|このリンクされたデータ型の更新モードを変更する要求を行います。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|リンクされたデータ型の一意の ID。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|リンクされたデータ型でサポートされているすべての更新モードを持つ配列を返します。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|新しいリンクされたデータ型の一意の ID。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|サービス ID 別にリンクされたデータ型を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|コレクション内のインデックスによってリンクされたデータ型を取得します。|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|ID によってリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|このセル値の種類を表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#valueAsJson)|この名前付きアイテムの値の JSON 表記。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#valuesAsJson)|この範囲内のセル内の値の JSON 表記。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|名前を使用してシート ビューを取得します。|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#type)|このセル値の種類を表します。|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|このセル値の種類を表します。|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|このセル値の種類を表します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|ピボットテーブルに適用されるスタイル。|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#getDataSourceString__)|ピボットテーブルのデータ ソースの文字列表現を返します。|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#getDataSourceType__)|ピボットテーブルのデータ ソースの種類を取得します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|コレクション内の最初のピボットテーブルを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|同じワークシートまたは複数のワークシート内のセルのすべての従属セルを含む範囲を表すオブジェクト `WorkbookRangeAreas` を返します。|
||[valuesAsJson](/javascript/api/excel/excel.range#valuesAsJson)|この範囲内のセル内の値の JSON 表記。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#valuesAsJson)|この範囲内のセル内の値の JSON 表記。|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|の種類を表します `RefErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|このセル値の種類を表します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|更新モードが変更されたオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が成功したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|更新要求が完了したオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求から生成された警告を含む配列。|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|図形の表示名を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|数式で使用するスライサーの名前を表します。|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|スライサーに適用されるスタイルを設定します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|スライサーに適用されるスタイル。|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|の種類を表します `SpillErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|データが表示された場合に流出する列の数を#SPILL! エラーを返します。|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|データが表示された場合に流出する行の数を表#SPILL! エラーを返します。|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|このセル値の種類を表します。|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|このセル値の種類を表します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|特定のテーブルにフィルターが適用されると発生します。|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|テーブルに適用されるスタイルを設定します。|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|テーブルに適用されるスタイル。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|ブックまたはワークシート内の任意のテーブルにフィルターが適用されると発生します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#valuesAsJson)|このテーブル列のセル内の値の JSON 表記。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|フィルターが適用されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|テーブルを含むワークシートの ID を取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#valuesAsJson)|この表の行のセル内の値の JSON 表記。|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|の種類を表します `ValueErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|の種類を表します `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|このセル値の種類を表します。|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|このセル値の種類を表します。|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|イメージのダウンロード先 URL を表します。|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|イメージが何を表すのかを説明するためにアクセシビリティ シナリオで使用できる代替テキストを表します。|
||[属性](/javascript/api/excel/excel.webimagecellvalue#attribution)|この画像を使用するソース要件とライセンス要件を説明する属性情報を表します。|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#basicType)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#basicValue)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[プロバイダー](/javascript/api/excel/excel.webimagecellvalue#provider)|画像を提供したエンティティまたは個人を表す情報を表します。|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|このページに関連すると見なされる画像を含む Web ページの URL を表します `WebImageCellValue` 。|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|このセル値の種類を表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|ブックの一部であるリンクされたデータ型のコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|ピボットテーブルのフィールド 一覧ウィンドウをブック レベルで表示するかどうかを指定します。|
||[タスク](/javascript/api/excel/excel.workbook#tasks)|ブックに存在するタスクのコレクションを返します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|特定のワークシートにフィルターが適用されると発生します。|
||[タスク](/javascript/api/excel/excel.worksheet#tasks)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|フィルターが適用されるワークシートの ID を取得します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|このワークシートで `AllowEditRangeCollection` 見つかったファイルを指定します。|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|このワークシートの保護を一時停止できる場合に指定します。|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|ワークシート保護のロック解除にパスワードを使用できる場合に指定します。|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|シートがパスワードで保護される場合に指定します。|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|ワークシートの保護を一時停止する場合に指定します。|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|特定のセッションのユーザーに対する、指定されたワークシート オブジェクトのワークシート保護を一時停止します。|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|特定のセッションのユーザーに対する、指定されたワークシート オブジェクトのワークシート保護を再開します。|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|オブジェクトに関連付けられているパスワードを変更 `WorksheetProtection` します。|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|オブジェクトに関連付けられているワークシート保護オプションを変更 `WorksheetProtection` します。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|オブジェクトが変更された `AllowEditRange` 場合に指定します。|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|変更された場合に `WorksheetProtectionOptions` 指定します。|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|ワークシートのパスワードが変更された場合に指定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
