---
title: Excel JavaScript プレビュー API
description: JavaScript API のExcel詳細。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f15a72631f83a5102fb4e042cc1357d179d1fa3d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747179"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

次の表に、API の簡潔な概要を示しますが、後続の [API リスト テーブル](#api-list) には詳細な一覧が示されています。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [データ型](../../excel/excel-data-types-overview.md) | 書式付き番号と web Excelのサポートを含む、既存のデータ型の拡張。 | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)、 [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)、 [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)、 [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)、 [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)、 [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)、 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)、 [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)、 [StringCellValue](/javascript/api/excel/excel.stringcellvalue)、 [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)、 [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [データ型のエラー](../../excel/excel-data-types-concepts.md#improved-error-support) | 拡張データ型をサポートするエラー オブジェクト。 | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)、[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)、[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)、[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)、[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)、[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)、[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)、[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)、[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)、[NullErrorCellValue、NumErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)、RefErrorCellValue、[](/javascript/api/excel/excel.numerrorcellvalue)[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)、[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue) [](/javascript/api/excel/excel.referrorcellvalue)|
| ドキュメント タスク | コメントをユーザーに割り当てられたタスクに変換します。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| ID | 表示名や電子メール アドレスなど、ユーザー ID を管理します。 | [IDENTITY](/javascript/api/excel/excel.identity)、[IdentityCollection、](/javascript/api/excel/excel.identitycollection)[IdentityEntity](/javascript/api/excel/excel.identityentity) |
| リンクされたデータ型 | 外部ソースからデータに接続されたデータExcelサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| テーブルのスタイル | フォント、罫線、塗りつぶしの色、および表のスタイルの他の側面のコントロールを提供します。 | [テーブル](/javascript/api/excel/excel.table)、[ピボットテーブル、](/javascript/api/excel/excel.pivottable)[スライサー](/javascript/api/excel/excel.slicer) |
| ワークシートの保護 | 承認されていないユーザーがワークシート内で指定した範囲に変更を加えなかねない。 | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中Excel JavaScript API の一覧を示します。 すべての JavaScript API (プレビュー API Excel以前にリリースされた API を含む) の完全な一覧については、[JavaScript API Excel参照してください](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|オブジェクトに関連付けられている範囲を指定します。|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|からこのオブジェクトを削除します `AllowEditRangeCollection`。|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|is is password `AllowEditRange` protected を指定します。|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|特定のセッションのユーザーの特定 `AllowEditRange` のオブジェクトに対するワークシートの保護を一時停止します。|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|に関連付けられているパスワードを変更します `AllowEditRange`。|
||[title](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|オブジェクトのタイトルを指定します。|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|コレクションに `AllowEditRange` オブジェクトを追加します。|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|コレクション内のオブジェクトの `AllowEditRange` 数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|タイトルによって `AllowEditRange` オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|コレクション内のインデックス `AllowEditRange` によってオブジェクトを返します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|タイトルによって `AllowEditRange` オブジェクトを取得します。|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|特定のセッションでユーザーに `AllowEditRange` 対して指定されたパスワードを持つコレクション内のすべてのオブジェクトに対するワークシート保護を一時停止します。|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|に関連付けられているパスワード `AllowEditRange`。|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[要素](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-elements-member)|配列の要素を表します。|
||[type](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-type-member)|このセル値の種類を表します。|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|の種類を表します `BlockedErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-type-member)|このセル値の種類を表します。|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|このセル値の種類を表します。|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|の種類を表します `BusyErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-type-member)|このセル値の種類を表します。|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|の種類を表します `CalcErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-type-member)|このセル値の種類を表します。|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[レイアウト](/javascript/api/excel/excel.cardlayoutlistsection#excel-excel-cardlayoutlistsection-layout-member)|このセクションのレイアウトの種類を表します。|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#excel-excel-cardlayoutpropertyreference-property-member)|カード レイアウトによって参照されるプロパティの名前。|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[折りたたむ](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsed-member)|カードのこのセクションが最初に折りたたまれるかどうかを表します。|
||[折りたたみ可能](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsible-member)|カードのこのセクションが折りたたみ可能かどうかを表します。|
||[プロパティ](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-properties-member)|このセクションのプロパティの名前を表します。|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-title-member)|カードのこのセクションのタイトルを表します。|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-mainimage-member)|カードのメイン イメージとして使用するプロパティを指定します。|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-sections-member)|カードのセクションを表します。|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-subtitle-member)|カードのサブタイトルを含むプロパティの仕様を表します。|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-title-member)|カードのタイトル、またはカードのタイトルを含むプロパティの仕様を表します。|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[レイアウト](/javascript/api/excel/excel.cardlayouttablesection#excel-excel-cardlayouttablesection-layout-member)|このセクションのレイアウトの種類を表します。|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licenseaddress-member)|このプロパティの使用方法を説明するライセンスまたはソースの URL を表します。|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licensetext-member)|このプロパティを管理するライセンスの名前を表します。|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourceaddress-member)|ソースの URL を表します `CellValue`。|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourcetext-member)|のソースの名前を表します `CellValue`。|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[属性](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-attribution-member)|このプロパティを使用するソース要件とライセンス要件を説明する属性情報を表します。|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-excludefrom-member)|このプロパティが除外される機能を表します。|
||[サブラベル](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-sublabel-member)|カード ビューに表示されるこのプロパティのサブラベルを表します。|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-autocomplete-member)|True は、プロパティがオートコンプリートによって表示されるプロパティから除外されます。|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-calccompare-member)|True は、プロパティが再計算時にセルの値を比較するために使用されるプロパティから除外されます。|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-cardview-member)|True は、プロパティがカード ビューで表示されるプロパティから除外されます。|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-dotnotation-member)|True は、プロパティが FIELDVALUE 関数を介してアクセスできるプロパティから除外されます。|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[説明](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-description-member)|ロゴが指定されていない場合にカード ビューで使用されるプロバイダーの説明プロパティを表します。|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logosourceaddress-member)|カード ビューでロゴとして使用される画像をダウンロードするために使用される URL を表します。|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logotargetaddress-member)|ユーザーがカード ビューのロゴ要素をクリックした場合のナビゲーション ターゲットの URL を表します。|
|[コメント](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|コメントに添付されたタスクを、割り当て先として指定されたユーザーに割り当てる。|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|このコメントに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|このコメントに関連付けられているタスクを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|コメントに添付されたタスクを、特定のユーザーに唯一の割り当て先として割り当てる。|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|このコメント返信のスレッドに関連付けられているタスクを取得します。|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|このコメント返信のスレッドに関連付けられているタスクを取得します。|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|の種類を表します `ConnectErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-type-member)|このセル値の種類を表します。|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-type-member)|このセル値の種類を表します。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|タスクの割り当て人のコレクションを返します。|
||[変更点](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|タスクの変更レコードを取得します。|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|タスクに関連付けられたコメントを取得します。|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|タスクを完了した最新のユーザーを取得します。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|タスクが完了した日時を取得します。|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|タスクを作成したユーザーを取得します。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|タスクが作成された日時を取得します。|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|タスクの ID を取得します。|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|タスクの完了率を指定します。|
||[優先度](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|タスクの優先度を指定します。|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-setstartandduedatetime-member(1))|タスクの開始日と期日を変更します。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|タスクを開始する日付と時刻を取得または設定します。期限が設定されます。|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|タスクのタイトルを指定します。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[割り当て先](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|変更レコードの種類のタスクに割 `assign` り当てられたユーザー、または変更レコードの種類のタスクから割り当てられていないユーザーを `unassign` 表します。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|タスクを作成または変更したユーザーを表します。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|タスクの変更をアンカーする `Comment` ID `CommentReply` を表します。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|タスク変更レコードの作成日時を表します。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|タスクの期日と時刻を UTC タイム ゾーンで表します。|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|タスク変更レコードの ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|タスクの完了率を表します。|
||[優先度](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|タスクの優先度を表します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|タスクの開始日時を UTC タイム ゾーンで表します。|
||[title](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|タスクのタイトルを表します。|
||[type](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|タスク変更レコードのアクションの種類を表します。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undohistoryid-member)|変更レコードの種類 `DocumentTaskChange.id` に対して元に戻されたプロパティを `undo` 表します。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|タスクのコレクション内の変更レコードの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|コレクション内のインデックスを使用してタスク変更レコードを取得します。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|コレクション内のタスクの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|ID を使用してタスクを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|コレクション内のインデックスによってタスクを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|ID を使用してタスクを取得します。|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|タスクが期限の日時を取得します。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|タスクを開始する日付と時刻を取得します。|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|このセル値の種類を表します。|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-type-member)|このセル値の種類を表します。|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[レイアウト](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|このレイアウトの種類を表します。|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-cardlayout-member)|カード ビューでこのエンティティのレイアウトを表します。|
||[properties: { [key: string]](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member)|このエンティティのプロパティとそのメタデータを表します。|
||[text](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-text-member)|この値を持つセルがレンダリングされる場合に表示されるテキストを表します。|
||[type](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-type-member)|このセル値の種類を表します。|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errorsubtype-member)|の種類を表します `FieldErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-type-member)|このセル値の種類を表します。|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-numberformat-member)|この値の表示に使用される数値書式指定文字列を返します。|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-type-member)|このセル値の種類を表します。|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-type-member)|このセル値の種類を表します。|
|[ID](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.identity#excel-excel-identity-email-member)|ユーザーの電子メール アドレスを表します。|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|ユーザーの一意の ID を表します。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-add-member(1))|コレクションにユーザー ID を追加します。|
||[clear()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-clear-member(1))|コレクションからすべてのユーザー ID を削除します。|
||[getCount()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getcount-member(1))|コレクション内のアイテムの数を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getitemat-member(1))|コレクション内のインデックスを使用してドキュメント ユーザー ID を取得します。|
||[items](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-remove-member(1))|コレクションからユーザー ID を削除します。|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-displayname-member)|ユーザーの表示名を表します。|
||[email](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-email-member)|ユーザーの電子メール アドレスを表します。|
||[id](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-id-member)|ユーザーの一意の ID を表します。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|リンクされたデータ型のデータ プロバイダーの名前。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|リンクされたデータ型が最後に更新されたときにブックが開か以降のローカルタイム ゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|リンクされたデータ型の名前。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|リンクされたデータ型が " `refreshMode` 定期的" に設定されている場合に更新される頻度 (秒)。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|リンクされたデータ型のデータを取得するメカニズム。|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|リンクされたデータ型を更新する要求を行います。|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|このリンクされたデータ型の更新モードを変更する要求を行います。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|リンクされたデータ型の一意の ID。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|リンクされたデータ型でサポートされているすべての更新モードを持つ配列を返します。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|新しいリンクされたデータ型の一意の ID。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|イベントの種類を取得します。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|サービス ID 別にリンクされたデータ型を取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|コレクション内のインデックスによってリンクされたデータ型を取得します。|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|ID によってリンクされたデータ型を取得します。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)|[basicType](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[id](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-id-member)|この値の情報を提供したサービス ソースを表します。|
||[properties: { [key: string]: CellValue & { propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-properties-member)|このエンティティのプロパティとそのメタデータを表します。|
||[propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-propertymetadata-member)||
||[プロバイダー](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-provider-member)|イメージを提供したサービスを説明する情報を表します。|
||[text](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-text-member)|この値を持つセルがレンダリングされる場合に表示されるテキストを表します。|
||[type](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-type-member)|このセル値の種類を表します。|
|[LinkedEntityId](/javascript/api/excel/excel.linkedentityid)|[カルチャ](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-culture-member)|これを作成するために使用された言語カルチャを表します `CellValue`。|
||[domainId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-domainid-member)|を作成するために使用するサービスに固有のドメインを表します `CellValue`。|
||[entityId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-entityid-member)|を作成するために使用するサービスに固有の識別子を表します `CellValue`。|
||[serviceId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-serviceid-member)|を作成するために使用されたサービスを表します `CellValue`。|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-type-member)|このセル値の種類を表します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjson-member)|この名前付きアイテムの値の JSON 表記。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjson-member)|この範囲内のセル内の値の JSON 表記。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|名前を使用してシート ビューを取得します。|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-type-member)|このセル値の種類を表します。|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-type-member)|このセル値の種類を表します。|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-type-member)|このセル値の種類を表します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。 |
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|ピボットテーブルに適用されるスタイル。|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|ピボットテーブルに適用されるスタイルを設定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcestring-member(1))|ピボットテーブルのデータ ソースの文字列表現を返します。|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcetype-member(1))|ピボットテーブルのデータ ソースの種類を取得します。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirstornullobject-member(1))|コレクション内の最初のピボットテーブルを取得します。|
|[PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue)|[basicType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[target](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-target-member)|`PlaceholderErrorCellValue` は処理中に使用され、データはダウンロードされます。|
||[type](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-type-member)|このセル値の種類を表します。|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1))|同じワークシート `WorkbookRangeAreas` または複数のワークシート内のセルのすべての従属セルを含む範囲を表すオブジェクトを返します。|
||[valuesAsJson](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member)|この範囲内のセル内の値の JSON 表記。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjson-member)|この範囲内のセル内の値の JSON 表記。|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|の種類を表します `RefErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-type-member)|このセル値の種類を表します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|更新モードが変更されたオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|イベントの種類を取得します。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|更新要求が成功したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|更新要求が完了したオブジェクトの一意の ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|イベントの種類を取得します。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|更新要求から生成された警告を含む配列。|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#excel-excel-shape-displayname-member)|図形の表示名を取得します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|数式で使用するスライサーの名前を表します。|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|スライサーに適用されるスタイルを設定します。|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|スライサーに適用されるスタイル。|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errorsubtype-member)|の種類を表します `SpillErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledcolumns-member)|データが表示された場合に流出する列の数を表#SPILL! エラーを返します。|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledrows-member)|データが表示された場合に流出する行の数を表#SPILL! エラーを返します。|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-type-member)|このセル値の種類を表します。|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|このセル値の種類を表します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|特定のテーブルにフィルターが適用されると発生します。|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|テーブルに適用されるスタイルを設定します。|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|テーブルに適用されるスタイル。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|ブックまたはワークシート内の任意のテーブルにフィルターが適用されると発生します。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjson-member)|このテーブル列のセル内の値の JSON 表記。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|フィルターが適用されるテーブルの ID を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|テーブルを含むワークシートの ID を取得します。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjson-member)|この表の行のセル内の値の JSON 表記。|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|の種類を表します `ValueErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errortype-member)|の種類を表します `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-type-member)|このセル値の種類を表します。|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-type-member)|このセル値の種類を表します。|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-address-member)|イメージのダウンロード先 URL を表します。|
||[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member)|イメージが何を表すのかを説明するためにアクセシビリティ シナリオで使用できる代替テキストを表します。|
||[属性](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member)|この画像を使用するソース要件とライセンス要件を説明する属性情報を表します。|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basictype-member)|この値を持つセルに対して返 `Range.valueTypes` される値を表します。|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basicvalue-member)|この値を持つセルに対して返 `Range.values` される値を表します。|
||[プロバイダー](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-provider-member)|画像を提供したエンティティまたは個人を表す情報を表します。|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-relatedimagesaddress-member)|このページに関連すると見なされる画像を含む Web ページの URL を表します `WebImageCellValue`。|
||[type](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-type-member)|このセル値の種類を表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getLinkedEntityCellValue(linkedEntityCellValueId: LinkedEntityId)](/javascript/api/excel/excel.workbook#excel-excel-workbook-getlinkedentitycellvalue-member(1))|指定した値に `LinkedEntityCellValue` 基づいて a を返します `LinkedEntityId`。|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|ブックの一部であるリンクされたデータ型のコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|ピボットテーブルのフィールド 一覧ウィンドウをブック レベルで表示するかどうかを指定します。|
||[タスク](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|ブックに存在するタスクのコレクションを返します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|特定のワークシートにフィルターが適用されると発生します。|
||[タスク](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|ワークシートに存在するタスクのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|フィルターが適用されるワークシートの ID を取得します。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|このワークシートにある `AllowEditRangeCollection` オブジェクトを指定します。|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|このワークシートの保護を一時停止できる場合に指定します。|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|ワークシート保護のロック解除にパスワードを使用できる場合に指定します。|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|シートがパスワードで保護される場合に指定します。|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|ワークシートの保護を一時停止する場合に指定します。|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|特定のセッションのユーザーに対する、指定されたワークシート オブジェクトのワークシート保護を一時停止します。|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|特定のセッションのユーザーに対する、指定されたワークシート オブジェクトのワークシート保護を再開します。|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|オブジェクトに関連付けられているパスワードを変更 `WorksheetProtection` します。|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|オブジェクトに関連付けられているワークシート保護オプションを変更 `WorksheetProtection` します。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|オブジェクトが変更された場合に `AllowEditRange` 指定します。|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|変更された場合に指定 `WorksheetProtectionOptions` します。|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|ワークシートのパスワードが変更された場合に指定します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
