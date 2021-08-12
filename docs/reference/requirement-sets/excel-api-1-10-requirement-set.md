---
title: ExcelJavaScript API 要件セット 1.10
description: ExcelApi 1.10 要件セットの詳細。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 34c21ad0e90593352ae4042c2be148e607c63164aac1845357e9f96371104f6f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087216"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>JavaScript API 1.10 Excel新機能

ExcelApi 1.10 では、コメント、アウトライン、スライサーなどの主要な機能が導入されました。 また、ワークシート レベルのクリックと並べ替えのイベントサポートも追加されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメント](../../excel/excel-add-ins-comments.md) | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [アウトライン](../../excel/excel-add-ins-ranges-group.md) | 折りたたみ可能なアウトラインを形成する行と列をグループ化します。 | [Range](/javascript/api/excel/excel.range)、 [Worksheet](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [その他のワークシート イベント](../../excel/excel-add-ins-events.md) | ワークシートでクリックイベントと並べ替えイベントをリッスンします。 | [ワークシート (イベント)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.10 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.10 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|コメントのコンテンツ。|
||[delete()](/javascript/api/excel/excel.comment#delete__)|コメントとすべての接続済み返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|このコメントがあるセルを取得します。|
||[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.comment#authorName)|コメント作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|コメントの作成日時を取得します。|
||[id](/javascript/api/excel/excel.comment#id)|コメント識別子を指定します。|
||[replies](/javascript/api/excel/excel.comment#replies)|コメントに関連付けられている返信オブジェクトのコレクションを表します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|指定したセルで、指定した内容の新しいコメントを作成します。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|ID に基づいてコレクションからコメントを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|位置に基づいてコレクションからコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|指定したセルからコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|指定した返信が接続されているコメントを取得します。|
||[items](/javascript/api/excel/excel.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|コメント返信のコンテンツ。|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|このコメント返信があるセルを取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|この返信の親コメントを取得します。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|コメント返信作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreply#id)|コメント返信識別子を指定します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|その ID で識別されるコメント返信を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|コレクション内の位置に基づいてコメント返信を取得します。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|フィールド リストを UI に表示できる場合に指定します。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|ピボットテーブル スタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|すべてのスタイル要素のコピーを含む、このピボットテーブル スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|ピボットテーブル スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|このオブジェクトが読 `PivotTableStyle` み取り専用の場合に指定します。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|指定した名前の `PivotTableStyle` 空白を作成します。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|コレクションに含まれる PivotTableStyle の数を取得します。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|親オブジェクトのスコープの既定のピボットテーブル スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|名前で `PivotTableStyle` 取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|名前で `PivotTableStyle` 取得します。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|親オブジェクトのスコープで使用する既定のピボットテーブル スタイルを設定します。|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group_groupOption_)|アウトラインの列と行をグループ分けします。|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|行または列グループの詳細を非表示にします。|
||[height](/javascript/api/excel/excel.range#height)|範囲の上端から範囲の下端までの 100% ズームの距離をポイントで返します。|
||[left](/javascript/api/excel/excel.range#left)|ワークシートの左側から範囲の左端までの距離をポイントで返します。100% ズームの場合。|
||[top](/javascript/api/excel/excel.range#top)|ワークシートの上端から範囲の上端までの 100% ズームの距離をポイントで返します。|
||[width](/javascript/api/excel/excel.range#width)|範囲の左端から範囲の右端までの距離をポイントで返します。100% ズームの場合。|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|行または列グループの詳細を表示します。|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup_groupOption_)|アウトラインの列と行のグループを解除します。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|オブジェクトをコピーして貼り付 `Shape` けます。|
||[placement](/javascript/api/excel/excel.shape#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|スライサーのキャプションを表します。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|現在スライサーに適用されているすべてのフィルターを消去します。|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|スライサーを削除します。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|選択されたアイテムのキーの配列を返します。|
||[height](/javascript/api/excel/excel.slicer#height)|スライサーの高さ (ポイント数) を表します。|
||[left](/javascript/api/excel/excel.slicer#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicer#name)|スライサーの名前を表します。|
||[id](/javascript/api/excel/excel.slicer#id)|スライサーの一意の ID を表します。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|値は `true` 、スライサーに現在適用されているフィルターすべてがクリアされている場合です。|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|スライサーの一部であるスライサー アイテムのコレクションを表します。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|スライサーを含んでいるワークシートを表します。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|キーに基づいてスライサー アイテムを選択します。|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|スライサーに含まれるアイテムの並べ替え順序を表します。|
||[style](/javascript/api/excel/excel.slicer#style)|スライサー スタイルを表す定数値。|
||[top](/javascript/api/excel/excel.slicer#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicer#width)|スライサーの幅 (ポイント数) を表します。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|ブックに新しいスライサーを追加します。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|コレクションに含まれるスライサーの数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|名前または ID を使用してスライサー オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|コレクション内の位置に基づいてスライサーを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|名前または ID を使用してスライサーを取得します。|
||[items](/javascript/api/excel/excel.slicercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|値は `true` 、スライサー アイテムが選択されている場合です。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasData)|値は `true` 、スライサー アイテムにデータがある場合です。|
||[key](/javascript/api/excel/excel.sliceritem#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritem#name)|ユーザー UI に表示されるタイトルExcelします。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|スライサーのスライサー アイテム数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|そのキーまたは名前を利用してスライサー アイテム オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|コレクション内の位置に基づいてスライサー アイテムを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|そのキーまたは名前を使用してスライサー アイテムを取得します。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|スライサー スタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|すべてのスタイル要素のコピーを使用して、このスライサー スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|スライサー スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|このオブジェクトが読 `SlicerStyle` み取り専用の場合に指定します。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|指定した名前の空白のスライサー スタイルを作成します。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|コレクション内のスライサー スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|親オブジェクトの `SlicerStyle` スコープの既定値を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|名前で `SlicerStyle` 取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|名前で `SlicerStyle` 取得します。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|親オブジェクトのスコープで使用する既定のスライサー スタイルを設定します。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|表のスタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|すべてのスタイル要素のコピーを含む、このテーブル スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.tablestyle#name)|テーブル スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|このオブジェクトが読 `TableStyle` み取り専用の場合に指定します。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|指定した名前の `TableStyle` 空白を作成します。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|コレクションに含まれるテーブル スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|親オブジェクトのスコープの既定のテーブル スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|名前で `TableStyle` 取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|名前で `TableStyle` 取得します。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|親オブジェクトのスコープで使用する既定のテーブル スタイルを設定します。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|表のスタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|すべてのスタイル要素のコピーを使用して、このタイムライン スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|タイムライン スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|このオブジェクトが読 `TimelineStyle` み取り専用の場合に指定します。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|指定した名前の `TimelineStyle` 空白を作成します。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|コレクションに含まれるタイムライン スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|親オブジェクトのスコープの既定のタイムライン スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|名前で `TimelineStyle` 取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|名前で `TimelineStyle` 取得します。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|親オブジェクトのスコープで使用する既定のタイムライン スタイルを設定します。|
|[ブック](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|ブックで現在アクティブになっているスライサーを取得します。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|ブックで現在アクティブになっているスライサーを取得します。|
||[comments](/javascript/api/excel/excel.workbook#comments)|ブックに関連付けられたコメントのコレクションを表します。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|ブックに関連付けられたスライサーのコレクションを表します。|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|ブックに関連付けられている TableStyle のコレクションを表します。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|1 つ以上の列を並べ替えたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|1 つ以上の行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|ワークシートで左クリック/タップ操作が行われると発生します。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|ワークシートの一部であるスライサーのコレクションを返します。|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|行または列のグループをアウトライン レベルで表示します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|1 つ以上の列を並べ替えたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|1 つ以上の行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|ワークシート コレクションで左クリック/タップ操作が実行された場合に発生します。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|並べ替えが行ったワークシートの ID を取得します。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|並べ替えが行ったワークシートの ID を取得します。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|特定のワークシートで左クリック/タップされたセルを表すアドレスを取得します。|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|左クリック/タップされたポイントから左クリック/タップされたセルの左 (または右から左の言語の場合は右) の枠線の端までの距離をポイントで指定します。|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|左クリック/タップされたポイントから、左クリック/タップされたセルの上側の目盛線までの距離を、ポイント単位で表します。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|セルが左クリック/タップされたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)