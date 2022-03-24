---
title: Excel JavaScript API 要件セット 1.10
description: ExcelApi 1.10 要件セットの詳細。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 53cf0ec55a26f02a615a3c5eee0b718b818790d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746341"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>JavaScript API 1.10 Excel新機能

ExcelApi 1.10 では、コメント、アウトライン、スライサーなどの主要な機能が導入されました。 また、ワークシート レベルのクリックと並べ替えのイベントサポートも追加されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメント](../../excel/excel-add-ins-comments.md) | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [アウトライン](../../excel/excel-add-ins-ranges-group.md) | 折りたたみ可能なアウトラインを形成する行と列をグループ化します。 | [範囲](/javascript/api/excel/excel.range)、 [ワークシート](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [その他のワークシート イベント](../../excel/excel-add-ins-events.md) | ワークシートでクリックイベントと並べ替えイベントをリッスンします。 | [ワークシート (イベント)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.10 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.10 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|コメント作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|コメントのコンテンツ。|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|コメントの作成日時を取得します。|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|コメントとすべての接続済み返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|このコメントがあるセルを取得します。|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|コメント識別子を指定します。|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|コメントに関連付けられている返信オブジェクトのコレクションを表します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|指定したセルで、指定した内容の新しいコメントを作成します。|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|ID に基づいてコレクションからコメントを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|位置に基づいてコレクションからコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|指定したセルからコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|指定した返信が接続されているコメントを取得します。|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|コメント返信作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|コメント返信のコンテンツ。|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|コメント返信の作成日時を取得します。|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|このコメント返信があるセルを取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|この返信の親コメントを取得します。|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|コメント返信識別子を指定します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|その ID で識別されるコメント返信を返します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|コレクション内の位置に基づいてコメント返信を取得します。|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|フィールド リストを UI に表示できる場合に指定します。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|ピボットテーブル スタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|すべてのスタイル要素のコピーを含む、このピボットテーブル スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|ピボットテーブル スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|このオブジェクトが読み取 `PivotTableStyle` り専用の場合に指定します。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|指定した名前の空白 `PivotTableStyle` を作成します。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|コレクションに含まれる PivotTableStyle の数を取得します。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|親オブジェクトのスコープの既定のピボットテーブル スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|名前で取得 `PivotTableStyle` します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|名前で取得 `PivotTableStyle` します。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|親オブジェクトのスコープで使用する既定のピボットテーブル スタイルを設定します。|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|アウトラインの列と行をグループ分けします。|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|範囲の上端から範囲の下端までの 100% ズームの距離をポイントで返します。|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|行または列グループの詳細を非表示にします。|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|ワークシートの左側から範囲の左端までの距離をポイントで返します。100% ズームの場合。|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|行または列グループの詳細を表示します。|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|ワークシートの上端から範囲の上端までの 100% ズームの距離をポイントで返します。|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|アウトラインの列と行のグループを解除します。|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|範囲の左端から範囲の右端までの距離をポイントで返します。100% ズームの場合。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|オブジェクトをコピーして貼り付 `Shape` けます。|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|オブジェクトがその下のセルに接続されている方法を表します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|スライサーのキャプションを表します。|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|現在スライサーに適用されているすべてのフィルターを消去します。|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|スライサーを削除します。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|選択されたアイテムのキーの配列を返します。|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|スライサーの高さ (ポイント数) を表します。|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|スライサーの一意の ID を表します。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|値は、 `true` スライサーに現在適用されているフィルターすべてがクリアされている場合です。|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|スライサーの名前を表します。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|キーに基づいてスライサー アイテムを選択します。|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|スライサーの一部であるスライサー アイテムのコレクションを表します。|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|スライサーに含まれるアイテムの並べ替え順序を表します。|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|スライサー スタイルを表す定数値。|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|スライサーの幅 (ポイント数) を表します。|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|スライサーを含んでいるワークシートを表します。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|ブックに新しいスライサーを追加します。|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|コレクションに含まれるスライサーの数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|名前または ID を使用してスライサー オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|コレクション内の位置に基づいてスライサーを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|名前または ID を使用してスライサーを取得します。|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|値は、 `true` スライサー アイテムにデータがある場合です。|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|値は、 `true` スライサー アイテムが選択されている場合です。|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|ユーザー UI に表示されるタイトルExcelします。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|スライサーのスライサー アイテム数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|そのキーまたは名前を利用してスライサー アイテム オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|コレクション内の位置に基づいてスライサー アイテムを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|そのキーまたは名前を使用してスライサー アイテムを取得します。|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|スライサー スタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|すべてのスタイル要素のコピーを使用して、このスライサー スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|スライサー スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|このオブジェクトが読み取 `SlicerStyle` り専用の場合に指定します。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|指定した名前の空白のスライサー スタイルを作成します。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|コレクション内のスライサー スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|親オブジェクトの `SlicerStyle` スコープの既定値を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|名前で取得 `SlicerStyle` します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|名前で取得 `SlicerStyle` します。|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|親オブジェクトのスコープで使用する既定のスライサー スタイルを設定します。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|表のスタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|すべてのスタイル要素のコピーを含む、このテーブル スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|テーブル スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|このオブジェクトが読み取 `TableStyle` り専用の場合に指定します。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|指定した名前の空白 `TableStyle` を作成します。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|コレクションに含まれるテーブル スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|親オブジェクトのスコープの既定のテーブル スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|名前で取得 `TableStyle` します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|名前で取得 `TableStyle` します。|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|親オブジェクトのスコープで使用する既定のテーブル スタイルを設定します。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|表のスタイルを削除します。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|すべてのスタイル要素のコピーを使用して、このタイムライン スタイルの複製を作成します。|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|タイムライン スタイルの名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|このオブジェクトが読み取 `TimelineStyle` り専用の場合に指定します。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|指定した名前の空白 `TimelineStyle` を作成します。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|コレクションに含まれるタイムライン スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|親オブジェクトのスコープの既定のタイムライン スタイルを取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|名前で取得 `TimelineStyle` します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|名前で取得 `TimelineStyle` します。|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|親オブジェクトのスコープで使用する既定のタイムライン スタイルを設定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|ブックに関連付けられたコメントのコレクションを表します。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|ブックで現在アクティブになっているスライサーを取得します。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|ブックで現在アクティブになっているスライサーを取得します。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|ブックに関連付けられている PivotTableStyle のコレクションを表します。|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|ブックに関連付けられている SlicerStyle のコレクションを表します。|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|ブックに関連付けられたスライサーのコレクションを表します。|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|ブックに関連付けられている TableStyle のコレクションを表します。|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|ブックに関連付けられている TimelineStyle のコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|ワークシート上のすべての Comments オブジェクトの集まりを返します。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|1 つ以上の列を並べ替えたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|1 つ以上の行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|ワークシートで左クリック/タップ操作が行われると発生します。|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|行または列のグループをアウトライン レベルで表示します。|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|ワークシートの一部であるスライサーのコレクションを返します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|1 つ以上の列を並べ替えたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|1 つ以上の行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|ワークシート コレクションで左クリック/タップ操作が実行された場合に発生します。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|並べ替えが行ったワークシートの ID を取得します。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|イベントのソースを取得します。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|並べ替えが行ったワークシートの ID を取得します。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|特定のワークシートで左クリック/タップされたセルを表すアドレスを取得します。|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|左クリック/タップされたポイントから左クリック/タップされたセルの左 (または右から左の言語の場合は右) の枠線の端までの距離をポイントで指定します。|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|左クリック/タップされたポイントから、左クリック/タップされたセルの上側の目盛線までの距離を、ポイント単位で表します。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|セルが左クリック/タップされたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)