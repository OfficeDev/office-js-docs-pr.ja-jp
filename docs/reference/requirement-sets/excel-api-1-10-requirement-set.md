---
title: Excel JavaScript API 要件セット1.10
description: ExcelApi 1.10 の要件セットの詳細
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 890d198f238e29d39744d87d754381543ebcaf6a
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431235"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Excel JavaScript API 1.10 の新機能

ExcelApi 1.10 には、コメント、アウトライン、スライサーなどの主要な機能が導入されています。 また、ワークシートレベルのクリックと並べ替えのイベントサポートも追加されました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメント](../../excel/excel-add-ins-comments.md) | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [アウトライン](../../excel/excel-add-ins-ranges-advanced.md#group-data-for-an-outline) | 行と列をグループ化して、折りたたみ可能なアウトラインを作成します。 | [範囲](/javascript/api/excel/excel.range)、 [ワークシート](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#slicers) | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [その他のワークシートイベント](../../excel/excel-add-ins-events.md) | ワークシートでクリックして並べ替えイベントを待機します。 | [ワークシート (イベント)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.10 の Api を示します。 Excel JavaScript API 要件セット1.10 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.10 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|コメントの内容を取得または設定します。 文字列はテキスト形式です。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|コメントと、接続されているすべての返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|このコメントが配置されているセルを取得します。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|コメント作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|コメントの作成日時を取得します。 コメントがメモから変換されている場合、コメントには作成日時がないため、null が返されます。|
||[id](/javascript/api/excel/excel.comment#id)|コメント ID を表します。 読み取り専用です。|
||[replies](/javascript/api/excel/excel.comment#replies)|コメントに関連付けられている返信オブジェクトのコレクションを表します。 読み取り専用です。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress: Range \| string, content: CommentRichContent \| String, contenttype?: Excel)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|指定したセルで、指定した内容の新しいコメントを作成します。 `InvalidArgument`指定した範囲が1つのセルより大きい場合は、エラーがスローされます。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|ID に基づいてコレクションからコメントを取得します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|位置に基づいてコレクションからコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|指定したセルからコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|指定した返信が接続されているコメントを取得します。|
||[items](/javascript/api/excel/excel.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|このコメントの返信があるセルを取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|この返信の親コメントを取得します。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|コメント返信作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreply#id)|コメント返信 ID を表します。 読み取り専用です。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add (content: CommentRichContent \| string, contenttype?: Excel)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|その ID で識別されるコメント返信を返します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|コレクション内の位置に基づいてコメント返信を取得します。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)||[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|PivotTableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|すべてのスタイル要素のコピーでこの PivotTableStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|この PivotTableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|指定された名前で空の PivotTableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|コレクションに含まれる PivotTableStyle の数を取得します。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の PivotTableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|名前に基づいて PivotTableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|名前に基づいて PivotTableStyle を取得します。 PivotTableStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の PivotTableStyle を設定します。|
|[Range](/javascript/api/excel/excel.range)|[group (groupOption: Excel. groupoption](/javascript/api/excel/excel.range#group-groupoption-)|アウトラインの列と行をグループ化します。|
||[hideGroupDetails (groupopoff: Excel. groupopoff)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|行または列グループの詳細を非表示にします。|
||[height](/javascript/api/excel/excel.range#height)|100% ズームの場合の、範囲の上端から範囲の下端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[left](/javascript/api/excel/excel.range#left)|100% ズームの場合の、ワークシートの左端から範囲の左端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[top](/javascript/api/excel/excel.range#top)|100% ズームの場合の、ワークシートの上端から範囲の上端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[width](/javascript/api/excel/excel.range#width)|100% ズームの場合の、範囲の左端から範囲の右端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[showGroupDetails (groupopoff: Excel. groupopoff)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|行または列グループの詳細を表示します。|
||[グループ解除 (groupOption: Excel. groupoption](/javascript/api/excel/excel.range#ungroup-groupoption-)|アウトラインの列と行のグループ化を解除します。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Shape オブジェクトをコピーして貼り付けます。|
||[placement](/javascript/api/excel/excel.shape#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|スライサーのキャプションを表します。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|現在スライサーに適用されているすべてのフィルターを消去します。|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|スライサーを削除します。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|選択されたアイテムのキーの配列を返します。 読み取り専用です。|
||[height](/javascript/api/excel/excel.slicer#height)|スライサーの高さ (ポイント数) を表します。|
||[left](/javascript/api/excel/excel.slicer#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicer#name)|スライサーの名前を表します。|
||[id](/javascript/api/excel/excel.slicer#id)|スライサーの一意の ID を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|スライサーに現在適用されているフィルターがすべて消去されている場合、true となります。|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|スライサーに含まれる SlicerItems のコレクションを表します。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|スライサーを含んでいるワークシートを表します。 読み取り専用です。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|キーに基づいてスライサーアイテムを選択します。 以前の選択はクリアされます。|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 可能な値は、"DataSourceOrder"、"昇順"、"降順" です。|
||[style](/javascript/api/excel/excel.slicer#style)|スライサー スタイルを表す定数値。 可能な値は次のとおりです。 "SlicerStyleLight1" は "SlicerStyleLight6"、"TableStyleOther1" ~ "TableStyleOther2"、"SlicerStyleDark1" ~ "SlicerStyleDark6" です。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicer#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicer#width)|スライサーの幅 (ポイント数) を表します。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|ブックに新しいスライサーを追加します。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|コレクションに含まれるスライサーの数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|名前または ID を使用してスライサー オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|コレクション内の位置に基づいてスライサーを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|名前または ID に基づいてスライサーを取得します。スライサーが存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.slicercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|スライサー アイテムが選択されている場合、true となります。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|スライサー アイテムにデータが含まれている場合、true となります。|
||[key](/javascript/api/excel/excel.sliceritem#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritem#name)|UI に表示されるタイトルを表します。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|スライサーのスライサー アイテム数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|そのキーまたは名前を利用してスライサー アイテム オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|コレクション内の位置に基づいてスライサー アイテムを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|そのキーまたは名前を使用してスライサー アイテムを取得します。 スライサー アイテムが存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|SlicerStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|すべてのスタイル要素のコピーでこの SlicerStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|この SlicerStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|指定された名前で空の SlicerStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|コレクション内のスライサー スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|親オブジェクトのスコープに対する既定の SlicerStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|名前で SlicerStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|名前で SlicerStyle を取得します。 SlicerStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の SlicerStyle を設定します。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|TableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|すべてのスタイル要素のコピーでこの TableStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.tablestyle#name)|TableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|この TableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|指定された名前で空の TableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|コレクションに含まれるテーブル スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|名前で TableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|名前で TableStyle を取得します。 TableStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TableStyle を設定します。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|TableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|すべてのスタイル要素のコピーでこの TimelineStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|この TimelineStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|指定された名前で空の TimelineStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|コレクションに含まれるタイムライン スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TimelineStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|名前で TimelineStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|名前で TimelineStyle を取得します。 TimelineStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TimelineStyle を設定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブなスライサーがない場合は、 `ItemNotFound` 例外がスローされます。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブになっているスライサーがない場合、null オブジェクトが返されます。|
||[comments](/javascript/api/excel/excel.workbook#comments)|ブックに関連付けられているコメントの集まりを表します。 読み取り専用です。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。 読み取り専用です。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|ブックに関連付けられているスライサーの集まりを表します。 読み取り専用です。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|ブックに関連付けられている TableStyle のコレクションを表します。 読み取り専用です。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。 読み取り専用です。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。 読み取り専用です。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|1 つ以上の列を並べ替えたときに発生します。 これは、左から右に並べ替えを実行したときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|1 つ以上の行を並べ替えたときに発生します。 これは、上から下に並べ替えを実行したときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|左クリック/タップ操作がワークシートで発生したときに発生します。 このイベントは、次のケースをクリックしても発生しません。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|ワークシートの一部であるスライサーのコレクションを返します。 読み取り専用です。|
||[showOutlineLevels (rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|アウトラインレベルで行または列のグループを表示します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|1 つ以上の列を並べ替えたときに発生します。 これは、左から右に並べ替えを実行したときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|1 つ以上の行を並べ替えたときに発生します。 これは、上から下に並べ替えを実行したときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|ワークシートのコレクションで左クリック/タップ操作が行われるときに発生します。 このイベントは、次のケースをクリックしても発生しません。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。 並べ替え操作の結果として変更された列のみが返されます。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。 並べ替え操作の結果として変更された行のみが返されます。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|特定のワークシートで左クリック/タップされたセルを表すアドレスを取得します。|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|左クリックまたは左にクリックされたポイント (右から左へ記述する言語の場合は右) からの距離をポイント単位で指定します。左クリック/タップしたセルの枠線の端点を指定します。|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|左クリック/タップされたポイントから、左クリック/タップされたセルの上側の目盛線までの距離を、ポイント単位で表します。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|左クリック/タップされたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)