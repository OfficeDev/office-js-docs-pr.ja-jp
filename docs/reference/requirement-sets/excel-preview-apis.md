---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 08/15/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c0b0a5757df2f5766f0e29b599e3e9b36dfd3983
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477908"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライサー](../../excel/excel-add-ins-pivottables.md#slicers-preview) | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [コメント](../../excel/excel-add-ins-workbooks.md#comments-preview) | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | ブックを保存して閉じます。  | [Workbook](/javascript/api/excel/excel.workbook) |
| [ブックを挿入する](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「[すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|コメントの内容を取得または設定します。 文字列はテキスト形式です。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|コメント スレッドを削除します。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|このコメントが配置されているセルを取得します。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|コメント作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|コメントの作成日時を取得します。 コメントがメモから変換されている場合、コメントには作成日時がないため、null が返されます。|
||[id](/javascript/api/excel/excel.comment#id)|コメント ID を表します。 読み取り専用です。|
||[replies](/javascript/api/excel/excel.comment#replies)|コメントに関連付けられている返信オブジェクトのコレクションを表します。 読み取り専用です。|
||[解析](/javascript/api/excel/excel.comment#resolved)|コメントスレッドの状態を取得または設定します。 値 "true" は、コメントスレッドが解決された状態であることを意味します。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|指定したセルに、指定されたコンテンツを含む新しいコメント (コメントスレッド) を作成します。 指定`InvalidArgument`した範囲が1つのセルより大きい場合は、エラーがスローされます。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|ID に基づいてコレクションからコメントを取得します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|位置に基づいてコレクションからコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|指定したセルからコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|コレクション内のその返信 ID に関連付けられているコメントを取得します。|
||[items](/javascript/api/excel/excel.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|このコメントの返信があるセルを取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|この返信の親コメントを取得します。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|コメント返信作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreply#id)|コメント返信 ID を表します。 読み取り専用です。|
||[解析](/javascript/api/excel/excel.commentreply#resolved)|コメントの返信状態を取得または設定します。 値 "true" は、コメント応答が解決された状態であることを意味します。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|その ID で識別されるコメント返信を返します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|コレクション内の位置に基づいてコメント返信を取得します。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
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
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 読み取り専用です。|
||[group (groupOption: Excel. groupoption](/javascript/api/excel/excel.range#group-groupoption-)|アウトラインの列と行をグループ化します。|
||[hideGroupDetails (groupopoff: Excel. groupopoff)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|行または列グループの詳細を非表示にします。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[height](/javascript/api/excel/excel.range#height)|100% ズームの場合の、範囲の上端から範囲の下端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[left](/javascript/api/excel/excel.range#left)|100% ズームの場合の、ワークシートの左端から範囲の左端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
||[top](/javascript/api/excel/excel.range#top)|100% ズームの場合の、ワークシートの上端から範囲の上端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[width](/javascript/api/excel/excel.range#width)|100% ズームの場合の、範囲の左端から範囲の右端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[showGroupDetails (groupopoff: Excel. groupopoff)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|行または列グループの詳細を表示します。|
||[グループ解除 (groupOption: Excel. groupoption](/javascript/api/excel/excel.range#ungroup-groupoption-)|アウトラインの列と行のグループ化を解除します。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Shape オブジェクトをコピーして貼り付けます。|
||[placement](/javascript/api/excel/excel.shape#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|スライサーのキャプションを表します。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|現在スライサーに適用されているすべてのフィルターを消去します。|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|スライサーを削除します。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|選択されたアイテムのキーの配列を返します。 読み取り専用です。|
||[height](/javascript/api/excel/excel.slicer#height)|スライサーの高さ (ポイント数) を表します。|
||[left](/javascript/api/excel/excel.slicer#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicer#name)|スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[id](/javascript/api/excel/excel.slicer#id)|スライサーの一意の ID を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|スライサーに現在適用されているフィルターがすべて消去されている場合、true となります。|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|スライサーに含まれる SlicerItems のコレクションを表します。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|スライサーを含んでいるワークシートを表します。 読み取り専用です。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|キーに基づいてスライサー アイテムを選択します。 前の選択は消去されます。|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
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
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されたテーブルの ID を表します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルが含まれるワークシートの ID を表します。|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブなスライサーがない場合は、 `ItemNotFound`例外がスローされます。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブになっているスライサーがない場合、null オブジェクトが返されます。|
||[comments](/javascript/api/excel/excel.workbook#comments)|ブックに関連付けられているコメントの集まりを表します。 読み取り専用です。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。 読み取り専用です。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|ブックに関連付けられているスライサーの集まりを表します。 読み取り専用です。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|ブックに関連付けられている TableStyle のコレクションを表します。 読み取り専用です。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。 読み取り専用です。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。 読み取り専用です。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|列を並べ替えたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで行の非表示の状態が変更されたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|ワークシートで左クリック/タップしたときに発生します。 このイベントは、次の場合にクリックしても発生しません。 [...]|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|ワークシートに含まれるスライサーをまとめて返します。 読み取り専用です。|
||[showOutlineLevels (rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|アウトラインレベルで行または列のグループを表示します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|列を並べ替えたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|ブック内のすべてのワークシートの行の非表示状態が変更されたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|ワークシートのコレクションで左クリック/タップ操作が行われるときに発生します。 このイベントは、次の場合にクリックしても発生しません。 [...]|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されたワークシートの ID を表します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|RowHiddenChanged イベントがトリガーされる方法を表す変更の種類を取得します。 詳細については、「RowHiddenChangeType」を参照してください。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|特定のワークシートで左クリック/タップされたセルを表すアドレスを取得します。|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|左クリックまたは左にクリックされたポイント (右から左へ記述する言語の場合は右) からの距離をポイント単位で指定します。左クリック/タップしたセルの枠線の端点を指定します。|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|左クリック/タップされたポイントから、左クリック/タップされたセルの上側の目盛線までの距離を、ポイント単位で表します。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|左クリック/タップされたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
