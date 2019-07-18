---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2199b7c115a1edd66bb7b1fef86eb3bc7bba473e
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771954"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには、CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)) で**ベータ** ライブラリを参照する必要があります。場合によっては、Office Insider プログラムに参加し、新しい Office ビルドを入手する必要があります。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [スライサー](../../excel/excel-add-ins-pivottables.md#slicers-preview) | テーブルやピボットテーブルにスライサーを挿入し、構成します。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [コメント](../../excel/excel-add-ins-workbooks.md#comments-preview) | コメントを追加、編集、削除します。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | ブックを保存して閉じます。  | [Workbook](/javascript/api/excel/excel.workbook) |
| [ブックを挿入する](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API リスト

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
||[set (properties: Excel. Comment)](/javascript/api/excel/excel.comment#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: CommentUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.comment#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|指定したセルに、指定されたコンテンツを含む新しいコメント (コメントスレッド) を作成します。 指定`InvalidArgument`した範囲が1つのセルより大きい場合は、エラーがスローされます。|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|指定したセルに、指定されたコンテンツを含む新しいコメント (コメントスレッド) を作成します。 指定`InvalidArgument`した範囲が1つのセルより大きい場合は、エラーがスローされます。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|コレクションに含まれるコメントの数を取得します。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|ID に基づいてコレクションからコメントを取得します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|位置に基づいてコレクションからコメントを取得します。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|指定したセルからコメントを取得します。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|コレクション内のその返信 ID に関連付けられているコメントを取得します。|
||[items](/javascript/api/excel/excel.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[コメントコレクションデータ](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[コメント Collectionloadoptions](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|コレクション内の各アイテムについて: コメントの作成者の電子メールを取得します。|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|コレクション内の各アイテムについて: コメントの作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|コレクション内の各アイテムについて: コメントの内容を取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|コレクション内の各アイテムについて: コメントの作成時刻を取得します。 コメントがメモから変換されている場合、コメントには作成日時がないため、null が返されます。|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|コレクション内の各アイテムについて: コメント識別子を表します。 読み取り専用です。|
|[CommentCollectionUpdateData](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[コメントデータ](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|コメント作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentdata#content)|コメントの内容を取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|コメントの作成日時を取得します。 コメントがメモから変換されている場合、コメントには作成日時がないため、null が返されます。|
||[id](/javascript/api/excel/excel.commentdata#id)|コメント ID を表します。 読み取り専用です。|
||[replies](/javascript/api/excel/excel.commentdata#replies)|コメントに関連付けられている返信オブジェクトのコレクションを表します。 読み取り専用です。|
|[コメント Loadoptions](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|コメント作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|コメントの内容を取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|コメントの作成日時を取得します。 コメントがメモから変換されている場合、コメントには作成日時がないため、null が返されます。|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|コメント ID を表します。 読み取り専用です。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|コメント返信を削除します。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|このコメントの返信があるセルを取得します。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|この返信の親コメントを取得します。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|コメント返信作成者の名前を取得します。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreply#id)|コメント返信 ID を表します。 読み取り専用です。|
||[set (プロパティ: Excel! 返信)](/javascript/api/excel/excel.commentreply#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: CommentReplyUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.commentreply#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|コメントのコメント返信を作成します。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|コレクションのコメント返信数を取得します。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|その ID で識別されるコメント返信を返します。 読み取り専用です。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|コレクション内の位置に基づいてコメント返信を取得します。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReplyCollectionData](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[CommentReplyCollectionLoadOptions](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|コレクション内の各アイテムについて: コメントの返信の作成者の電子メールを取得します。|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|コレクション内の各アイテムについて: コメントの返信の作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|コレクション内の各アイテムについて: コメント応答のコンテンツを取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|コレクション内の各アイテムについて: コメント応答の作成時刻を取得します。|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|コレクション内の各アイテムについて: コメント応答識別子を表します。 読み取り専用です。|
|[CommentReplyCollectionUpdateData](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[CommentReplyData](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|コメント返信作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentreplydata#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreplydata#id)|コメント返信 ID を表します。 読み取り専用です。|
|[CommentReplyLoadOptions](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|コメント返信作成者の名前を取得します。|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|コメント返信の作成日時を取得します。|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|コメント返信 ID を表します。 読み取り専用です。|
|[CommentReplyUpdateData](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|コメント返信の内容を取得または設定します。 文字列はテキスト形式です。|
|[CommentUpdateData](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|コメントの内容を取得または設定します。 文字列はテキスト形式です。|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|コレクション内の各アイテムについて: オブジェクトがその下のセルにどのように接続されるかを表します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|フィールド リストを UI に表示できるかどうかを指定します。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|PivotTableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|すべてのスタイル要素のコピーでこの PivotTableStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|この PivotTableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
||[set (properties: PivotTableStyle)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotTableStyleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|指定された名前で空の PivotTableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|コレクションに含まれる PivotTableStyle の数を取得します。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の PivotTableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|名前に基づいて PivotTableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|名前に基づいて PivotTableStyle を取得します。 PivotTableStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の PivotTableStyle を設定します。|
|[PivotTableStyleCollectionData](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[PivotTableStyleCollectionLoadOptions](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|コレクション内の各アイテムについて: PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|コレクション内の各アイテムについて: この PivotTableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[PivotTableStyleCollectionUpdateData](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[ピボットのスタイルデータ](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|この PivotTableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[ピボットピボットスタイル Loadoptions](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|PivotTableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|この PivotTableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[PivotTableStyleUpdateData](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|PivotTableStyle の名前を取得します。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 読み取り専用です。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[height](/javascript/api/excel/excel.range#height)|100% ズームの場合の、範囲の上端から範囲の下端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[left](/javascript/api/excel/excel.range#left)|100% ズームの場合の、ワークシートの左端から範囲の左端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
||[top](/javascript/api/excel/excel.range#top)|100% ズームの場合の、ワークシートの上端から範囲の上端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[width](/javascript/api/excel/excel.range#width)|100% ズームの場合の、範囲の左端から範囲の右端までの距離を、ポイント単位で返します。  読み取り専用です。|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|コレクション内の各アイテムについて: すべてのセルにスピル境界線があるかどうかを表します。|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|コレクション内の各項目に対して、範囲の上端から下端までの距離をポイント単位で 100% ズームで返します。 読み取り専用です。|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|コレクション内の各項目について、次のようにします。 100% ズームの場合は、ワークシートの左端から範囲の左端までの距離をポイント単位で返します。 読み取り専用です。|
||[savedAsArray](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|コレクション内の各アイテムについて: すべてのセルが配列数式として保存されるかどうかを表します。|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|コレクション内の各項目について、次のように、ワークシートの上端から範囲の上端までの距離をポイント単位で 100% ズームで返します。 読み取り専用です。|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|コレクション内の各項目について、範囲の左端から右端までの距離をポイント単位で 100% ズームで返します。 読み取り専用です。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[height](/javascript/api/excel/excel.rangedata#height)|100% ズームの場合の、範囲の上端から範囲の下端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[left](/javascript/api/excel/excel.rangedata#left)|100% ズームの場合の、ワークシートの左端から範囲の左端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[savedAsArray](/javascript/api/excel/excel.rangedata#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
||[top](/javascript/api/excel/excel.rangedata#top)|100% ズームの場合の、ワークシートの上端から範囲の上端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[width](/javascript/api/excel/excel.rangedata#width)|100% ズームの場合の、範囲の左端から範囲の右端までの距離を、ポイント単位で返します。  読み取り専用です。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|100% ズームの場合の、範囲の上端から範囲の下端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|100% ズームの場合の、ワークシートの左端から範囲の左端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[savedAsArray](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|100% ズームの場合の、ワークシートの上端から範囲の上端までの距離を、ポイント単位で返します。  読み取り専用です。|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|100% ズームの場合の、範囲の左端から範囲の右端までの距離を、ポイント単位で返します。  読み取り専用です。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Shape オブジェクトをコピーして貼り付けます。|
||[placement](/javascript/api/excel/excel.shape#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|コレクション内の各アイテムについて: オブジェクトがその下のセルにどのように接続されるかを表します。|
|[図形データ](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[図形 Loadoptions](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|オブジェクトがその下のセルに接続されている方法を表します。|
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
||[set (properties: Excel. スライサー)](/javascript/api/excel/excel.slicer#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: SlicerUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.slicer#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
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
|[SlicerCollectionData](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[SlicerCollectionLoadOptions](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|コレクション内の各アイテムについて: スライサーのキャプションを表します。|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|コレクション内の各アイテムについて: スライサーの高さをポイント単位で表します。|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|コレクション内の各アイテムについて: スライサーの一意の id を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|コレクション内の各アイテムについて: True を指定すると、スライサーに現在適用されているすべてのフィルターがクリアされます。|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|コレクション内の各項目の場合: ワークシートの左側にあるスライサーの左側からの距離をポイント単位で表します。|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|コレクション内の各アイテムについて: スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|コレクション内の各アイテムについて: 数式で使用されるスライサー名を表します。|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|コレクション内の各アイテムについて: スライサー内のアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|コレクション内の各項目について: スライサースタイルを表す定数値。 可能な値は次のとおりです。 "SlicerStyleLight1" は "SlicerStyleLight6"、"TableStyleOther1" ~ "TableStyleOther2"、"SlicerStyleDark1" ~ "SlicerStyleDark6" です。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|コレクション内の各項目について、: スライサーの上端からワークシートの上端までの距離をポイント単位で表します。|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|コレクション内の各アイテムについて: スライサーの幅をポイント単位で表します。|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|コレクション内の各アイテムについて: スライサーを含むワークシートを表します。|
|[SlicerCollectionUpdateData](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[SlicerData](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|スライサーのキャプションを表します。|
||[height](/javascript/api/excel/excel.slicerdata#height)|スライサーの高さ (ポイント数) を表します。|
||[id](/javascript/api/excel/excel.slicerdata#id)|スライサーの一意の ID を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|スライサーに現在適用されているフィルターがすべて消去されている場合、true となります。|
||[left](/javascript/api/excel/excel.slicerdata#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicerdata#name)|スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|数式で使用するスライサーの名前を表します。|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|スライサーに含まれる SlicerItems のコレクションを表します。 読み取り専用です。|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
||[style](/javascript/api/excel/excel.slicerdata#style)|スライサー スタイルを表す定数値。 可能な値は次のとおりです。 "SlicerStyleLight1" は "SlicerStyleLight6"、"TableStyleOther1" ~ "TableStyleOther2"、"SlicerStyleDark1" ~ "SlicerStyleDark6" です。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicerdata#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicerdata#width)|スライサーの幅 (ポイント数) を表します。|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|スライサーを含んでいるワークシートを表します。 読み取り専用です。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|スライサー アイテムが選択されている場合、true となります。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|スライサー アイテムにデータが含まれている場合、true となります。|
||[key](/javascript/api/excel/excel.sliceritem#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritem#name)|UI に表示されるタイトルを表します。|
||[set (properties: SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: SlicerItemUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|スライサーのスライサー アイテム数を返します。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|そのキーまたは名前を利用してスライサー アイテム オブジェクトを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|コレクション内の位置に基づいてスライサー アイテムを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|そのキーまたは名前を使用してスライサー アイテムを取得します。 スライサー アイテムが存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlicerItemCollectionData](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[SlicerItemCollectionLoadOptions](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|コレクション内の各アイテムについて: True の場合は、スライサーアイテムにデータが含まれています。|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|コレクション内の各アイテムに対して、スライサーアイテムが選択されている場合は True。|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|コレクション内の各アイテムについて: スライサーアイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|コレクション内の各アイテムについて: UI に表示されるタイトルを表します。|
|[SlicerItemCollectionUpdateData](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[SlicerItemData](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|スライサー アイテムにデータが含まれている場合、true となります。|
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|スライサー アイテムが選択されている場合、true となります。|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|UI に表示されるタイトルを表します。|
|[SlicerItemLoadOptions](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|スライサー アイテムにデータが含まれている場合、true となります。|
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|スライサー アイテムが選択されている場合、true となります。|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|スライサー アイテムを表す一意の値を表します。|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|UI に表示されるタイトルを表します。|
|[SlicerItemUpdateData](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|スライサー アイテムが選択されている場合、true となります。|
|[SlicerLoadOptions](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|スライサーのキャプションを表します。|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|スライサーの高さ (ポイント数) を表します。|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|スライサーの一意の ID を表します。 読み取り専用です。|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|スライサーに現在適用されているフィルターがすべて消去されている場合、true となります。|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|数式で使用するスライサーの名前を表します。|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|スライサー スタイルを表す定数値。 可能な値は次のとおりです。 "SlicerStyleLight1" は "SlicerStyleLight6"、"TableStyleOther1" ~ "TableStyleOther2"、"SlicerStyleDark1" ~ "SlicerStyleDark6" です。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|スライサーの幅 (ポイント数) を表します。|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|スライサーを含んでいるワークシートを表します。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|SlicerStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|すべてのスタイル要素のコピーでこの SlicerStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|この SlicerStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
||[set (properties: SlicerStyle)](/javascript/api/excel/excel.slicerstyle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: SlicerStyleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|指定された名前で空の SlicerStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|コレクション内のスライサー スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|親オブジェクトのスコープに対する既定の SlicerStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|名前で SlicerStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|名前で SlicerStyle を取得します。 SlicerStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の SlicerStyle を設定します。|
|[SlicerStyleCollectionData](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[SlicerStyleCollectionLoadOptions](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|コレクション内の各アイテムについて: SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|コレクション内の各アイテムについて: この SlicerStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[SlicerStyleCollectionUpdateData](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[SlicerStyleData](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|この SlicerStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[SlicerStyleLoadOptions](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|SlicerStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|この SlicerStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[SlicerStyleUpdateData](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|SlicerStyle の名前を取得します。|
|[SlicerUpdateData](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|スライサーのキャプションを表します。|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|スライサーの高さ (ポイント数) を表します。|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|スライサーの左側からワークシートの左までの距離を表します (ポイント数)。|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|スライサーの名前を表します。|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|数式で使用するスライサーの名前を表します。|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|スライサーに含まれるアイテムの並べ替え順序を表します。 指定可能な値は DataSourceOrder、Ascending、Descending です。|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|スライサー スタイルを表す定数値。 可能な値は次のとおりです。 "SlicerStyleLight1" は "SlicerStyleLight6"、"TableStyleOther1" ~ "TableStyleOther2"、"SlicerStyleDark1" ~ "SlicerStyleDark6" です。 ブックに存在するカスタムのユーザー定義スタイルも指定できます。|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|スライサーの上端からワークシートの上端までの距離を表します (ポイント数)。|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|スライサーの幅 (ポイント数) を表します。|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|スライサーを含んでいるワークシートを表します。|
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
||[set (properties: TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TableStyleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|指定された名前で空の TableStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|コレクションに含まれるテーブル スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TableStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|名前で TableStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|名前で TableStyle を取得します。 TableStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TableStyle を設定します。|
|[TableStyleCollectionData](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[TableStyleCollectionLoadOptions](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|コレクション内の各アイテムについて: TableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|コレクション内の各アイテムについて: この TableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TableStyleCollectionUpdateData](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[Tableのスタイルデータ](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|TableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|この TableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[Tableスタイル Loadoptions](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|TableStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|この TableStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TableStyleUpdateData](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|TableStyle の名前を取得します。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|TableStyle を削除します。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|すべてのスタイル要素のコピーでこの TimelineStyle の複製を作成します。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|この TimelineStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
||[set (properties: TimelineStyle)](/javascript/api/excel/excel.timelinestyle#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: TimelineStyleUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|指定された名前で空の TimelineStyle を作成します。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|コレクションに含まれるタイムライン スタイルの数を取得します。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|親オブジェクトのスコープに対する既定の TimelineStyle を取得します。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|名前で TimelineStyle を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|名前で TimelineStyle を取得します。 TimelineStyle が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|親オブジェクトのスコープで使用する既定の TimelineStyle を設定します。|
|[TimelineStyleCollectionData](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[TimelineStyleCollectionLoadOptions](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|コレクション内の各アイテムについて: TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|コレクション内の各アイテムについて: この TimelineStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TimelineStyleCollectionUpdateData](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[TimelineStyleData](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|この TimelineStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TimelineStyleLoadOptions](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|TimelineStyle の名前を取得します。|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|この TimelineStyle オブジェクトが読み取り専用であるかどうかを指定します。 読み取り専用です。|
|[TimelineStyleUpdateData](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|TimelineStyle の名前を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブなスライサーがない場合は、 `ItemNotFound`例外がスローされます。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|ブックで現在アクティブになっているスライサーを取得します。 アクティブになっているスライサーがない場合、null オブジェクトが返されます。|
||[comments](/javascript/api/excel/excel.workbook#comments)|ブックに関連付けられているコメントの集まりを表します。 読み取り専用です。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。 読み取り専用です。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|ブックに関連付けられているスライサーの集まりを表します。 読み取り専用です。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|ブックに関連付けられている TableStyle のコレクションを表します。 読み取り専用です。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。 読み取り専用です。|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|ブックに関連付けられているコメントの集まりを表します。 読み取り専用です。|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|ブックに関連付けられている PivotTableStyle のコレクションを表します。 読み取り専用です。|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|ブックに関連付けられている SlicerStyle のコレクションを表します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.workbookdata#slicers)|ブックに関連付けられているスライサーの集まりを表します。 読み取り専用です。|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|ブックに関連付けられている TableStyle のコレクションを表します。 読み取り専用です。|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|ブックに関連付けられている TimelineStyle のコレクションを表します。 読み取り専用です。|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。 読み取り専用です。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|列を並べ替えたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで行の非表示の状態が変更されたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|ワークシートで左クリック/タップしたときに発生します。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|ワークシートに含まれるスライサーをまとめて返します。 読み取り専用です。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|列を並べ替えたときに発生します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|ブック内のすべてのワークシートの行の非表示状態が変更されたときに発生します。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|行を並べ替えたときに発生します。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|ワークシートのコレクションで左クリック/タップ操作が行われるときに発生します。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|ワークシート上のすべての Comments オブジェクトの集まりを返します。 読み取り専用です。|
||[slicers](/javascript/api/excel/excel.worksheetdata#slicers)|ワークシートに含まれるスライサーをまとめて返します。 読み取り専用です。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されたワークシートの ID を表します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Changed イベントがトリガーされる方法を表す変更の種類を取得します。 詳細については、「RowHiddenChangeType」を参照してください。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|特定のワークシートで並べ替えられたエリアを表す範囲のアドレスを取得します。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|並べ替えが発生したワークシートの ID を取得します。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|特定のワークシートで左クリック/タップされたセルを表すアドレスを取得します。|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|左クリック/タップされたポイントから、左クリック/タップされたセルの左側 (RTL の場合は右側) の目盛線までの距離を、ポイント単位で表します。|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|左クリック/タップされたポイントから、左クリック/タップされたセルの上側の目盛線までの距離を、ポイント単位で表します。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|左クリック/タップされたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
