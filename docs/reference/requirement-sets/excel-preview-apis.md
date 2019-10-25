---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: dc0a2a3b23fbf4ccffb5de3b0689b0de0ed08b75
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682544"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [コメントメンション](../../excel/excel-add-ins-comments.md#mentions-preview) | 通知を送信するコメントに他のユーザーについて言及します。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [ブックを挿入する](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | あるブックを別のブックに挿入します。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| ブックを[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)して[閉じる](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | ブックを保存して閉じます。  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「[すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|グラフの系列の1つの次元から値を取得します。 指定できるのは、指定された次元と、グラフ系列に対するデータのマッピング方法によって異なります。|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|コメントに記載されているエンティティ (人物など) を取得します。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|リッチコメントの内容 (コメントに含まれるメンションなど) を取得します。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[解析](/javascript/api/excel/excel.comment#resolved)|コメントスレッドの状態を取得または設定します。 値 "true" は、コメントスレッドが解決された状態であることを意味します。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[コメントについて](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|コメントで言及されているエンティティの電子メールアドレスを取得または設定します。|
||[id](/javascript/api/excel/excel.commentmention#id)|エンティティの id を取得または設定します。 これは、の`CommentRichContent.richContent`id 情報に合わせて配置されます。|
||[name](/javascript/api/excel/excel.commentmention#name)|コメントで言及されているエンティティの名前を取得または設定します。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|コメントに記載されているエンティティ (人物など) を取得します。|
||[解析](/javascript/api/excel/excel.commentreply#resolved)|コメントの返信状態を取得または設定します。 値 "true" は、コメント応答が解決された状態であることを意味します。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|リッチコメントの内容 (コメントに含まれるメンションなど) を取得します。 この文字列は、エンドユーザーに表示されることを意図したものではありません。 アドインでは、リッチコメントコンテンツを解析するためにのみ使用する必要があります。|
||[updateMentions (contentWithMentions ション: CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|特別に書式設定された文字列とメンションの一覧を使用して、コメントの内容を更新します。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|コメント内で記述されているすべてのエンティティ (人物など) を含む配列。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|スピルするセルのアンカー セルを含む範囲オブジェクトを取得します。 読み取り専用です。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 複数のセルを含む範囲に適用される場合は失敗します。 読み取り専用です。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|アンカー セルで呼び出されたとき、スピル範囲を含む範囲オブジェクトを取得します。 読み取り専用です。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|すべてのセルにスピル ボーダーがあるかどうかを表します。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|すべてのセルが配列数式として保存されるかどうかを表します。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (金額: 数値)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|範囲の書式のインデントを調整します。 インデントの値の範囲は 0 ~ 250 です。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの id を表します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルが含まれるワークシートの ID を表します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|現在のブックを閉じます。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|現在のブックを保存します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|特定のワークシートで、1つまたは複数の行の非表示の状態が変更されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を表します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されたワークシートの ID を表します。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|特定のワークシートで変更されたエリアを表す範囲のアドレスを取得します。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|イベントがトリガーされた方法を表す変更の種類を取得します。 詳細は「`Excel.RowHiddenChangeType`」をご覧ください。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|データが変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
