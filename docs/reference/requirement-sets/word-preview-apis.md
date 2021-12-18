---
title: Word JavaScript プレビュー API
description: 今後の Word JavaScript API の詳細。
ms.date: 12/14/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: c68a63dc57fbcaa8282343c3f3271778c43bc28d
ms.sourcegitcommit: 9b6556563451f9907cb5da50cba757eb9960aa39
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/17/2021
ms.locfileid: "61565365"
---
# <a name="word-javascript-preview-apis"></a>Word JavaScript プレビュー API

新しい Word JavaScript API は、最初に "プレビュー" で導入され、後で十分なテストが行われるとユーザーフィードバックが取得された後、特定の番号付き要件セットの一部になります。

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Word JavaScript API の一覧を示します。ただし、現在プレビュー中の Api は、現在のバージョン[でのみ使用Word on the web。](#web-only-api-list) すべての Word JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧を表示するには、 [すべての Word JavaScript API を参照してください](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|コンテンツ コントロール内のデータが変更された場合に発生します。|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|コンテンツ コントロールが削除された場合に発生します。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|コンテンツ コントロール内の選択が変更された場合に発生します。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|イベントを発生させたオブジェクト。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|イベントの種類。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|カスタム XML パーツを削除します。|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|xpath で識別される要素から、指定された名前の属性を削除します。|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|xpath で識別される要素を削除します。|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|カスタム XML パーツの完全な XML コンテンツを取得します。|
||[id](/javascript/api/word/word.customxmlpart#id)|カスタム XML パーツの ID を取得します。|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|指定された名前と値を持つ属性を、xpath で識別される要素に挿入します。|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|xpath で識別される親要素の下に、指定された XML を子位置インデックスに挿入します。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|カスタム XML パーツの名前空間 URI を取得します。|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|カスタム XML パーツの XML コンテンツを照会します。|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|カスタム XML パーツの完全な XML コンテンツを設定します。|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|xpath で識別される要素の指定された名前を持つ属性の値を更新します。|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|xpath で識別される要素の XML を更新します。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|新しいカスタム XML パーツをドキュメントに追加します。|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#customXmlParts)|ドキュメント内のカスタム XML パーツを取得します。|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|ブックマークの範囲を取得します。|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|コンテンツ コントロールが追加された場合に発生します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#search_searchText__searchOptions_)|文書全体の範囲で指定された検索オプションを使用して検索を実行します。|
||[settings](/javascript/api/word/word.document#settings)|ドキュメント内のアドインの設定を取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|ドキュメント内のカスタム XML パーツを取得します。|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|ブックマークの範囲を取得します。|
||[settings](/javascript/api/word/word.documentcreated#settings)|ドキュメント内のアドインの設定を取得します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|インライン イメージの形式を取得します。|
|[リスト](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|リスト内の指定されたレベルの箇条書き、数字、または図のフォントを取得します。|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|リスト内の指定されたレベルの図の base64 エンコードされた文字列表現を取得します。|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|箇条書き、番号、または図のフォントを、リスト内の指定されたレベルでリセットします。|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|リスト内の指定されたレベルで図を設定します。|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|範囲内または範囲に重なるすべてのブックマークの名前を取得します。|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|範囲にブックマークを挿入します。|
|[設定](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|設定を削除します。|
||[key](/javascript/api/word/word.setting#key)|設定のキーを取得します。|
||[value](/javascript/api/word/word.setting#value)|設定の値を取得または設定します。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|新しい設定を作成するか、既存の設定を設定します。|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|このアドインのすべての設定を削除します。|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|設定の数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[items](/javascript/api/word/word.settingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|最初のセルと最後のセルで結合されたセルを結合します。|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|セルを指定した数の行と列に分割します。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|行にコンテンツ コントロールを挿入します。|
||[merge()](/javascript/api/word/word.tablerow#merge__)|行を 1 つのセルに結合します。|

## <a name="web-only-api-list"></a>Web 専用 API リスト

次の表に、現在プレビュー中の Word JavaScript API の一覧を、Word on the web。 すべての Word JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧を表示するには、 [すべての Word JavaScript API を参照してください](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#endnotes)|本文の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.body#footnotes)|本文の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.body#getComments__)|本文に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#getReviewedText_changeTrackingVersion_)|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
||[type](/javascript/api/word/word.body#type)|本文の種類を取得します。|
|[コメント](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#authorEmail)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/word/word.comment#authorName)|コメント作成者の名前を取得します。|
||[content](/javascript/api/word/word.comment#content)|コメントのコンテンツをプレーン テキストとして取得または設定します。|
||[creationDate](/javascript/api/word/word.comment#creationDate)|コメントの作成日を取得します。|
||[delete()](/javascript/api/word/word.comment#delete__)|コメントとその返信を削除します。|
||[getRange()](/javascript/api/word/word.comment#getRange__)|コメントがオンのメイン ドキュメント内の範囲を取得します。|
||[id](/javascript/api/word/word.comment#id)|ID|
||[replies](/javascript/api/word/word.comment#replies)|コメントに関連付けられた返信オブジェクトのコレクションを取得します。|
||[reply(replyText: string)](/javascript/api/word/word.comment#reply_replyText_)|コメント スレッドの末尾に新しい返信を追加します。|
||[解決済み](/javascript/api/word/word.comment#resolved)|コメント スレッドの状態を取得または設定します。|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|コレクション内の最初のコメントを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|コレクション内の最初のコメントまたは null オブジェクトを取得します。|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#getItem_index_)|コレクション内のインデックスによってコメント オブジェクトを取得します。|
||[items](/javascript/api/word/word.commentcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/word/word.commentreply#authorName)|コメント返信作成者の名前を取得します。|
||[content](/javascript/api/word/word.commentreply#content)|コメント返信の内容を取得または設定します。|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|コメント返信の作成日を取得します。|
||[delete()](/javascript/api/word/word.commentreply#delete__)|コメント返信を削除します。|
||[id](/javascript/api/word/word.commentreply#id)|ID|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|この返信の親コメントを取得します。|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|コレクション内の最初のコメント返信を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|コレクション内の最初のコメント返信または null オブジェクトを取得します。|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#getItem_index_)|コレクション内のインデックスによってコメント返信オブジェクトを取得します。|
||[items](/javascript/api/word/word.commentreplycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#endnotes)|コンテンツ コントロール内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.contentcontrol#footnotes)|コンテンツ コントロール内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.contentcontrol#getComments__)|本文に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#getReviewedText_changeTrackingVersion_)|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#changeTrackingMode)|ChangeTracking モードを取得または設定します。|
||[getEndnoteBody()](/javascript/api/word/word.document#getEndnoteBody__)|1 つの本文でドキュメントの文末脚注を取得します。|
||[getFootnoteBody()](/javascript/api/word/word.document#getFootnoteBody__)|1 つの本文でドキュメントの脚注を取得します。|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#body)|メモ アイテムの body オブジェクトを表します。|
||[delete()](/javascript/api/word/word.noteitem#delete__)|メモ アイテムを削除します。|
||[getNext()](/javascript/api/word/word.noteitem#getNext__)|同じ種類の次のノート アイテムを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#getNextOrNullObject__)|同じ種類の次のノート アイテムを取得します。|
||[reference](/javascript/api/word/word.noteitem#reference)|メイン ドキュメントの脚注または文末脚注参照を表します。|
||[type](/javascript/api/word/word.noteitem#type)|メモ アイテムの種類である脚注または文末脚注を表します。|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|このコレクションの最初のノート アイテムを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|このコレクションの最初のノート アイテムを取得します。|
||[items](/javascript/api/word/word.noteitemcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#endnotes)|段落内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.paragraph#footnotes)|段落内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.paragraph#getComments__)|段落に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#getReviewedText_changeTrackingVersion_)|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#endnotes)|範囲内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.range#footnotes)|範囲内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.range#getComments__)|範囲に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#getReviewedText_changeTrackingVersion_)|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
||[insertComment(commentText: string)](/javascript/api/word/word.range#insertComment_commentText_)|範囲にコメントを挿入します。|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#insertEndnote_insertText_)|文末脚注を挿入します。|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#insertFootnote_insertText_)|脚注を挿入します。|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#endnotes)|テーブル内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.table#footnotes)|テーブル内の脚注のコレクションを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#endnotes)|テーブル行の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.tablerow#footnotes)|テーブル行の脚注のコレクションを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
