---
title: Word JavaScript プレビュー API
description: 今後の Word JavaScript API の詳細。
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
---

# <a name="word-javascript-preview-apis"></a>Word JavaScript プレビュー API

新しい Word JavaScript API は、最初に "プレビュー" で導入され、後で十分なテストが行われるとユーザーフィードバックが取得された後、特定の番号付き要件セットの一部になります。

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Word JavaScript API の一覧を示します。ただし、この API は、現在のバージョン[でのみWord on the web](#web-only-api-list)。 すべての Word JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧を表示するには、 [すべての Word JavaScript API を参照してください](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|コンテンツ コントロール内のデータが変更された場合に発生します。|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|コンテンツ コントロールが削除された場合に発生します。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|コンテンツ コントロール内の選択が変更された場合に発生します。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|イベントを発生させたオブジェクト。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|イベントの種類。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|カスタム XML パーツを削除します。|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|xpath で識別される要素から、指定された名前の属性を削除します。|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|xpath で識別される要素を削除します。|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|カスタム XML パーツの完全な XML コンテンツを取得します。|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|カスタム XML パーツの ID を取得します。|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|指定された名前と値を持つ属性を、xpath で識別される要素に挿入します。|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|xpath で識別される親要素の下に、指定された XML を子位置インデックスに挿入します。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|カスタム XML パーツの名前空間 URI を取得します。|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|カスタム XML パーツの XML コンテンツを照会します。|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|カスタム XML パーツの完全な XML コンテンツを設定します。|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|xpath で識別される要素の指定された名前を持つ属性の値を更新します。|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|xpath で識別される要素の XML を更新します。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|新しいカスタム XML パーツをドキュメントに追加します。|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|ドキュメント内のカスタム XML パーツを取得します。|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|ブックマークの範囲を取得します。|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|コンテンツ コントロールが追加された場合に発生します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|文書全体の範囲で指定された検索オプションを使用して検索を実行します。|
||[settings](/javascript/api/word/word.document#word-word-document-settings-member)|ドキュメント内のアドインの設定を取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|ドキュメント内のカスタム XML パーツを取得します。|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|ブックマークの範囲を取得します。|
||[settings](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|ドキュメント内のアドインの設定を取得します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|インライン イメージの形式を取得します。|
|[リスト](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|リスト内の指定されたレベルの箇条書き、数字、または図のフォントを取得します。|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|リスト内の指定されたレベルの図の base64 エンコードされた文字列表現を取得します。|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|箇条書き、番号、または図のフォントを、リスト内の指定されたレベルでリセットします。|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|リスト内の指定されたレベルで図を設定します。|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|範囲内または範囲に重なるすべてのブックマークの名前を取得します。|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|範囲にブックマークを挿入します。|
|[設定](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|設定を削除します。|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|設定のキーを取得します。|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|設定の値を取得または設定します。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|新しい設定を作成するか、既存の設定を設定します。|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|このアドインのすべての設定を削除します。|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|設定の数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|最初のセルと最後のセルで結合されたセルを結合します。|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|セルを指定した数の行と列に分割します。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|行にコンテンツ コントロールを挿入します。|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|行を 1 つのセルに結合します。|

## <a name="web-only-api-list"></a>Web 専用 API リスト

次の表に、現在プレビュー中の Word JavaScript API の一覧を、Word on the web。 すべての Word JavaScript API (プレビュー API と以前にリリースされた API を含む) の完全な一覧を表示するには、 [すべての Word JavaScript API を参照してください](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|本文の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|本文の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|本文に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|本文の種類を取得します。|
|[コメント](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|コメント作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|コメント作成者の名前を取得します。|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|コメントのコンテンツをプレーン テキストとして取得または設定します。|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|コメント スレッドの状態を取得または設定します。|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|コメントの作成日を取得します。|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|コメントとその返信を削除します。|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|コメントがオンのメイン ドキュメント内の範囲を取得します。|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|コメントに関連付けられた返信オブジェクトのコレクションを取得します。|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|コメント スレッドの末尾に新しい返信を追加します。|
||[解決済み](/javascript/api/word/word.comment#word-word-comment-resolved-member)|コメント スレッドの状態を取得または設定します。|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|コレクション内の最初のコメントを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|コレクション内の最初のコメントを取得します。|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|コレクション内のインデックスによってコメント オブジェクトを取得します。|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|コメント テキストが太字かどうかを示す値を取得または設定します。|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|指定した場所にテキストを挿入します。|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|範囲の長さが 0 であるかどうかを確認します。|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|コメント テキストがイタル化されているかどうかを示す値を取得または設定します。|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|コメント テキストに取り消し線が設定されているかどうかを示す値を取得または設定します。|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|コメント範囲のテキストを取得します。|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|コメント テキストの下線の種類を示す値を取得または設定します。|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|コメント返信作成者のメール アドレスを取得します。|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|コメント返信作成者の名前を取得します。|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|コメント返信の内容を取得または設定します。|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|commentReply のコンテンツ範囲を取得または設定します。|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|コメント返信の作成日を取得します。|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|コメント返信を削除します。|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|この返信の親コメントを取得します。|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|コレクション内の最初のコメント返信を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|コレクション内の最初のコメント返信を取得します。|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|コレクション内のインデックスによってコメント返信オブジェクトを取得します。|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|コンテンツ コントロール内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|コンテンツ コントロール内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|本文に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|ChangeTracking モードを取得または設定します。|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|1 つの本文でドキュメントの文末脚注を取得します。|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|1 つの本文でドキュメントの脚注を取得します。|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|メモ アイテムの body オブジェクトを表します。|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|メモ アイテムを削除します。|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|同じ種類の次のノート アイテムを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|同じ種類の次のノート アイテムを取得します。|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|メイン ドキュメントの脚注または文末脚注参照を表します。|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|メモ アイテムの種類である脚注または文末脚注を表します。|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|このコレクションの最初のノート アイテムを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|このコレクションの最初のノート アイテムを取得します。|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|段落内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|段落内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|段落に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|範囲内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|範囲内の脚注のコレクションを取得します。|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|範囲に関連付けられたコメントを取得します。|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|ChangeTrackingVersion の選択に基づいて確認されたテキストを取得します。|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|範囲にコメントを挿入します。|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|文末脚注を挿入します。|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|脚注を挿入します。|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|テーブル内の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|テーブル内の脚注のコレクションを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|テーブル行の文末脚注のコレクションを取得します。|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|テーブル行の脚注のコレクションを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
