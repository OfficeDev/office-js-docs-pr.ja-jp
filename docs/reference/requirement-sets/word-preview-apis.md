---
title: Word JavaScript プレビュー Api
description: 今後の Word JavaScript Api の詳細
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 6a3b67e65c4ced3f1b89d98afe45d5d6c33f63b6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996404"
---
# <a name="word-javascript-preview-apis"></a>Word JavaScript プレビュー Api

新しい Word JavaScript Api は最初に "プレビュー" で導入され、後でテストが行われ、ユーザーのフィードバックが取得された後に、特定の番号付き要件の一部になります。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Word JavaScript Api を示します。 すべての Word JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストを表示するには、「 [すべての Word Javascript api](/javascript/api/word?view=word-js-preview&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|コンテンツコントロール内のデータが変更されるときに発生します。|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|コンテンツコントロールが削除されるときに発生します。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|コンテンツコントロール内の選択範囲が変更されたときに発生します。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|イベントを発生させたオブジェクト。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|イベントの種類。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|カスタム XML パーツを削除します。|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Xpath で識別される要素から、指定した名前の属性を削除します。|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Xpath で識別される要素を削除します。|
||[getXml ()](/javascript/api/word/word.customxmlpart#getxml--)|カスタム XML パーツの完全な XML コンテンツを取得します。|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|指定した名前と値を持つ属性を、xpath で識別される要素に挿入します。|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Xpath で識別される親要素の下に、子の位置インデックスで指定した XML を挿入します。|
||[query (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|カスタム XML パーツの XML コンテンツを照会します。|
||[id](/javascript/api/word/word.customxmlpart#id)|カスタム XML パーツの ID を取得します。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|カスタム XML パーツの名前空間 URI を取得します。|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|カスタム XML パーツの完全な XML コンテンツを設定します。|
||[updateAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Xpath で識別される要素の指定した名前で属性の値を更新します。|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Xpath で識別される要素の XML を更新します。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add (xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|文書に新しいカスタム XML 部分を追加します。|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem ()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (name: string)](/javascript/api/word/word.document#deletebookmark-name-)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange (name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject (name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|ブックマークの範囲を取得します。|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|ドキュメント内のカスタム XML パーツを取得します。|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|コンテンツコントロールが追加されると発生します。|
||[settings](/javascript/api/word/word.document#settings)|文書内のアドインの設定を取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange (name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject (name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|ブックマークの範囲を取得します。|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|ドキュメント内のカスタム XML パーツを取得します。|
||[settings](/javascript/api/word/word.documentcreated#settings)|文書内のアドインの設定を取得します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|インライン画像の形式を取得します。|
|[List](/javascript/api/word/word.list)|[getLevelFont (level: number)](/javascript/api/word/word.list#getlevelfont-level-)|リスト内の指定されたレベルの行頭文字、番号、または図のフォントを取得します。|
||[getLevelPicture (level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|リスト内の指定されたレベルにある画像の base64 エンコード文字列表現を取得します。|
||[resetLevelFont (level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|リスト内の指定されたレベルの行頭文字、番号、または図のフォントをリセットします。|
||[setLevelPicture (level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|リスト内の指定されたレベルで画像を設定します。|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: boolean, Includehidden?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|指定した範囲に含まれるすべてのブックマークの名前を取得します。|
||[insertBookmark (name: string)](/javascript/api/word/word.range#insertbookmark-name-)|範囲にブックマークを挿入します。|
|[設定](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|設定を削除します。|
||[key](/javascript/api/word/word.setting#key)|設定のキーを取得します。|
||[value](/javascript/api/word/word.setting#value)|設定の値を取得または設定します。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add (key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|新しい設定を作成するか、既存の設定を設定します。|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|このアドインのすべての設定を削除します。|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|設定の数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[items](/javascript/api/word/word.settingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow: number, firstCell: number, 下端行: 数値, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|最初と最後のセルによって制限されたセルを結合します。|
|[TableCell](/javascript/api/word/word.tablecell)|[split (rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|指定された行数と列数にセルを分割します。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|行にコンテンツコントロールを挿入します。|
||[merge ()](/javascript/api/word/word.tablerow#merge--)|1つのセルに行を結合します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
