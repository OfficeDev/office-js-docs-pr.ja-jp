---
title: Word JavaScript API 要件セット 1.1
description: WordApi 1.1 要件セットの詳細
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: bb2ad35e3dfe690437a6081dc5790dc5c36ec84c
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681703"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 の新機能

WordApi 1.1 は、Word JavaScript API の最初の要件セットです。 これは、ユーザーがサポートする唯一の Word API 要件セットWord 2016。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.1 の API を示します。 Word JavaScript API 要件セット 1.1 でサポートされるすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.1 の Word API」を参照してください](/javascript/api/word?view=word-js-1.1&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|本文オブジェクトの内容を消去します。|
||[contentControls](/javascript/api/word/word.body#contentControls)|本文内のリッチ テキスト コンテンツ コントロール オブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.body#font)|本文のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|body オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|本文内の InlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|文書を本文の指定された位置に挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|指定した位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|指定した位置に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertText_text__insertLocation_)|テキストを本文の指定された位置に挿入します。|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[paragraphs](/javascript/api/word/word.body#paragraphs)|本文内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|本文を含むコンテンツ コントロールを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|body オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select_selectionMode_)|本文を選択し、その本文に Word の UI を移動します。|
||[style](/javascript/api/word/word.body#style)|本文のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.body#text)|本文のテキストを取得します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[外観](/javascript/api/word/word.contentcontrol#appearance)|コンテンツ コントロールの外観を取得または設定します。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|ユーザーがコンテンツ コントロールを削除できるかどうかを示す値を取得または設定します。|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|ユーザーがコンテンツ コントロールのコンテンツを編集できるかどうかを示す値を取得または設定します。|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|コンテンツ コントロールの内容をクリアします。|
||[color](/javascript/api/word/word.contentcontrol#color)|コンテンツ コントロールの色を取得または設定します。|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|コンテンツ コントロールのコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete_keepContent_)|コンテンツ コントロールとそのコンテンツを削除します。|
||[font](/javascript/api/word/word.contentcontrol#font)|コンテンツ コントロールのテキストの書式設定を取得します。|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|コンテンツ コントロール オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[id](/javascript/api/word/word.contentcontrol#id)|コンテンツ コントロールの識別子を表す整数値を取得します。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|コンテンツ コントロールに含まれる inlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|指定した場所にあるコンテンツ コントロールにドキュメントを挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|コンテンツ コントロール内の指定された位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|指定した場所にあるコンテンツ コントロールに OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|コンテンツ コントロール内の指定された位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|コンテンツ コントロール内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|コンテンツ コントロールのプレースホルダー テキストを取得または設定します。|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|コンテンツ コントロールを編集後に削除できるかどうかを示す値を取得または設定します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|コンテンツ コントロール オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select_selectionMode_)|コンテンツ コントロールを選択します。|
||[style](/javascript/api/word/word.contentcontrol#style)|コンテンツ コントロールのスタイル名を取得または設定します。|
||[タグ](/javascript/api/word/word.contentcontrol#tag)|コンテンツ コントロールを識別するタグを取得または設定します。|
||[text](/javascript/api/word/word.contentcontrol#text)|コンテンツ コントロールのテキストを取得します。|
||[title](/javascript/api/word/word.contentcontrol#title)|コンテンツ コントロールのタイトルを取得または設定します。|
||[type](/javascript/api/word/word.contentcontrol#type)|コンテンツ コントロールの種類を取得します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|指定されたタグを含むコンテンツ コントロールを取得します。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|指定されたタイトルを含むコンテンツ コントロールを取得します。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|コレクション内のインデックスによってコンテンツ コントロールを取得します。|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#body)|メイン ドキュメントの body オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.document#contentControls)|ドキュメント内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[getSelection()](/javascript/api/word/word.document#getSelection__)|ドキュメントの現在の選択範囲を取得します。|
||[save()](/javascript/api/word/word.document#save__)|ドキュメントを保存します。|
||[保存済み](/javascript/api/word/word.document#saved)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.document#sections)|ドキュメント内のセクション オブジェクトのコレクションを取得します。|
|[フォント](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|フォントが太字かどうかを示す値を取得または設定します。|
||[color](/javascript/api/word/word.font#color)|指定されたフォントの色を取得または設定します。|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|フォントに二重取り消し線が設定されているかどうかを示す値を取得または設定します。|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|強調表示の色を取得または設定します。|
||[italic](/javascript/api/word/word.font#italic)|フォントが斜体かどうかを示す値を取得または設定します。|
||[name](/javascript/api/word/word.font#name)|フォント名を表す値を取得または設定します。|
||[size](/javascript/api/word/word.font#size)|フォント サイズをポイント単位で表す値を取得または設定します。|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|フォントに取り消し線が設定されているかどうかを示す値を取得または設定します。|
||[subscript](/javascript/api/word/word.font#subscript)|フォントが下付き文字かどうかを示す値を取得または設定します。|
||[superscript](/javascript/api/word/word.font#superscript)|フォントが上付き文字かどうかを示す値を取得または設定します。|
||[underline](/javascript/api/word/word.font#underline)|フォントの下線の種類を示す値を取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|インライン イメージに関連付けられた代替テキストを表す文字列を取得または設定します。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|インライン画像のタイトルを含む文字列を取得または設定します。|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|インライン画像の base64 エンコード文字列形式を取得します。|
||[height](/javascript/api/word/word.inlinepicture#height)|インライン画像の高さを表す数値を取得するか設定します。|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|イメージ上のハイパーリンクを取得または設定します。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|リッチ テキストのコンテンツ コントロールでインライン画像をラップします。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|インライン画像を含むコンテンツ コントロールを取得します。|
||[width](/javascript/api/word/word.inlinepicture#width)|インライン画像の幅を表す数値を取得するか設定します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[配置](/javascript/api/word/word.paragraph#alignment)|段落の配置を取得または設定します。|
||[clear()](/javascript/api/word/word.paragraph#clear__)|段落オブジェクトの内容をクリアします。|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|段落内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete()](/javascript/api/word/word.paragraph#delete__)|文書から段落と、その段落の内容を削除します。|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|最初の行またはぶら下げインデントの値をポイントで取得または設定します。|
||[font](/javascript/api/word/word.paragraph#font)|段落のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|段落オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|段落内の InlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|指定した場所の段落にドキュメントを挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|段落の指定した位置に、HTML を挿入します。|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|段落の指定した位置に、図を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|指定した場所の段落に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|段落の指定した位置に、テキストを挿入します。|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|段落の左インデントの値をポイント数単位で取得または設定します。|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|段落の行間をポイント数単位で取得または設定します。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|段落の後のグリッド線の間隔の量を取得または設定します。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|段落前の間隔の幅をグリッド線数単位で取得または設定します。|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|段落のアウトライン レベルを取得または設定します。|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|段落を格納しているコンテンツ コントロールを取得します。|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|段落の右インデントの値をポイント数単位で取得または設定します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Paragraph オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select_selectionMode_)|段落を選択して、その段落に Word の UI を移動します。|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|段落後の間隔をポイント数単位で取得または設定します。|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|段落前の間隔をポイント数単位で取得または設定します。|
||[style](/javascript/api/word/word.paragraph#style)|段落のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.paragraph#text)|段落のテキストを取得します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|範囲オブジェクトの内容をクリアします。|
||[contentControls](/javascript/api/word/word.range#contentControls)|範囲内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete()](/javascript/api/word/word.range#delete__)|文書から範囲と、その範囲の内容を削除します。|
||[font](/javascript/api/word/word.range#font)|範囲のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|範囲オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|Range オブジェクトの OOXML 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|指定した位置に文書を挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|指定した位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|指定した位置に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertText_text__insertLocation_)|指定した位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[paragraphs](/javascript/api/word/word.range#paragraphs)|範囲内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|範囲を格納するコンテンツ コントロールを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|range オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select_selectionMode_)|範囲を選択して、その範囲に Word の UI を移動します。|
||[style](/javascript/api/word/word.range#style)|範囲のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.range#text)|範囲のテキストを取得します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|単語間のすべての区切り記号を無視するかどうかを示す値を取得または設定します。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|検索文字列で終わる語句と一致するかどうかを示す値を取得または設定します。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|長い単語の一部ではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#body)|セクションの body オブジェクトを取得します。|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getFooter_type_)|セクションのフッターの 1 つを取得します。|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getHeader_type_)|セクションのヘッダーの 1 つを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
