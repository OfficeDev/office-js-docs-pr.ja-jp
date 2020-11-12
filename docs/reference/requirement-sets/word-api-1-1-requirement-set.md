---
title: Word JavaScript API 要件セット1.1
description: WordApi 1.1 要件セットの詳細
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 371638c18cff882f2b3907f1adedb6748761cc0c
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996439"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 の新機能

WordApi 1.1 は、Word JavaScript API の最初の要件セットです。 Word 2016 でサポートされている唯一の Word API 要件セットです。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット1.1 の Api を示します。 Word JavaScript API 要件セット1.1 でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.1 の Word api](/javascript/api/word?view=word-js-1.1&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|本文オブジェクトの内容を消去します。|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Body オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|文書を本文の指定された位置に挿入します。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|指定した位置に HTML を挿入します。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|指定した位置に OOXML を挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|テキストを本文の指定された位置に挿入します。|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|本文に含まれるリッチテキストコンテンツコントロールオブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.body#font)|本文のテキスト形式を取得します。|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|本文にある InlinePicture オブジェクトのコレクションを取得します。|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|本文に含まれる paragraph オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|本文を含むコンテンツ コントロールを取得します。|
||[text](/javascript/api/word/word.body#text)|本文のテキストを取得します。|
||[search (searchText: string, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: boolean Matchcase?: Boolean matchWholeWord?: ブール型の一致ワイルドカード?: boolean})](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Body オブジェクトのスコープで、指定された SearchOptions を使用して検索を実行します。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|本文を選択し、その本文に Word の UI を移動します。|
||[style](/javascript/api/word/word.body#style)|本文のスタイル名を取得または設定します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[外観](/javascript/api/word/word.contentcontrol#appearance)|コンテンツ コントロールの外観を取得または設定します。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|ユーザーがコンテンツ コントロールを削除できるかどうかを示す値を取得または設定します。|
||[Canメモ Dit](/javascript/api/word/word.contentcontrol#cannotedit)|ユーザーがコンテンツ コントロールのコンテンツを編集できるかどうかを示す値を取得または設定します。|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|コンテンツ コントロールの内容をクリアします。|
||[color](/javascript/api/word/word.contentcontrol#color)|コンテンツ コントロールの色を取得または設定します。|
||[削除 (keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|コンテンツ コントロールとそのコンテンツを削除します。|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|コンテンツコントロールオブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|指定した位置にコンテンツコントロールにドキュメントを挿入します。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|コンテンツ コントロール内の指定された位置に HTML を挿入します。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|指定した位置に、コンテンツコントロールに OOXML を挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|コンテンツ コントロール内の指定された位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|コンテンツ コントロールのプレースホルダー テキストを取得または設定します。|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|コンテンツ コントロールのコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.contentcontrol#font)|コンテンツ コントロールのテキストの書式設定を取得します。|
||[id](/javascript/api/word/word.contentcontrol#id)|コンテンツ コントロールの識別子を表す整数値を取得します。|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|コンテンツ コントロールに含まれる inlinePicture オブジェクトのコレクションを取得します。|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|コンテンツ コントロールにある Paragraph オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[text](/javascript/api/word/word.contentcontrol#text)|コンテンツ コントロールのテキストを取得します。|
||[type](/javascript/api/word/word.contentcontrol#type)|コンテンツ コントロールの種類を取得します。|
||[removeWhenEdited 済み](/javascript/api/word/word.contentcontrol#removewhenedited)|コンテンツ コントロールを編集後に削除できるかどうかを示す値を取得または設定します。|
||[search (searchText: string, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: boolean Matchcase?: Boolean matchWholeWord?: ブール型の一致ワイルドカード?: boolean})](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|コンテンツコントロールオブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|コンテンツ コントロールを選択します。|
||[style](/javascript/api/word/word.contentcontrol#style)|コンテンツコントロールのスタイル名を取得または設定します。|
||[マーク](/javascript/api/word/word.contentcontrol#tag)|コンテンツコントロールを識別するタグを取得または設定します。|
||[title](/javascript/api/word/word.contentcontrol#title)|コンテンツ コントロールのタイトルを取得または設定します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|指定されたタグを含むコンテンツ コントロールを取得します。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|指定されたタイトルを含むコンテンツ コントロールを取得します。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|コレクション内のインデックスによってコンテンツコントロールを取得します。|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[getSelection ()](/javascript/api/word/word.document#getselection--)|ドキュメントの現在の選択範囲を取得します。|
||[body](/javascript/api/word/word.document#body)|文書の本文オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|文書内のコンテンツコントロールオブジェクトのコレクションを取得します。|
||[更新](/javascript/api/word/word.document#saved)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.document#sections)|ドキュメント内の section オブジェクトのコレクションを取得します。|
||[save()](/javascript/api/word/word.document#save--)|ドキュメントを保存します。|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|フォントが太字かどうかを示す値を取得または設定します。|
||[color](/javascript/api/word/word.font#color)|指定されたフォントの色を取得または設定します。|
||[[Doublestrikethrough]](/javascript/api/word/word.font#doublestrikethrough)|フォントに二重取り消し線があるかどうかを示す値を取得または設定します。|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|強調表示の色を取得または設定します。|
||[italic](/javascript/api/word/word.font#italic)|フォントが斜体かどうかを示す値を取得または設定します。|
||[name](/javascript/api/word/word.font#name)|フォント名を表す値を取得または設定します。|
||[size](/javascript/api/word/word.font#size)|フォント サイズをポイント単位で表す値を取得または設定します。|
||[打ち消し](/javascript/api/word/word.font#strikethrough)|フォントに取り消し線を表示するかどうかを示す値を取得または設定します。|
||[subscript](/javascript/api/word/word.font#subscript)|フォントが下付き文字かどうかを示す値を取得または設定します。|
||[superscript](/javascript/api/word/word.font#superscript)|フォントが上付き文字かどうかを示す値を取得または設定します。|
||[underline](/javascript/api/word/word.font#underline)|フォントの下線の種類を示す値を取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|インライン画像に関連付けられている代替テキストを表す文字列を取得または設定します。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|インライン画像のタイトルを含む文字列を取得または設定します。|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|インライン画像の base64 エンコード文字列形式を取得します。|
||[height](/javascript/api/word/word.inlinepicture#height)|インライン画像の高さを表す数値を取得するか設定します。|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|画像のハイパーリンクを取得または設定します。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|リッチ テキストのコンテンツ コントロールでインライン画像をラップします。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|インライン画像を含むコンテンツ コントロールを取得します。|
||[width](/javascript/api/word/word.inlinepicture#width)|インライン画像の幅を表す数値を取得するか設定します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[策定](/javascript/api/word/word.paragraph#alignment)|段落の配置を取得または設定します。|
||[clear()](/javascript/api/word/word.paragraph#clear--)|段落オブジェクトの内容をクリアします。|
||[delete()](/javascript/api/word/word.paragraph#delete--)|文書から段落と、その段落の内容を削除します。|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|最初の行またはぶら下げインデントの値をポイント単位で取得または設定します。|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Paragraph オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|指定した位置に段落に文書を挿入します。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|段落の指定した位置に、HTML を挿入します。|
||[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|段落の指定した位置に、図を挿入します。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|指定した位置の段落に OOXML を挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|段落の指定した位置に、テキストを挿入します。|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|段落の左インデントの値をポイント数単位で取得または設定します。|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|段落の行間をポイント数単位で取得または設定します。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|段落後の間隔の量 (グリッド線単位) を取得または設定します。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|段落前の間隔の幅をグリッド線数単位で取得または設定します。|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|段落のアウトライン レベルを取得または設定します。|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|段落内のコンテンツコントロールオブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.paragraph#font)|段落のテキスト形式を取得します。|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|段落内の InlinePicture オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|段落を格納しているコンテンツ コントロールを取得します。|
||[text](/javascript/api/word/word.paragraph#text)|段落のテキストを取得します。|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|段落の右インデントの値をポイント数単位で取得または設定します。|
||[search (searchText: string, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: boolean Matchcase?: Boolean matchWholeWord?: ブール型の一致ワイルドカード?: boolean})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Paragraph オブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|段落を選択して、その段落に Word の UI を移動します。|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|段落後の間隔をポイント数単位で取得または設定します。|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|段落前の間隔をポイント数単位で取得または設定します。|
||[style](/javascript/api/word/word.paragraph#style)|段落のスタイル名を取得または設定します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|範囲オブジェクトの内容をクリアします。|
||[delete()](/javascript/api/word/word.range#delete--)|文書から範囲と、その範囲の内容を削除します。|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Range オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Range オブジェクトの OOXML 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|指定した位置に文書を挿入します。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|指定した位置に HTML を挿入します。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|指定した位置に OOXML を挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|指定した位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|範囲内のコンテンツコントロールオブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.range#font)|範囲のテキスト形式を取得します。|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|範囲内の paragraph オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|範囲を格納するコンテンツ コントロールを取得します。|
||[text](/javascript/api/word/word.range#text)|範囲のテキストを取得します。|
||[search (searchText: string, searchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignoreSpace?: Boolean matchCase?: Boolean matchPrefix?: boolean Matchcase?: Boolean matchWholeWord?: ブール型の一致ワイルドカード?: boolean})](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Range オブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|範囲を選択して、その範囲に Word の UI を移動します。|
||[style](/javascript/api/word/word.range#style)|範囲のスタイル名を取得または設定します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|単語間のすべての区切り記号を無視するかどうかを示す値を取得または設定します。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|単語間のすべての空白文字を無視するかどうかを示す値を取得または設定します。|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|検索文字列で終わる語句と一致するかどうかを示す値を取得または設定します。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|長い単語の一部ではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。|
|[Section](/javascript/api/word/word.section)|[getFooter (type: Word Headerfooter Type)](/javascript/api/word/word.section#getfooter-type-)|セクションのフッターの 1 つを取得します。|
||[getHeader (type: Word Headerフッターの種類)](/javascript/api/word/word.section#getheader-type-)|セクションのヘッダーの 1 つを取得します。|
||[body](/javascript/api/word/word.section#body)|セクションの本文オブジェクトを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
