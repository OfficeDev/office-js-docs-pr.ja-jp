---
title: Word JavaScript API 要件セット 1.1
description: WordApi 1.1 要件セットの詳細。
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: dfcb1954cd9522de6165130cc115fddbb5f3ec45
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744213"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 の新機能

WordApi 1.1 は、Word JavaScript API の最初の要件セットです。 これは、ユーザーがサポートする唯一の Word API 要件セットWord 2016。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.1 の API を示します。 Word JavaScript API 要件セット 1.1 でサポートされるすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.1 の Word API」を参照してください](/javascript/api/word?view=word-js-1.1&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|本文オブジェクトの内容を消去します。|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|本文内のリッチ テキスト コンテンツ コントロール オブジェクトのコレクションを取得します。|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|本文のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|body オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|本文内の InlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|文書を本文の指定された位置に挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|指定した位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|指定した位置に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|テキストを本文の指定された位置に挿入します。|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|本文内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|本文を含むコンテンツ コントロールを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.body#word-word-body-search-member(1))|body オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|本文を選択し、その本文に Word の UI を移動します。|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|本文のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|本文のテキストを取得します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[外観](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|コンテンツ コントロールの外観を取得または設定します。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|ユーザーがコンテンツ コントロールを削除できるかどうかを示す値を取得または設定します。|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|ユーザーがコンテンツ コントロールのコンテンツを編集できるかどうかを示す値を取得または設定します。|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|コンテンツ コントロールの内容をクリアします。|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|コンテンツ コントロールの色を取得または設定します。|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|コンテンツ コントロールのコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|コンテンツ コントロールとそのコンテンツを削除します。|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|コンテンツ コントロールのテキストの書式設定を取得します。|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|コンテンツ コントロール オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|コンテンツ コントロールの識別子を表す整数値を取得します。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|コンテンツ コントロールに含まれる inlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|メイン文書の指定した位置に、区切りを挿入します。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|指定した場所にあるコンテンツ コントロールにドキュメントを挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|コンテンツ コントロール内の指定された位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|指定した場所にあるコンテンツ コントロールに OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|コンテンツ コントロール内の指定された位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|コンテンツ コントロール内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|コンテンツ コントロールのプレースホルダー テキストを取得または設定します。|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|コンテンツ コントロールを編集後に削除できるかどうかを示す値を取得または設定します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|コンテンツ コントロール オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|コンテンツ コントロールを選択します。|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|コンテンツ コントロールのスタイル名を取得または設定します。|
||[タグ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|コンテンツ コントロールを識別するタグを取得または設定します。|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|コンテンツ コントロールのテキストを取得します。|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|コンテンツ コントロールのタイトルを取得または設定します。|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|コンテンツ コントロールの種類を取得します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|指定されたタグを含むコンテンツ コントロールを取得します。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|指定されたタイトルを含むコンテンツ コントロールを取得します。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|コレクション内のインデックスによってコンテンツ コントロールを取得します。|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|メイン ドキュメントの body オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|ドキュメント内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getselection-member(1))|ドキュメントの現在の選択範囲を取得します。|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|ドキュメントを保存します。|
||[保存済み](/javascript/api/word/word.document#word-word-document-saved-member)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|ドキュメント内のセクション オブジェクトのコレクションを取得します。|
|[フォント](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|フォントが太字かどうかを示す値を取得または設定します。|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|指定されたフォントの色を取得または設定します。|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|フォントに二重取り消し線が設定されているかどうかを示す値を取得または設定します。|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|強調表示の色を取得または設定します。|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|フォントが斜体かどうかを示す値を取得または設定します。|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|フォント名を表す値を取得または設定します。|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|フォント サイズをポイント単位で表す値を取得または設定します。|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|フォントに取り消し線が設定されているかどうかを示す値を取得または設定します。|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|フォントが下付き文字かどうかを示す値を取得または設定します。|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|フォントが上付き文字かどうかを示す値を取得または設定します。|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|フォントの下線の種類を示す値を取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|インライン イメージに関連付けられた代替テキストを表す文字列を取得または設定します。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|インライン画像のタイトルを含む文字列を取得または設定します。|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|インライン画像の base64 エンコード文字列形式を取得します。|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|インライン画像の高さを表す数値を取得するか設定します。|
||[hyperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|イメージ上のハイパーリンクを取得または設定します。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|リッチ テキストのコンテンツ コントロールでインライン画像をラップします。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrol-member)|インライン画像を含むコンテンツ コントロールを取得します。|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|インライン画像の幅を表す数値を取得するか設定します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[配置](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|段落の配置を取得または設定します。|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|段落オブジェクトの内容をクリアします。|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|段落内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|文書から段落と、その段落の内容を削除します。|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|最初の行またはぶら下げインデントの値をポイントで取得または設定します。|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|段落のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|段落オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|段落内の InlinePicture オブジェクトのコレクションを取得します。|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|指定した場所の段落にドキュメントを挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|段落の指定した位置に、HTML を挿入します。|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|段落の指定した位置に、図を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|指定した場所の段落に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|段落の指定した位置に、テキストを挿入します。|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|段落の左インデントの値をポイント数単位で取得または設定します。|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|段落の行間をポイント数単位で取得または設定します。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|段落の後のグリッド線の間隔の量を取得または設定します。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|段落前の間隔の幅をグリッド線数単位で取得または設定します。|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|段落のアウトライン レベルを取得または設定します。|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|段落を格納しているコンテンツ コントロールを取得します。|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|段落の右インデントの値をポイント数単位で取得または設定します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Paragraph オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|段落を選択して、その段落に Word の UI を移動します。|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|段落後の間隔をポイント数単位で取得または設定します。|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|段落前の間隔をポイント数単位で取得または設定します。|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|段落のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|段落のテキストを取得します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|範囲オブジェクトの内容をクリアします。|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|範囲内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|文書から範囲と、その範囲の内容を削除します。|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|範囲のテキスト形式を取得します。|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|範囲オブジェクトの HTML 表記を取得します。|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|Range オブジェクトの OOXML 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|メイン文書の指定した位置に、区切りを挿入します。|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|指定した位置に文書を挿入します。|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|指定した位置に HTML を挿入します。|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|指定した位置に OOXML を挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|指定した位置にテキストを挿入します。|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|範囲内の段落オブジェクトのコレクションを取得します。|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|範囲を格納するコンテンツ コントロールを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.range#word-word-range-search-member(1))|range オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|範囲を選択して、その範囲に Word の UI を移動します。|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|範囲のスタイル名を取得または設定します。|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|範囲のテキストを取得します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|単語間のすべての区切り記号を無視するかどうかを示す値を取得または設定します。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|検索文字列で終わる語句と一致するかどうかを示す値を取得または設定します。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|長い単語の一部ではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|セクションの body オブジェクトを取得します。|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|セクションのフッターの 1 つを取得します。|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getheader-member(1))|セクションのヘッダーの 1 つを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
