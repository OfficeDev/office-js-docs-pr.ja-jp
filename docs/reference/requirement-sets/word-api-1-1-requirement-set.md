---
title: Word JavaScript API 要件セット1.1
description: WordApi 1.1 要件セットの詳細
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: a2839a2553d42701956fd2e75a86564c133d9a93
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064915"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 の新機能

WordApi 1.1 は、Word JavaScript API の最初の要件セットです。 Word 2016 でサポートされている唯一の Word API 要件セットです。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット1.1 の Api を示します。 Word JavaScript API 要件セット1.1 でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「[要件セット1.1 の Word api](/javascript/api/word?view=word-js-1.1)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|本文オブジェクトの内容を消去します。ユーザーは、消去された内容を元に戻す操作を実行できます。|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Body オブジェクトの HTML 表記を取得します。 Web ページまたは HTML ビューアーでレンダリングされる場合、書式設定は、ドキュメントの書式設定と完全に一致しますが、完全に一致するとは限りません。 このメソッドは、異なるプラットフォーム (Windows、Mac など) の同じドキュメントに対して、まったく同じ HTML を返しません。 厳密な忠実性、または複数のプラットフォーム間で`Body.getOoxml()`の一貫性が必要な場合は、を使用して、返された XML を HTML に変換します。|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|本文オブジェクトの OOXML (Office オープン XML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。 insertLocation の値には、'Start' または 'End' を指定できます。|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|リッチ テキスト コンテンツ コントロールで本文オブジェクトをラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|文書を本文の指定された位置に挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|指定した位置に HTML を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|指定した位置に OOXML を挿入します。  insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 insertLocation の値には、'Start' または 'End' を指定できます。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|テキストを本文の指定された位置に挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|本文に含まれるリッチテキストコンテンツコントロールオブジェクトのコレクションを取得します。 読み取り専用です。|
||[font](/javascript/api/word/word.body#font)|本文のテキスト形式を取得します。 フォント名、サイズ、色、およびその他のプロパティを取得および設定するために使用します。 読み取り専用です。|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|本文にある InlinePicture オブジェクトのコレクションを取得します。 コレクションに浮動イメージは含まれません。 読み取り専用です。|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|本文に含まれる paragraph オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|本文を含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[text](/javascript/api/word/word.body#text)|本文のテキストを取得します。 insertText メソッドを使用して、テキストを挿入します。 読み取り専用です。|
||[search (searchText: string, searchOptions?: Word SearchOptions)](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Body オブジェクトのスコープで、指定された SearchOptions を使用して検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|本文を選択し、その本文に Word の UI を移動します。|
||[style](/javascript/api/word/word.body#style)|本文のスタイル名を取得または設定します。カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[外観](/javascript/api/word/word.contentcontrol#appearance)|コンテンツ コントロールの外観を取得または設定します。 値には、' BoundingBox '、' Tags '、または ' Hidden ' を指定できます。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|ユーザーがコンテンツ コントロールを削除できるかどうかを示す値を取得または設定します。 removeWhenEdited と同時に使用することはできません。|
||[Canメモ Dit](/javascript/api/word/word.contentcontrol#cannotedit)|ユーザーがコンテンツ コントロールのコンテンツを編集できるかどうかを示す値を取得または設定します。|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|コンテンツ コントロールの内容をクリアします。 ユーザーは、消去された内容を元に戻す操作を実行できます。|
||[color](/javascript/api/word/word.contentcontrol#color)|コンテンツ コントロールの色を取得または設定します。 色は、' #RRGGBB ' 形式で指定するか、色名を使用して指定します。|
||[削除 (keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|コンテンツ コントロールとそのコンテンツを削除します。keepContent が true の場合、コンテンツは削除されません。|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|コンテンツコントロールオブジェクトの HTML 表記を取得します。 Web ページまたは HTML ビューアーでレンダリングされる場合、書式設定は、ドキュメントの書式設定と完全に一致しますが、完全に一致するとは限りません。 このメソッドは、異なるプラットフォーム (Windows、Mac など) の同じドキュメントに対して、まったく同じ HTML を返しません。 厳密な忠実性、または複数のプラットフォーム間で`ContentControl.getOoxml()`の一貫性が必要な場合は、を使用して、返された XML を HTML に変換します。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|コンテンツ コントロール オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。 InsertLocation の値には、' Start '、' End '、' Before '、または ' After ' を指定できます。 このメソッドは、' RichTextTable '、' RichTextTableRow '、および ' RichTextTableCell ' のコンテンツコントロールと共に使用することはできません。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|指定した位置にコンテンツコントロールにドキュメントを挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|コンテンツ コントロール内の指定された位置に HTML を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|指定した位置に、コンテンツコントロールに OOXML を挿入します。  insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 InsertLocation の値には、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|コンテンツ コントロール内の指定された位置にテキストを挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|コンテンツ コントロールのプレースホルダー テキストを取得または設定します。 コンテンツ コントロールが空の場合は、淡色のテキストが表示されます。|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|コンテンツ コントロールのコンテンツ コントロール オブジェクトのコレクションを取得します。 読み取り専用です。|
||[font](/javascript/api/word/word.contentcontrol#font)|コンテンツ コントロールのテキストの書式設定を取得します。 これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。 読み取り専用です。|
||[id](/javascript/api/word/word.contentcontrol#id)|コンテンツ コントロールの識別子を表す整数値を取得します。 読み取り専用です。|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|コンテンツ コントロールに含まれる inlinePicture オブジェクトのコレクションを取得します。 コレクションに浮動イメージは含まれません。 読み取り専用です。|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|コンテンツ コントロールにある Paragraph オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[text](/javascript/api/word/word.contentcontrol#text)|コンテンツ コントロールのテキストを取得します。 読み取り専用です。|
||[type](/javascript/api/word/word.contentcontrol#type)|コンテンツ コントロールの種類を取得します。 現在、リッチ テキストのコンテンツ コントロールのみがサポートされています。 読み取り専用です。|
||[removeWhenEdited 済み](/javascript/api/word/word.contentcontrol#removewhenedited)|コンテンツ コントロールを編集後に削除できるかどうかを示す値を取得または設定します。 cannotDelete と同時に使用することはできません。|
||[search (searchText: string, searchOptions?: Word SearchOptions)](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|コンテンツコントロールオブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|コンテンツ コントロールを選択します。 その結果、Word は選択範囲にスクロールされます。|
||[style](/javascript/api/word/word.contentcontrol#style)|コンテンツコントロールのスタイル名を取得または設定します。 カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。 ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|
||[マーク](/javascript/api/word/word.contentcontrol#tag)|コンテンツコントロールを識別するタグを取得または設定します。|
||[title](/javascript/api/word/word.contentcontrol#title)|コンテンツ コントロールのタイトルを取得または設定します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。 このコレクション内の識別子を持つコンテンツコントロールがない場合は、例外をスローします。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|指定されたタグを含むコンテンツ コントロールを取得します。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|指定されたタイトルを含むコンテンツ コントロールを取得します。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|コレクション内のインデックスによってコンテンツコントロールを取得します。|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[getSelection ()](/javascript/api/word/word.document#getselection--)|ドキュメントの現在の選択範囲を取得します。 複数選択はサポートされていません。|
||[本文](/javascript/api/word/word.document#body)|文書の本文オブジェクトを取得します。 本文は、ヘッダー、フッター、脚注、テキストボックスなどを除いたテキストです。 読み取り専用です。|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|文書内のコンテンツコントロールオブジェクトのコレクションを取得します。 これには、文書、ヘッダー、フッター、テキストボックスなどの本文にコンテンツコントロールが含まれます。 読み取り専用です。|
||[更新](/javascript/api/word/word.document#saved)|ドキュメント内の変更が保存されているかどうかを示します。値 true は、ドキュメントが保存されてから変更されていないことを示します。読み取り専用です。|
||[sections](/javascript/api/word/word.document#sections)|ドキュメント内の section オブジェクトのコレクションを取得します。 読み取り専用です。|
||[save()](/javascript/api/word/word.document#save--)|ドキュメントを保存します。 ここでは、ドキュメントが保存されたことがない場合は、Word の既定のファイルの名前付け規則を使用します。|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|フォントが太字かどうかを示す値を取得または設定します。 フォントの書式設定が太字の場合は true、それ以外の場合は false です。|
||[color](/javascript/api/word/word.font#color)|指定されたフォントの色を取得または設定します。 値は、"#RRGGBB" の形式または色の名前で指定できます。|
||[[Doublestrikethrough]](/javascript/api/word/word.font#doublestrikethrough)|フォントに二重取り消し線があるかどうかを示す値を取得または設定します。 フォントの書式が二重取り消し線付きのテキストである場合は true、それ以外の場合は false です。|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|強調表示の色を取得または設定します。 このプロパティを設定するには、' #RRGGBB ' 形式または色名のいずれかの値を使用します。 蛍光ペンの色を削除するには、その色を null に設定します。 強調表示色は、"#RRGGBB" 形式で指定できます。強調表示色が混在している場合は空の文字列、または強調表示色なしの場合は null になります。|
||[italic](/javascript/api/word/word.font#italic)|フォントが斜体かどうかを示す値を取得または設定します。 フォントが斜体の場合は true、それ以外の場合は false です。|
||[name](/javascript/api/word/word.font#name)|フォント名を表す値を取得または設定します。|
||[size](/javascript/api/word/word.font#size)|フォント サイズをポイント単位で表す値を取得または設定します。|
||[打ち消し](/javascript/api/word/word.font#strikethrough)|フォントに取り消し線を表示するかどうかを示す値を取得または設定します。 フォントの書式が取り消し線付きのテキストである場合は true、それ以外の場合は false です。|
||[subscript](/javascript/api/word/word.font#subscript)|フォントが下付き文字かどうかを示す値を取得または設定します。 フォントの書式が下付き文字である場合は true、それ以外の場合は false です。|
||[superscript](/javascript/api/word/word.font#superscript)|フォントが上付き文字かどうかを示す値を取得または設定します。 フォントの書式が上付き文字である場合は true、それ以外の場合は false です。|
||[underline](/javascript/api/word/word.font#underline)|フォントの下線の種類を示す値を取得または設定します。 フォントに下線が付いていない場合は ' None '。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|インライン画像に関連付けられている代替テキストを表す文字列を取得または設定します。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|インライン画像のタイトルを含む文字列を取得または設定します。|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|インライン画像の base64 エンコード文字列形式を取得します。|
||[height](/javascript/api/word/word.inlinepicture#height)|インライン画像の高さを表す数値を取得するか設定します。|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|画像のハイパーリンクを取得または設定します。 省略可能な location パーツから address パーツを区切るには、' # ' を使用します。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|リッチ テキストのコンテンツ コントロールでインライン画像をラップします。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|インライン画像のサイズを変更する際にその元の縦横比を保持するかどうかを示す値を取得または設定します。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|インライン画像を含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[width](/javascript/api/word/word.inlinepicture#width)|インライン画像の幅を表す数値を取得するか設定します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[策定](/javascript/api/word/word.paragraph#alignment)|段落の配置を取得または設定します。 値には、"left"、"centered"、"right"、または "justified" を指定できます。|
||[clear()](/javascript/api/word/word.paragraph#clear--)|段落オブジェクトの内容をクリアします。ユーザーは、消去された内容を元に戻す操作を実行できます。|
||[delete()](/javascript/api/word/word.paragraph#delete--)|文書から段落と、その段落の内容を削除します。|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|最初の行のインデントまたはぶら下げインデントの値をポイント数単位で取得または設定します。最初の行のインデントを設定するには、正の値を使用します。また、ぶら下げインデントを設定するには、負の値を使用します。|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Paragraph オブジェクトの HTML 表記を取得します。 Web ページまたは HTML ビューアーでレンダリングされる場合、書式設定は、ドキュメントの書式設定と完全に一致しますが、完全に一致するとは限りません。 このメソッドは、異なるプラットフォーム (Windows、Mac など) の同じドキュメントに対して、まったく同じ HTML を返しません。 厳密な忠実性、または複数のプラットフォーム間で`Paragraph.getOoxml()`の一貫性が必要な場合は、を使用して、返された XML を HTML に変換します。|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Paragraph オブジェクトの Office Open XML (OOXML) 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|段落オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|指定した位置に段落に文書を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|段落の指定した位置に、HTML を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|段落の指定した位置に、図を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|指定した位置の段落に OOXML を挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|段落の指定した位置に、テキストを挿入します。 insertLocation 値には、'Replace'、'Start'、'End' のいずれかを指定できます。|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|段落の左インデントの値をポイント数単位で取得または設定します。|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|段落の行間をポイント数単位で取得または設定します。 Word UI では、この値が 12 で除算されます。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|段落後の間隔の量 (グリッド線単位) を取得または設定します。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|段落前の間隔の幅をグリッド線数単位で取得または設定します。|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|段落のアウトライン レベルを取得または設定します。|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|段落内のコンテンツコントロールオブジェクトのコレクションを取得します。 読み取り専用です。|
||[font](/javascript/api/word/word.paragraph#font)|段落のテキスト形式を取得します。 これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。 読み取り専用。|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|段落内の InlinePicture オブジェクトのコレクションを取得します。 コレクションに浮動イメージは含まれません。 読み取り専用です。|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|段落を格納しているコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[text](/javascript/api/word/word.paragraph#text)|段落のテキストを取得します。 読み取り専用です。|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|段落の右インデントの値をポイント数単位で取得または設定します。|
||[検索 (searchText: string, searchOptions:: Word SearchOptions})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Paragraph オブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|段落を選択して、その段落に Word の UI を移動します。|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|段落後の間隔をポイント数単位で取得または設定します。|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|段落前の間隔をポイント数単位で取得または設定します。|
||[style](/javascript/api/word/word.paragraph#style)|段落のスタイル名を取得または設定します。 カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。 ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|範囲オブジェクトの内容をクリアします。ユーザーは、クリアしたコンテンツを元に戻す操作を実行できます。|
||[delete()](/javascript/api/word/word.range#delete--)|文書から範囲と、その範囲の内容を削除します。|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Range オブジェクトの HTML 表記を取得します。 Web ページまたは HTML ビューアーでレンダリングされる場合、書式設定は、ドキュメントの書式設定と完全に一致しますが、完全に一致するとは限りません。 このメソッドは、異なるプラットフォーム (Windows、Mac など) の同じドキュメントに対して、まったく同じ HTML を返しません。 厳密な忠実性、または複数のプラットフォーム間で`Range.getOoxml()`の一貫性が必要な場合は、を使用して、返された XML を HTML に変換します。|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Range オブジェクトの OOXML 表記を取得します。|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType: BreakType, Insertbreak: Word Insertbreak)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|メイン文書の指定した位置に、区切りを挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|範囲オブジェクトを、リッチ テキストのコンテンツ コントロールでラップします。|
||[insertFileFromBase64 (base64File: string, insertLocation: Word InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|指定した位置に文書を挿入します。 InsertLocation の値には、' Replace '、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[insertHtml (html: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|指定した位置に HTML を挿入します。 InsertLocation の値には、' Replace '、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[insertOoxml (ooxml: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|指定した位置に OOXML を挿入します。  InsertLocation の値には、' Replace '、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[insertText (text: string, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|指定した位置にテキストを挿入します。 InsertLocation の値には、' Replace '、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|範囲内のコンテンツコントロールオブジェクトのコレクションを取得します。 読み取り専用です。|
||[font](/javascript/api/word/word.range#font)|範囲のテキスト形式を取得します。 これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。 読み取り専用。|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|範囲内の paragraph オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|範囲を格納するコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[text](/javascript/api/word/word.range#text)|範囲のテキストを取得します。 読み取り専用です。|
||[search (searchText: string, searchOptions?: Word SearchOptions)](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Range オブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|範囲を選択して、その範囲に Word の UI を移動します。|
||[style](/javascript/api/word/word.range#style)|範囲のスタイル名を取得または設定します。 カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。 ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|単語間のすべての区切り記号を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに相当します。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|単語間のすべての空白文字を無視するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログボックスの [空白文字を無視する] チェックボックスに対応します。|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログボックスの [大文字と小文字を区別する] チェックボックスに対応します。|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接頭辞に一致する] チェック ボックスに相当します。|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|検索文字列で終わる語句と一致するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接尾辞に一致する] に相当します。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|長い単語の一部ではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [完全に一致する単語だけを検索する] チェック ボックスに相当します。|
||[matchWildCards](/javascript/api/word/word.searchoptions#matchwildcards)||
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに相当します。|
|[Section](/javascript/api/word/word.section)|[getFooter (type: Word Headerfooter Type)](/javascript/api/word/word.section#getfooter-type-)|セクションのフッターの 1 つを取得します。|
||[getHeader (type: Word Headerフッターの種類)](/javascript/api/word/word.section#getheader-type-)|セクションのヘッダーの 1 つを取得します。|
||[本文](/javascript/api/word/word.section#body)|セクションの本文オブジェクトを取得します。 これには、ヘッダー/フッターおよびその他のセクションメタデータは含まれません。 読み取り専用です。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンスドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
