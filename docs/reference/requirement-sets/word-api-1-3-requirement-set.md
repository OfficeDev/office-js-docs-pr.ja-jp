---
title: Word JavaScript API 要件セット1.3
description: WordApi 1.3 要件セットの詳細
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: fe72a3047fdbdd719fd115858e4010fbc2c639e5
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268560"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 の新機能

WordApi 1.3 コンテンツコントロール、カスタム XML、およびドキュメントレベルの設定のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット1.3 の Api を示します。 Word JavaScript API 要件セット1.3 またはそれ以前のバージョンでサポートされているすべての Api の API リファレンスドキュメントを表示するには、「[要件セット1.3 またはそれ以前の Word api](/javascript/api/word?view=word-js-1.3)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|オプションの base64 でエンコードされた .docx ファイルを使用して、新しいドキュメントを作成します。|
|[Body](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|範囲として、本文全体、あるいは本文の開始点または終了点を取得します。|
||[insertTable (rowCount: number, columnCount: number, Inserttable: Word Inserttable, values?: string [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。 insertLocation の値には、'Start' または 'End' を指定できます。|
||[サイト](/javascript/api/word/word.body#lists)|本文に含まれるリスト オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentBody](/javascript/api/word/word.body#parentbody)|本文の親の本文を取得します。たとえば、テーブル セル本文の親本文にはヘッダーを指定できます。親本文がない場合は、スローします。読み取り専用。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|本文の親の本文を取得します。たとえば、テーブル セル本文の親本文にはヘッダーを指定できます。親本文がない場合は、null オブジェクトを返します。読み取り専用。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|本文を含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentSection](/javascript/api/word/word.body#parentsection)|本文の親セクションを取得します。 親セクションが存在しない場合にスローされます。 読み取り専用です。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|本文の親セクションを取得します。 親セクションが存在しない場合は、null オブジェクトを返します。 読み取り専用です。|
||[テーブル](/javascript/api/word/word.body#tables)|本文に含まれるテーブル オブジェクトのコレクションを取得します。 読み取り専用です。|
||[type](/javascript/api/word/word.body#type)|本文の種類を取得します。 種類は、'MainDoc'、'Section'、'Header'、'Footer'、または 'TableCell' にできます。 読み取り専用です。|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|本文の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|範囲として、コンテンツ コントロール全体、あるいはコンテンツ コントロールの開始点または終了点を取得します。|
||[getTextRanges (endingMarks: string [], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|句読点やその他の終了マークを使用して、コンテンツコントロール内のテキスト範囲を取得します。|
||[insertTable (rowCount: number, columnCount: number, Inserttable: Word Inserttable, values?: string [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを、コンテンツ コントロール内またはコンテンツ コントロールの横に挿入します。 InsertLocation の値には、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[サイト](/javascript/api/word/word.contentcontrol#lists)|コンテンツ コントロールに含まれるリスト オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|コンテンツ コントロールの親の本文を取得します。 読み取り専用です。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|コンテンツ コントロールを含むテーブルを取得します。 テーブルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|コンテンツ コントロールを含むテーブル セルを取得します。 テーブルのセルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|コンテンツ コントロールを含むテーブル セルを取得します。 テーブル セルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|コンテンツ コントロールを含むテーブルを取得します。 テーブルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[サブ](/javascript/api/word/word.contentcontrol#subtype)|コンテンツ コントロールのサブタイプを取得します。 リッチ テキスト コンテンツ コントロールの場合、サブタイプは、'RichTextInline'、'RichTextParagraphs'、'RichTextTableCell'、'RichTextTableRow' および 'RichTextTable' にできます。 読み取り専用です。|
||[テーブル](/javascript/api/word/word.contentcontrol#tables)|コンテンツ コントロールに含まれるテーブル オブジェクトのコレクションを取得します。 読み取り専用。|
||[split (区切り文字: string [], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|区切り記号を使用して、コンテンツ コントロールを子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|コンテンツ コントロールの組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。 このコレクション内の識別子を持つコンテンツコントロールがない場合は、null オブジェクトを返します。|
||[getByTypes (types: Word ContentControlType [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|指定した種類またはサブタイプのコンテンツコントロールを取得します。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|このコレクション内の最初のコンテンツ コントロールを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|このコレクション内の最初のコンテンツ コントロールを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/word/word.customproperty#key)|カスタム プロパティのキーを取得します。 読み取り専用です。|
||[type](/javascript/api/word/word.customproperty#type)|カスタム プロパティの値の型を取得します。 可能な値は、String、Number、Date、Boolean です。 読み取り専用です。|
||[value](/javascript/api/word/word.customproperty#value)|カスタム プロパティの値を取得または設定します。 Web 上の Word および .docx ファイル形式では、これらのプロパティを任意に長くすることができますが、Word のデスクトップ版では文字列値が 255 16 ビット文字に切り捨てられます (サロゲートペアを分割することで、無効な unicode を作成する可能性があります)。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add (key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタムプロパティが存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/word/word.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[プロパティ](/javascript/api/word/word.document#properties)|ドキュメントのプロパティを取得します。 読み取り専用です。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open ()](/javascript/api/word/word.documentcreated#open--)|図面を開きます。|
||[本文](/javascript/api/word/word.documentcreated#body)|文書の本文オブジェクトを取得します。 本文は、ヘッダー、フッター、脚注、テキストボックスなどを除いたテキストです。 読み取り専用です。|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|文書内のコンテンツコントロールオブジェクトのコレクションを取得します。 これには、文書、ヘッダー、フッター、テキストボックスなどの本文にコンテンツコントロールが含まれます。 読み取り専用です。|
||[プロパティ](/javascript/api/word/word.documentcreated#properties)|ドキュメントのプロパティを取得します。 読み取り専用です。|
||[更新](/javascript/api/word/word.documentcreated#saved)|ドキュメント内の変更が保存されているかどうかを示します。値 true は、ドキュメントが保存されてから変更されていないことを示します。読み取り専用です。|
||[sections](/javascript/api/word/word.documentcreated#sections)|ドキュメント内の section オブジェクトのコレクションを取得します。 読み取り専用です。|
||[save()](/javascript/api/word/word.documentcreated#save--)|ドキュメントを保存します。 ここでは、ドキュメントが保存されたことがない場合は、Word の既定のファイルの名前付け規則を使用します。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[判別](/javascript/api/word/word.documentproperties#author)|ドキュメントの作成者を取得または設定します。|
||[項目](/javascript/api/word/word.documentproperties#category)|ドキュメントのカテゴリを取得または設定します。|
||[comments](/javascript/api/word/word.documentproperties#comments)|ドキュメントのコメントを取得または設定します。|
||[company](/javascript/api/word/word.documentproperties#company)|ドキュメントの会社を取得または設定します。|
||[format](/javascript/api/word/word.documentproperties#format)|ドキュメントの書式設定を取得または設定します。|
||[キーワード](/javascript/api/word/word.documentproperties#keywords)|ドキュメントのキーワードを取得または設定します。|
||[manager](/javascript/api/word/word.documentproperties#manager)|ドキュメントのマネージャーを取得または設定します。|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|ドキュメントのアプリケーション名を取得します。 読み取り専用です。|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|ドキュメントの作成日を取得します。 読み取り専用です。|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|ドキュメントのカスタム プロパティのコレクションを取得します。 読み取り専用です。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|ドキュメントの最後の作成者を取得します。 読み取り専用です。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|ドキュメントを最後に印刷した日を取得します。 読み取り専用。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|ドキュメントを最後に保存した時刻を取得します。 読み取り専用です。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|ドキュメントのリビジョン番号を取得します。 読み取り専用です。|
||[security](/javascript/api/word/word.documentproperties#security)|ドキュメントのセキュリティを取得します。 読み取り専用です。|
||[template](/javascript/api/word/word.documentproperties#template)|ドキュメントのテンプレートを取得します。 読み取り専用です。|
||[subject](/javascript/api/word/word.documentproperties#subject)|ドキュメントの件名を取得または設定します。|
||[title](/javascript/api/word/word.documentproperties#title)|ドキュメントのタイトルを取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ()](/javascript/api/word/word.inlinepicture#getnext--)|次のインライン画像を取得します。 このインライン画像が最後にある場合は、例外をスローします。|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|次のインライン画像を取得します。 このインライン画像が最後にある場合は、null オブジェクトを返します。|
||[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|範囲として、画像、あるいは画像の開始点または終了点を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|インライン画像を含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|インライン イメージを含むテーブルを取得します。 テーブルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|インライン イメージを含むテーブルのセルを取得します。 テーブルのセルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|インライン イメージを含むテーブルのセルを取得します。 テーブル セルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|インライン イメージを含むテーブルを取得します。 テーブルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|このコレクション内の最初のインライン イメージを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|このコレクション内の最初のインライン イメージを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|リスト内の指定したレベルで発生する段落を取得します。|
||[getLevelString (level: number)](/javascript/api/word/word.list#getlevelstring-level-)|指定したレベルで行頭文字、番号、または画像を文字列として取得します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 InsertLocation の値には、' Start '、' End '、' Before '、または ' After ' を指定できます。|
||[id](/javascript/api/word/word.list#id)|リストの id を取得します。|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|リスト内に 9 つの各レベルが存在するかどうかを確認します。値が true の場合は、レベルが存在することを示します。つまり、そのレベルに少なくとも 1 つのリスト アイテムがあることを意味します。読み取り専用。|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|リスト内の 9 レベルのすべての種類を取得します。 各種類は、' Bullet '、' Number '、または ' Picture ' にすることができます。 読み取り専用です。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|リスト内の段落を取得します。 読み取り専用です。|
||[setLevelAlignment (level: number, alignment: Word. 配置)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|リスト内の指定したレベルで行頭文字の配置、番号、画像のいずれかを設定します。|
||[setLevelBullet (level: number, listBullet: Word. ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|リスト内の指定したレベルで行頭文字の書式を設定します。 行頭文字が 'Custom' の場合は、charCode が必要です。|
||[setLevelIndents (level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|リスト内の指定したレベルの 2 つのインデントを設定します。|
||[setLevelNumbering (level: number, listNumbering: Word. ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|リスト内の指定したレベルで番号付け書式を設定します。|
||[setLevelStartingNumber (level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|リスト内の指定したレベルで開始番号を設定します。 既定値は 1 です。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|識別子を使用してリストを取得します。 このコレクションに識別子のリストがない場合は、例外をスローします。|
||[getByIdOrNullObject (id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|識別子を使用してリストを取得します。 このコレクションに識別子が含まれているリストがない場合は、null オブジェクトを返します。|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|このコレクション内の最初のリストを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|このコレクション内の最初のリストを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|コレクション内のインデックスを使用して、リスト オブジェクトを取得します。|
||[items](/javascript/api/word/word.listcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。 リストアイテムが祖先を持たない場合にスローされます。|
||[getAncestorOrNullObject (parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。 リストアイテムに祖先がない場合は、null オブジェクトを返します。|
||[getDescendants (directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|リスト アイテムのすべての子孫のリスト アイテムを取得します。|
||[level](/javascript/api/word/word.listitem#level)|リスト内のアイテムのレベルを取得または設定します。|
||[listString](/javascript/api/word/word.listitem#liststring)|リストアイテムの行頭文字、番号、または画像を文字列として取得します。 読み取り専用です。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|兄弟を基準にしてリスト アイテムの注文番号を取得します。 読み取り専用。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|指定したレベルで段落を既存のリストに結合させます。段落をリストに結合できない場合、または段落が既にリスト アイテムである場合は、失敗します。|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|段落がリスト アイテムである場合は、この段落をリストから移動します。|
||[getNext ()](/javascript/api/word/word.paragraph#getnext--)|次の段落を取得します。 段落が最後のものである場合は、例外をスローします。|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|次の段落を取得します。 段落が最後のものである場合は、null オブジェクトを返します。|
||[getPrevious ()](/javascript/api/word/word.paragraph#getprevious--)|前の段落を取得します。 段落が最初のものである場合は、例外をスローします。|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|前の段落を取得します。 段落が最初のものである場合は、null オブジェクトを返します。|
||[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|段落全体、あるいは段落の開始点または終了点を範囲として取得します。|
||[getTextRanges (endingMarks: string [], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|句読点やその他の終了記号を使用して、段落内のテキスト範囲を取得します。|
||[insertTable (rowCount: number, columnCount: number, Inserttable: Word Inserttable, values?: string [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。 insertLocation の値には、'Before' または 'After' を指定できます。|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|段落がその親の本文内の最後の段落であることを示します。 読み取り専用です。|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|段落がリスト アイテムであるかどうかを確認します。 読み取り専用です。|
||[list](/javascript/api/word/word.paragraph#list)|この段落が属するリストを取得します。 段落がリストに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[listItem](/javascript/api/word/word.paragraph#listitem)|段落の ListItem を取得します。 段落がリストに含まれていない場合にスローされます。 読み取り専用です。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|段落の ListItem を取得します。 段落がリストの一部でない場合は、null オブジェクトを返します。 読み取り専用です。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|この段落が属するリストを取得します。 段落がリスト内にない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|段落の親の本文を取得します。 読み取り専用。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|段落を格納しているコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|段落を含むテーブルを取得します。 テーブルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|段落を含むテーブルのセルを取得します。 テーブルのセルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|段落を含むテーブルのセルを取得します。 テーブル セルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|段落を含むテーブルを取得します。 テーブルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|段落のテーブルのレベルを取得します。 段落がテーブル内にない場合は、0 を返します。 読み取り専用です。|
||[split (区切り文字: string [], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|区切り記号を使用して、段落を子の範囲に分割します。|
||[startNewList ()](/javascript/api/word/word.paragraph#startnewlist--)|この段落を含む新しいリストを開始します。 段落が既にリスト アイテムである場合は失敗します。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|段落の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|このコレクション内の最初の段落を取得します。 コレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|このコレクション内の最初の段落を取得します。 コレクションが空の場合は、null オブジェクトを返します。|
||[getLast ()](/javascript/api/word/word.paragraphcollection#getlast--)|このコレクション内の最後の段落を取得します。 コレクションが空の場合にスローされます。|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|このコレクション内の最後の段落を取得します。 コレクションが空の場合は、null オブジェクトを返します。|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (range: Word Range)](/javascript/api/word/word.range#comparelocationwith-range-)|この範囲の場所を別の範囲の場所と比較します。|
||[expandTo (range: Word Range)](/javascript/api/word/word.range#expandto-range-)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。 この範囲は変更されません。 2つの範囲がユニオンを持たない場合にスローされます。|
||[expandToOrNullObject (範囲: Word)](/javascript/api/word/word.range#expandtoornullobject-range-)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。 この範囲は変更されません。 2つの範囲がユニオンを持たない場合は、null オブジェクトを返します。|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|範囲内のハイパーリンクの子の範囲を取得します。|
||[getNextTextRange (endingMarks: string [], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|句読点やその他の終了記号を使用して、次のテキスト範囲を取得します。 このテキスト範囲が最後のものである場合は、例外をスローします。|
||[getNextTextRangeOrNullObject (endingMarks: string [], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|句読点やその他の終了記号を使用して、次のテキスト範囲を取得します。 このテキスト範囲が最後のものである場合は、null オブジェクトを返します。|
||[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|範囲の複製を作成するか、新しい範囲として開始点または終了点を取得します。|
||[getTextRanges (endingMarks: string [], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|句読点やその他の終了記号を使用して、範囲内のテキストの子の範囲を取得します。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。 範囲に新しいハイパーリンクを設定すると、範囲内のすべてのハイパーリンクが削除されます。 省略可能な location パーツから address パーツを区切るには、' # ' を使用します。|
||[insertTable (rowCount: number, columnCount: number, Inserttable: Word Inserttable, values?: string [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。 insertLocation の値には、'Before' または 'After' を指定できます。|
||[intersectWith (範囲: Word)](/javascript/api/word/word.range#intersectwith-range-)|別の範囲とこの範囲の交点として、新しい範囲を返します。 この範囲は変更されません。 2つの範囲が重なっていない場合、または隣接していない場合にスローされます。|
||[intersectWithOrNullObject (範囲: Word)](/javascript/api/word/word.range#intersectwithornullobject-range-)|別の範囲とこの範囲の交点として、新しい範囲を返します。 この範囲は変更されません。 2つの範囲が重なっていないか隣接していない場合は、null オブジェクトを返します。|
||[isEmpty](/javascript/api/word/word.range#isempty)|範囲の長さが 0 であるかどうかを確認します。 読み取り専用です。|
||[サイト](/javascript/api/word/word.range#lists)|範囲内のリスト オブジェクトのコレクションを取得します。 読み取り専用です。|
||[parentBody](/javascript/api/word/word.range#parentbody)|範囲の親の本文を取得します。 読み取り専用です。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|範囲を格納するコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.range#parenttable)|範囲を含むテーブルを取得します。 テーブルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|範囲を含むテーブルのセルを取得します。 テーブルのセルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|範囲を含むテーブルのセルを取得します。 テーブル セルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|範囲を含むテーブルを取得します。 テーブルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[テーブル](/javascript/api/word/word.range#tables)|範囲内のテーブル オブジェクトのコレクションを取得します。 読み取り専用。|
||[split (区切り文字: string [], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|区切り記号を使用して、範囲を子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|範囲の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|このコレクション内の最初の範囲を取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|このコレクション内の最初の範囲を取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
|[Section](/javascript/api/word/word.section)|[getNext ()](/javascript/api/word/word.section#getnext--)|次のセクションを取得します。 このセクションが最後のセクションの場合は、例外をスローします。|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|次のセクションを取得します。 このセクションが最後のセクションの場合は、null オブジェクトを返します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|このコレクション内の最初のセクションを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|このコレクション内の最初のセクションを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation: Word InsertLocation, columnCount: number, values?: string [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|最初または最後の既存の列をテンプレートとして使用して、テーブルの最初または最後に列を追加します。これは、統一されたテーブルに適用可能です。指定すると、文字列値は新しく挿入された行に設定されます。|
||[addRows (insertLocation: Word. InsertLocation, rowCount: number, values?: string [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|最初または最後の既存の行をテンプレートとして使用して、テーブルの最初または最後に行を追加します。指定すると、文字列値は新しく挿入された行に設定されます。|
||[策定](/javascript/api/word/word.table#alignment)|ページの列に対するテーブルの配置を取得または設定します。 値には、' Left '、' センタリング '、または ' Right ' を指定できます。|
||[autoFitWindow ()](/javascript/api/word/word.table#autofitwindow--)|テーブルの列をウィンドウの幅に合わせて自動調整します。|
||[clear()](/javascript/api/word/word.table#clear--)|テーブルの内容をクリアします。|
||[delete()](/javascript/api/word/word.table#delete--)|テーブル全体を削除します。|
||[deleteColumns (columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|特定の列を削除します。 これは、統一されたテーブルに適用可能です。|
||[deleteRows (rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|特定の行を削除します。|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|列の幅を揃えます。 これは、統一されたテーブルに適用可能です。|
||[getBorder (borderLocation: Word BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|指定した罫線の罫線スタイルを取得します。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|指定された行と列のテーブル セルを取得します。 指定した表のセルが存在しない場合にスローされます。|
||[getCellOrNullObject (rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|指定された行と列のテーブル セルを取得します。 指定したテーブルセルが存在しない場合は、null オブジェクトを返します。|
||[getCellPadding (cellPaddingLocation: CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext ()](/javascript/api/word/word.table#getnext--)|次のテーブルを取得します。 このテーブルが最後のものである場合は、例外をスローします。|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|次のテーブルを取得します。 このテーブルが最後のものである場合は、null オブジェクトを返します。|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|テーブルの後の段落を取得します。 表の後に段落がない場合にスローされます。|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|テーブルの後の段落を取得します。 表の後に段落がない場合は、null オブジェクトを返します。|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|テーブルの前の段落を取得します。 表の前に段落がない場合にスローされます。|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|テーブルの前の段落を取得します。 表の前に段落がない場合は、null オブジェクトを返します。|
||[getRange (rangeLocation?: Word RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|このテーブルを含む範囲、あるいはテーブルの開始または終了の範囲を取得します。|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|ヘッダー行の数を取得および設定します。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|テーブル内のすべてのセルの水平方向の配置を取得および設定します。 値は、' Left '、' センタリング '、' Right '、または ' ジャスティファイ ' にすることができます。|
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|テーブルにコンテンツ コントロールを挿入します。|
||[insertParagraph (paragraphText: string, Insertparagraph: Word. Insertparagraph)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[insertTable (rowCount: number, columnCount: number, Inserttable: Word Inserttable, values?: string [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。 有効な insertLocation の値は、'Before' または 'After' です。|
||[font](/javascript/api/word/word.table#font)|フォントを取得します。 これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。 読み取り専用です。|
||[isUniform](/javascript/api/word/word.table#isuniform)|すべてのテーブル行が均一かどうかを示します。 読み取り専用です。|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|テーブルの入れ子のレベルを取得します。 最上位のテーブルのレベルは、レベル 1 です。 読み取り専用です。|
||[parentBody](/javascript/api/word/word.table#parentbody)|テーブルの親の本文を取得します。 読み取り専用です。|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|テーブルを含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合にスローされます。 読み取り専用です。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|テーブルを含むコンテンツ コントロールを取得します。 親コンテンツコントロールがない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.table#parenttable)|このテーブルを含むテーブルを取得します。 テーブルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|このテーブルを含むテーブルのセルを取得します。 テーブルのセルに含まれていない場合は、例外をスローします。 読み取り専用です。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|このテーブルを含むテーブルのセルを取得します。 テーブル セルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|このテーブルを含むテーブルを取得します。 テーブルに含まれていない場合は、null オブジェクトを返します。 読み取り専用です。|
||[rowCount](/javascript/api/word/word.table#rowcount)|表の行数を取得します。 読み取り専用です。|
||[rows](/javascript/api/word/word.table#rows)|すべてのテーブルの行を取得します。 読み取り専用です。|
||[テーブル](/javascript/api/word/word.table#tables)|1 レベル深く入れ子にされた子テーブルを取得します。 読み取り専用。|
||[search (searchText: string, searchOptions?: Word SearchOptions](/javascript/api/word/word.table#search-searchtext--searchoptions-)|Table オブジェクトの範囲に対して、指定した SearchOptions を使用して検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|テーブル、あるいはテーブルの開始位置または終了位置を選択して、Word の UI に移動します。|
||[setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|網かけの色を取得および設定します。 色は、"#RRGGBB" 形式で指定するか、色の名前を使用して指定します。|
||[style](/javascript/api/word/word.table#style)|テーブルのスタイル名を取得または設定します。カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|テーブルの列を縞模様にするかどうかを取得および設定します。|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|テーブルの行を縞模様にするかどうかを取得および設定します。|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|テーブルの組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|テーブルの最初の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|テーブルの最後の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|テーブルの集計 (最後) 行に特別なスタイルを指定するかどうかを取得および設定します。|
||[values](/javascript/api/word/word.table#values)|2D の Javascript 配列として、テーブルのテキスト値を取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|テーブル内のすべてのセルの垂直方向の配置を取得および設定します。 値には、' Top '、' Center '、または ' Bottom ' を指定できます。|
||[width](/javascript/api/word/word.table#width)|テーブルの幅をポイント単位で取得および設定します。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|表の罫線の色を取得または設定します。|
||[type](/javascript/api/word/word.tableborder#type)|テーブルの罫線の種類を取得または設定します。|
||[width](/javascript/api/word/word.tableborder#width)|テーブルの罫線の幅をポイント単位で得または設定します。 幅が固定されているテーブルの罫線の種類には適用できません。|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|セルの列の幅をポイント単位で取得または設定します。 これは、統一されたテーブルに適用可能です。|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|このセルを含む列を削除します。 これは、統一されたテーブルに適用可能です。|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|このセルを含む行を削除します。|
||[getBorder (borderLocation: Word BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|指定した罫線の罫線スタイルを取得します。|
||[getCellPadding (cellPaddingLocation: CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext ()](/javascript/api/word/word.tablecell#getnext--)|次のセルを取得します。 このセルが最後のセルである場合は、スローします。|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|次のセルを取得します。 このセルが最後のセルの場合は、null オブジェクトを返します。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|セルの水平方向の配置を取得および設定します。 値は、' Left '、' センタリング '、' Right '、または ' ジャスティファイ ' にすることができます。|
||[insertColumns (Insertcolumns: Word Insertcolumns, columnCount: number, values?: string [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|セルの列をテンプレートとして使用して、列をセルの左または右に追加します。これは、統一されたテーブルに適用可能です。指定すると、文字列値は新しく挿入された行に設定されます。|
||[insertRows (Insertrows: Word Insertrows, rowCount: number, values?: string [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|セルの行をテンプレートとして使用して、行をセルの上または下に挿入します。指定すると、文字列値は新しく挿入された行に設定されます。|
||[本文](/javascript/api/word/word.tablecell#body)|セルの本文オブジェクトを取得します。 読み取り専用です。|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|その行のセルのインデックスを取得します。 読み取り専用です。|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|セルの親行を取得します。 読み取り専用です。|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|セルの親テーブルを取得します。 読み取り専用。|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|テーブルのセル行のインデックスを取得します。 読み取り専用です。|
||[width](/javascript/api/word/word.tablecell#width)|セルの幅をポイント単位で取得します。 読み取り専用です。|
||[setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|セルの網かけの色を取得または設定します。 色は、"#RRGGBB" 形式で指定するか、色の名前を使用して指定します。|
||[value](/javascript/api/word/word.tablecell#value)|セルのテキストを取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|セルの垂直方向の配置を取得および設定します。 値には、' Top '、' Center '、または ' Bottom ' を指定できます。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|このコレクション内の最初のテーブル セルを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|このコレクション内の最初のテーブル セルを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
||[items](/javascript/api/word/word.tablecellcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|このコレクション内の最初のテーブルを取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|このコレクション内の最初のテーブルを取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
||[items](/javascript/api/word/word.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|行の内容をクリアします。|
||[delete()](/javascript/api/word/word.tablerow#delete--)|行全体を削除します。|
||[getBorder (borderLocation: Word BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|行のセルの罫線スタイルを取得します。|
||[getCellPadding (cellPaddingLocation: CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext ()](/javascript/api/word/word.tablerow#getnext--)|次の行を取得します。 この行が最後の行である場合にスローします。|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|次の行を取得します。 この行が最後の行の場合は、null オブジェクトを返します。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|行のすべてのセルの水平方向の配置を取得および設定します。 値は、' Left '、' センタリング '、' Right '、または ' ジャスティファイ ' にすることができます。|
||[insertRows (Insertrows: Word Insertrows, rowCount: number, values?: string [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|この行をテンプレートとして使用して、行を挿入します。 値を指定すると、新しい行に値を挿入します。|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|適切な行の高さをポイント単位で取得および設定します。|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|行のセルの数を取得します。 読み取り専用です。|
||[cells](/javascript/api/word/word.tablerow#cells)|セルを取得します。 読み取り専用です。|
||[font](/javascript/api/word/word.tablerow#font)|フォントを取得します。 これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。 読み取り専用です。|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|行がヘッダー行であるかどうかを確認します。 読み取り専用。 ヘッダー行の数を設定するには、テーブル オブジェクトの HeaderRowCount を使用します。|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|親テーブルを取得します。 読み取り専用。|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|親テーブル内の行のインデックスを取得します。 読み取り専用です。|
||[search (searchText: string, searchOptions?: Word SearchOptions)](/javascript/api/word/word.tablerow#search-searchtext--searchoptions-)|指定した SearchOptions を使用して、行の範囲に基づいて検索を実行します。 検索結果は、範囲オブジェクトのコレクションです。|
||[select (selectionMode?:. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|行を選択し、その行に Word の UI を移動します。|
||[setCellPadding (cellPaddingLocation: CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|網かけの色を取得および設定します。 色は、"#RRGGBB" 形式で指定するか、色の名前を使用して指定します。|
||[values](/javascript/api/word/word.tablerow#values)|2D の Javascript 配列として、行のテキスト値を取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|行のセルの垂直方向の配置を取得および設定します。 値には、' Top '、' Center '、または ' Bottom ' を指定できます。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|このコレクション内の最初の行を取得します。 このコレクションが空の場合にスローされます。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|このコレクション内の最初の行を取得します。 このコレクションが空の場合は、null オブジェクトを返します。|
||[items](/javascript/api/word/word.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンスドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
