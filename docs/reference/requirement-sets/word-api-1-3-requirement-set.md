---
title: Word JavaScript API 要件セット 1.3
description: WordApi 1.3 要件セットの詳細。
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 0291a8a96e0ae38bf9d1061a09dac3d3b9cc3ddb
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178105"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 の新機能

WordApi 1.3 では、コンテンツ コントロールとドキュメント レベルの設定のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.3 の API の一覧を示します。 Word JavaScript API 要件セット 1.3 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)以前の Word API」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|オプションの base64 エンコードされた .docx ファイルを使用して、新しいドキュメントを作成します。|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|範囲として、本文全体、あるいは本文の開始点または終了点を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。|
||[lists](/javascript/api/word/word.body#lists)|本文に含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.body#parentbody)|本文の親の本文を取得します。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|本文の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|本文を含むコンテンツ コントロールを取得します。|
||[parentSection](/javascript/api/word/word.body#parentsection)|本文の親セクションを取得します。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|本文の親セクションを取得します。|
||[テーブル](/javascript/api/word/word.body#tables)|本文に含まれるテーブル オブジェクトのコレクションを取得します。|
||[type](/javascript/api/word/word.body#type)|本文の種類を取得します。|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|本文の組み込みスタイル名を取得または設定します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|範囲として、コンテンツ コントロール全体、あるいはコンテンツ コントロールの開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|句読点や他の終了記号を使用して、コンテンツ コントロール内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを、コンテンツ コントロール内またはコンテンツ コントロールの横に挿入します。|
||[lists](/javascript/api/word/word.contentcontrol#lists)|コンテンツ コントロールに含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|コンテンツ コントロールの親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|コンテンツ コントロールを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|コンテンツ コントロールを含むテーブルを取得します。|
||[サブタイプ](/javascript/api/word/word.contentcontrol#subtype)|コンテンツ コントロールのサブタイプを取得します。|
||[テーブル](/javascript/api/word/word.contentcontrol#tables)|コンテンツ コントロールに含まれるテーブル オブジェクトのコレクションを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|区切り記号を使用して、コンテンツ コントロールを子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|コンテンツ コントロールの組み込みスタイル名を取得または設定します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|指定した種類またはサブタイプを持つコンテンツ コントロールを取得します。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|このコレクション内の最初のコンテンツ コントロールを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|このコレクション内の最初のコンテンツ コントロールを取得します。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|カスタム プロパティを削除します。|
||[key](/javascript/api/word/word.customproperty#key)|カスタム プロパティのキーを取得します。|
||[type](/javascript/api/word/word.customproperty#type)|カスタム プロパティの値の型を取得します。|
||[value](/javascript/api/word/word.customproperty#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/word/word.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[プロパティ](/javascript/api/word/word.document#properties)|ドキュメントのプロパティを取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|ドキュメントを開きます。|
||[body](/javascript/api/word/word.documentcreated#body)|ドキュメントの body オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|ドキュメント内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[プロパティ](/javascript/api/word/word.documentcreated#properties)|ドキュメントのプロパティを取得します。|
||[保存済み](/javascript/api/word/word.documentcreated#saved)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.documentcreated#sections)|ドキュメント内のセクション オブジェクトのコレクションを取得します。|
||[save()](/javascript/api/word/word.documentcreated#save--)|ドキュメントを保存します。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|ドキュメントの作成者を取得または設定します。|
||[category](/javascript/api/word/word.documentproperties#category)|ドキュメントのカテゴリを取得または設定します。|
||[comments](/javascript/api/word/word.documentproperties#comments)|ドキュメントのコメントを取得または設定します。|
||[company](/javascript/api/word/word.documentproperties#company)|ドキュメントの会社を取得または設定します。|
||[format](/javascript/api/word/word.documentproperties#format)|ドキュメントの書式設定を取得または設定します。|
||[キーワード](/javascript/api/word/word.documentproperties#keywords)|ドキュメントのキーワードを取得または設定します。|
||[上司](/javascript/api/word/word.documentproperties#manager)|ドキュメントのマネージャーを取得または設定します。|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|ドキュメントのアプリケーション名を取得します。|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|ドキュメントの作成日を取得します。|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|ドキュメントのカスタム プロパティのコレクションを取得します。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|ドキュメントの最後の作成者を取得します。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|ドキュメントを最後に印刷した日を取得します。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|ドキュメントを最後に保存した時刻を取得します。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|ドキュメントのリビジョン番号を取得します。|
||[セキュリティ](/javascript/api/word/word.documentproperties#security)|ドキュメントのセキュリティ設定を取得します。|
||[template](/javascript/api/word/word.documentproperties#template)|ドキュメントのテンプレートを取得します。|
||[subject](/javascript/api/word/word.documentproperties#subject)|ドキュメントの件名を取得または設定します。|
||[title](/javascript/api/word/word.documentproperties#title)|ドキュメントのタイトルを取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|次のインライン画像を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|次のインライン画像を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|範囲として、画像、あるいは画像の開始点または終了点を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|インライン画像を含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|インライン イメージを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|インライン イメージを含むテーブルを取得します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|このコレクション内の最初のインライン イメージを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|このコレクション内の最初のインライン イメージを取得します。|
|[リスト](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|リスト内の指定したレベルで発生する段落を取得します。|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|指定したレベルの行頭文字、数値、または図を文字列として取得します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[id](/javascript/api/word/word.list#id)|リストの ID を取得します。|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|9 つの各レベルがリストに存在するかどうかを確認します。|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|リスト内の 9 レベルのすべての種類を取得します。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|リスト内の段落を取得します。|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|リスト内の指定されたレベルでの箇条書き、数字、または図の配置を設定します。|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|リスト内の指定したレベルで行頭文字の書式を設定します。|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|リスト内の指定したレベルの 2 つのインデントを設定します。|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: 配列<文字列 \|>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|リスト内の指定したレベルで番号付け書式を設定します。|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|リスト内の指定したレベルで開始番号を設定します。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|識別子を使用してリストを取得します。|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|識別子を使用してリストを取得します。|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|このコレクション内の最初のリストを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|このコレクション内の最初のリストを取得します。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|コレクション内のインデックスを使用して、リスト オブジェクトを取得します。|
||[items](/javascript/api/word/word.listcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|リスト アイテムのすべての子孫のリスト アイテムを取得します。|
||[level](/javascript/api/word/word.listitem#level)|リスト内のアイテムのレベルを取得または設定します。|
||[listString](/javascript/api/word/word.listitem#liststring)|リスト アイテムの箇条書き、数値、または図を文字列として取得します。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|兄弟を基準にしてリスト アイテムの注文番号を取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|指定したレベルで段落を既存のリストに結合させます。|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|段落がリスト アイテムである場合は、この段落をリストから移動します。|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|次の段落を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|次の段落を取得します。|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|前の段落を取得します。|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|段落全体、あるいは段落の開始点または終了点を範囲として取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|句読点や他の終了記号を使用して、段落内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|段落がその親の本文内の最後の段落であることを示します。|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|段落がリスト アイテムであるかどうかを確認します。|
||[リスト](/javascript/api/word/word.paragraph#list)|この段落が属するリストを取得します。|
||[listItem](/javascript/api/word/word.paragraph#listitem)|段落の ListItem を取得します。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|段落の ListItem を取得します。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|この段落が属するリストを取得します。|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|段落の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|段落を格納しているコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|段落を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|段落を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|段落を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|段落を含むテーブルを取得します。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|段落のテーブルのレベルを取得します。|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|区切り記号を使用して、段落を子の範囲に分割します。|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|この段落を含む新しいリストを開始します。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|段落の組み込みスタイル名を取得または設定します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|このコレクション内の最初の段落を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|このコレクション内の最初の段落を取得します。|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|このコレクション内の最後の段落を取得します。|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|このコレクション内の最後の段落を取得します。|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|この範囲の場所を別の範囲の場所と比較します。|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|範囲内のハイパーリンクの子の範囲を取得します。|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|範囲の複製を作成するか、新しい範囲として開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|句読点や他の終了記号を使用して、範囲内のテキストの子範囲を取得します。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[isEmpty](/javascript/api/word/word.range#isempty)|範囲の長さが 0 であるかどうかを確認します。|
||[lists](/javascript/api/word/word.range#lists)|範囲内のリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.range#parentbody)|範囲の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|範囲を格納するコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.range#parenttable)|範囲を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|範囲を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|範囲を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|範囲を含むテーブルを取得します。|
||[テーブル](/javascript/api/word/word.range#tables)|範囲内のテーブル オブジェクトのコレクションを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|区切り記号を使用して、範囲を子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|範囲の組み込みスタイル名を取得または設定します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|このコレクション内の最初の範囲を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|このコレクション内の最初の範囲を取得します。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api セット: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|次のセクションを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|次のセクションを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|このコレクション内の最初のセクションを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|このコレクション内の最初のセクションを取得します。|
|[表](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|最初または最後の既存の列をテンプレートとして使用して、テーブルの最初または最後に列を追加します。|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|最初または最後の既存の行をテンプレートとして使用して、テーブルの最初または最後に行を追加します。|
||[配置](/javascript/api/word/word.table#alignment)|ページ列に対するテーブルの配置を取得または設定します。|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|テーブルの列をウィンドウの幅に合わせて自動調整します。|
||[clear()](/javascript/api/word/word.table#clear--)|テーブルの内容をクリアします。|
||[delete()](/javascript/api/word/word.table#delete--)|テーブル全体を削除します。|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|特定の列を削除します。|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|特定の行を削除します。|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|列の幅を揃えます。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|指定した罫線の罫線スタイルを取得します。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|指定された行と列のテーブル セルを取得します。|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|指定された行と列のテーブル セルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.table#getnext--)|次のテーブルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|次のテーブルを取得します。|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|テーブルの後の段落を取得します。|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|テーブルの後の段落を取得します。|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|テーブルの前の段落を取得します。|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|テーブルの前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|このテーブルを含む範囲、あるいはテーブルの開始または終了の範囲を取得します。|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|ヘッダー行の数を取得および設定します。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|テーブル内のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|テーブルにコンテンツ コントロールを挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|指定した位置に、段落を挿入します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|指定した数の行と列を含むテーブルを挿入します。|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|フォントを取得します。|
||[isUniform](/javascript/api/word/word.table#isuniform)|すべてのテーブル行が均一かどうかを示します。|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|テーブルの入れ子のレベルを取得します。|
||[parentBody](/javascript/api/word/word.table#parentbody)|テーブルの親の本文を取得します。|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.table#parenttable)|このテーブルを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|このテーブルを含むテーブルを取得します。|
||[rowCount](/javascript/api/word/word.table#rowcount)|表の行数を取得します。|
||[rows](/javascript/api/word/word.table#rows)|すべてのテーブルの行を取得します。|
||[テーブル](/javascript/api/word/word.table#tables)|1 レベル深く入れ子にされた子テーブルを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|テーブル オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|テーブル、あるいはテーブルの開始位置または終了位置を選択して、Word の UI に移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|網かけの色を取得および設定します。|
||[style](/javascript/api/word/word.table#style)|テーブルのスタイル名を取得または設定します。|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|テーブルの列を縞模様にするかどうかを取得および設定します。|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|テーブルの行を縞模様にするかどうかを取得および設定します。|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|テーブルの組み込みスタイル名を取得または設定します。|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|テーブルの最初の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|テーブルの最後の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|テーブルの集計 (最後) 行に特別なスタイルを指定するかどうかを取得および設定します。|
||[values](/javascript/api/word/word.table#values)|2D の Javascript 配列として、テーブルのテキスト値を取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|テーブル内のすべてのセルの垂直方向の配置を取得および設定します。|
||[width](/javascript/api/word/word.table#width)|テーブルの幅をポイント単位で取得および設定します。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|テーブルの罫線の色を取得または設定します。|
||[type](/javascript/api/word/word.tableborder#type)|テーブルの罫線の種類を取得または設定します。|
||[width](/javascript/api/word/word.tableborder#width)|テーブルの罫線の幅をポイント単位で得または設定します。|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|セルの列の幅をポイント単位で取得または設定します。|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|このセルを含む列を削除します。|
||[deleteRow()](/javascript/api/word/word.tablecell#deleterow--)|このセルを含む行を削除します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|指定した罫線の罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablecell#getnext--)|次のセルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|次のセルを取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|セルの水平方向の配置を取得および設定します。|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|セルの列をテンプレートとして使用して、列をセルの左または右に追加します。|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|セルの行をテンプレートとして使用して、行をセルの上または下に挿入します。|
||[body](/javascript/api/word/word.tablecell#body)|セルの本文オブジェクトを取得します。|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|その行のセルのインデックスを取得します。|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|セルの親行を取得します。|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|セルの親テーブルを取得します。|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|テーブルのセル行のインデックスを取得します。|
||[width](/javascript/api/word/word.tablecell#width)|セルの幅をポイント単位で取得します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|セルの網かけの色を取得または設定します。|
||[value](/javascript/api/word/word.tablecell#value)|セルのテキストを取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|セルの垂直方向の配置を取得および設定します。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|このコレクション内の最初のテーブル セルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|このコレクション内の最初のテーブル セルを取得します。|
||[items](/javascript/api/word/word.tablecellcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|このコレクション内の最初のテーブルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|このコレクション内の最初のテーブルを取得します。|
||[items](/javascript/api/word/word.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|行の内容をクリアします。|
||[delete()](/javascript/api/word/word.tablerow#delete--)|行全体を削除します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|行のセルの罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablerow#getnext--)|次の行を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|次の行を取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|行のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|この行をテンプレートとして使用して、行を挿入します。|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|適切な行の高さをポイント単位で取得および設定します。|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|行のセルの数を取得します。|
||[セル](/javascript/api/word/word.tablerow#cells)|セルを取得します。|
||[font](/javascript/api/word/word.tablerow#font)|フォントを取得します。|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|行がヘッダー行であるかどうかを確認します。|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|親テーブルを取得します。|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|親テーブル内の行のインデックスを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|行のスコープで指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|行を選択し、その行に Word の UI を移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|網かけの色を取得および設定します。|
||[values](/javascript/api/word/word.tablerow#values)|行のテキスト値を 2D Javascript 配列として取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|行のセルの垂直方向の配置を取得および設定します。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|このコレクション内の最初の行を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|このコレクション内の最初の行を取得します。|
||[items](/javascript/api/word/word.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
