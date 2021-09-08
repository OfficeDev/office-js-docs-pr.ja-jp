---
title: Word JavaScript API 要件セット 1.3
description: WordApi 1.3 要件セットの詳細。
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: b58bb99e664e982d1d9047f4348755d807ad216d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936808"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 の新機能

WordApi 1.3 では、コンテンツ コントロールとドキュメント レベルの設定のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.3 の API の一覧を示します。 Word JavaScript API 要件セット 1.3 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)以前の Word API」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createDocument_base64File_)|オプションの base64 エンコードファイルを使用して新しい.docxします。|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getRange_rangeLocation_)|範囲として、本文全体、あるいは本文の開始点または終了点を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#insertTable_rowCount__columnCount__insertLocation__values_)|指定した数の行と列を含むテーブルを挿入します。|
||[lists](/javascript/api/word/word.body#lists)|本文に含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.body#parentBody)|本文の親の本文を取得します。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentBodyOrNullObject)|本文の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentContentControlOrNullObject)|本文を含むコンテンツ コントロールを取得します。|
||[parentSection](/javascript/api/word/word.body#parentSection)|本文の親セクションを取得します。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentSectionOrNullObject)|本文の親セクションを取得します。|
||[テーブル](/javascript/api/word/word.body#tables)|本文に含まれるテーブル オブジェクトのコレクションを取得します。|
||[type](/javascript/api/word/word.body#type)|本文の種類を取得します。|
||[styleBuiltIn](/javascript/api/word/word.body#styleBuiltIn)|本文の組み込みスタイル名を取得または設定します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getRange_rangeLocation_)|範囲として、コンテンツ コントロール全体、あるいはコンテンツ コントロールの開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#getTextRanges_endingMarks__trimSpacing_)|句読点や他の終了記号を使用して、コンテンツ コントロール内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#insertTable_rowCount__columnCount__insertLocation__values_)|指定した数の行と列を含むテーブルを、コンテンツ コントロール内またはコンテンツ コントロールの横に挿入します。|
||[lists](/javascript/api/word/word.contentcontrol#lists)|コンテンツ コントロールに含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentBody)|コンテンツ コントロールの親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentContentControlOrNullObject)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.contentcontrol#parentTable)|コンテンツ コントロールを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parentTableCell)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parentTableCellOrNullObject)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parentTableOrNullObject)|コンテンツ コントロールを含むテーブルを取得します。|
||[サブタイプ](/javascript/api/word/word.contentcontrol#subtype)|コンテンツ コントロールのサブタイプを取得します。|
||[テーブル](/javascript/api/word/word.contentcontrol#tables)|コンテンツ コントロールに含まれるテーブル オブジェクトのコレクションを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|区切り記号を使用して、コンテンツ コントロールを子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#styleBuiltIn)|コンテンツ コントロールの組み込みスタイル名を取得または設定します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getByIdOrNullObject_id_)|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getByTypes_types_)|指定した種類またはサブタイプを持つコンテンツ コントロールを取得します。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getFirst__)|このコレクション内の最初のコンテンツ コントロールを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getFirstOrNullObject__)|このコレクション内の最初のコンテンツ コントロールを取得します。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete__)|カスタム プロパティを削除します。|
||[key](/javascript/api/word/word.customproperty#key)|カスタム プロパティのキーを取得します。|
||[type](/javascript/api/word/word.customproperty#type)|カスタム プロパティの値の型を取得します。|
||[value](/javascript/api/word/word.customproperty#value)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add_key__value_)|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteAll__)|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getCount__)|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getItem_key_)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getItemOrNullObject_key_)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/word/word.custompropertycollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[プロパティ](/javascript/api/word/word.document#properties)|ドキュメントのプロパティを取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open__)|ドキュメントを開きます。|
||[body](/javascript/api/word/word.documentcreated#body)|ドキュメントの body オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.documentcreated#contentControls)|ドキュメント内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[プロパティ](/javascript/api/word/word.documentcreated#properties)|ドキュメントのプロパティを取得します。|
||[保存済み](/javascript/api/word/word.documentcreated#saved)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.documentcreated#sections)|ドキュメント内のセクション オブジェクトのコレクションを取得します。|
||[save()](/javascript/api/word/word.documentcreated#save__)|ドキュメントを保存します。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|ドキュメントの作成者を取得または設定します。|
||[category](/javascript/api/word/word.documentproperties#category)|ドキュメントのカテゴリを取得または設定します。|
||[comments](/javascript/api/word/word.documentproperties#comments)|ドキュメントのコメントを取得または設定します。|
||[company](/javascript/api/word/word.documentproperties#company)|ドキュメントの会社を取得または設定します。|
||[format](/javascript/api/word/word.documentproperties#format)|ドキュメントの書式設定を取得または設定します。|
||[キーワード](/javascript/api/word/word.documentproperties#keywords)|ドキュメントのキーワードを取得または設定します。|
||[上司](/javascript/api/word/word.documentproperties#manager)|ドキュメントのマネージャーを取得または設定します。|
||[applicationName](/javascript/api/word/word.documentproperties#applicationName)|ドキュメントのアプリケーション名を取得します。|
||[creationDate](/javascript/api/word/word.documentproperties#creationDate)|ドキュメントの作成日を取得します。|
||[customProperties](/javascript/api/word/word.documentproperties#customProperties)|ドキュメントのカスタム プロパティのコレクションを取得します。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastAuthor)|ドキュメントの最後の作成者を取得します。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastPrintDate)|ドキュメントを最後に印刷した日を取得します。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastSaveTime)|ドキュメントを最後に保存した時刻を取得します。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionNumber)|ドキュメントのリビジョン番号を取得します。|
||[セキュリティ](/javascript/api/word/word.documentproperties#security)|ドキュメントのセキュリティ設定を取得します。|
||[template](/javascript/api/word/word.documentproperties#template)|ドキュメントのテンプレートを取得します。|
||[subject](/javascript/api/word/word.documentproperties#subject)|ドキュメントの件名を取得または設定します。|
||[title](/javascript/api/word/word.documentproperties#title)|ドキュメントのタイトルを取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getNext__)|次のインライン画像を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getNextOrNullObject__)|次のインライン画像を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getRange_rangeLocation_)|範囲として、画像、あるいは画像の開始点または終了点を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentContentControlOrNullObject)|インライン画像を含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.inlinepicture#parentTable)|インライン イメージを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parentTableCell)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parentTableCellOrNullObject)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parentTableOrNullObject)|インライン イメージを含むテーブルを取得します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getFirst__)|このコレクション内の最初のインライン イメージを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getFirstOrNullObject__)|このコレクション内の最初のインライン イメージを取得します。|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getLevelParagraphs_level_)|リスト内の指定したレベルで発生する段落を取得します。|
||[getLevelString(level: number)](/javascript/api/word/word.list#getLevelString_level_)|指定したレベルの行頭文字、数値、または図を文字列として取得します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[id](/javascript/api/word/word.list#id)|リストの ID を取得します。|
||[levelExistences](/javascript/api/word/word.list#levelExistences)|9 つの各レベルがリストに存在するかどうかを確認します。|
||[levelTypes](/javascript/api/word/word.list#levelTypes)|リスト内の 9 レベルのすべての種類を取得します。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|リスト内の段落を取得します。|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setLevelAlignment_level__alignment_)|リスト内の指定されたレベルでの箇条書き、数字、または図の配置を設定します。|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setLevelBullet_level__listBullet__charCode__fontName_)|リスト内の指定したレベルで行頭文字の書式を設定します。|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setLevelIndents_level__textIndent__bulletNumberPictureIndent_)|リスト内の指定したレベルの 2 つのインデントを設定します。|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: 配列<文字列 \|>)](/javascript/api/word/word.list#setLevelNumbering_level__listNumbering__formatString_)|リスト内の指定したレベルで番号付け書式を設定します。|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setLevelStartingNumber_level__startingNumber_)|リスト内の指定したレベルで開始番号を設定します。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getById_id_)|識別子を使用してリストを取得します。|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getByIdOrNullObject_id_)|識別子を使用してリストを取得します。|
||[getFirst()](/javascript/api/word/word.listcollection#getFirst__)|このコレクション内の最初のリストを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getFirstOrNullObject__)|このコレクション内の最初のリストを取得します。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getItem_index_)|コレクション内のインデックスを使用して、リスト オブジェクトを取得します。|
||[items](/javascript/api/word/word.listcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestor_parentOnly_)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestorOrNullObject_parentOnly_)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getDescendants_directChildrenOnly_)|リスト アイテムのすべての子孫のリスト アイテムを取得します。|
||[level](/javascript/api/word/word.listitem#level)|リスト内のアイテムのレベルを取得または設定します。|
||[listString](/javascript/api/word/word.listitem#listString)|リスト アイテムの箇条書き、数値、または図を文字列として取得します。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingIndex)|兄弟を基準にしてリスト アイテムの注文番号を取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachToList_listId__level_)|指定したレベルで段落を既存のリストに結合させます。|
||[detachFromList()](/javascript/api/word/word.paragraph#detachFromList__)|段落がリスト アイテムである場合は、この段落をリストから移動します。|
||[getNext()](/javascript/api/word/word.paragraph#getNext__)|次の段落を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getNextOrNullObject__)|次の段落を取得します。|
||[getPrevious()](/javascript/api/word/word.paragraph#getPrevious__)|前の段落を取得します。|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getPreviousOrNullObject__)|前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getRange_rangeLocation_)|段落全体、あるいは段落の開始点または終了点を範囲として取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#getTextRanges_endingMarks__trimSpacing_)|句読点や他の終了記号を使用して、段落内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#insertTable_rowCount__columnCount__insertLocation__values_)|指定した数の行と列を含むテーブルを挿入します。|
||[isLastParagraph](/javascript/api/word/word.paragraph#isLastParagraph)|段落がその親の本文内の最後の段落であることを示します。|
||[isListItem](/javascript/api/word/word.paragraph#isListItem)|段落がリスト アイテムであるかどうかを確認します。|
||[list](/javascript/api/word/word.paragraph#list)|この段落が属するリストを取得します。|
||[listItem](/javascript/api/word/word.paragraph#listItem)|段落の ListItem を取得します。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listItemOrNullObject)|段落の ListItem を取得します。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listOrNullObject)|この段落が属するリストを取得します。|
||[parentBody](/javascript/api/word/word.paragraph#parentBody)|段落の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentContentControlOrNullObject)|段落を格納しているコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.paragraph#parentTable)|段落を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.paragraph#parentTableCell)|段落を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parentTableCellOrNullObject)|段落を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parentTableOrNullObject)|段落を含むテーブルを取得します。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tableNestingLevel)|段落のテーブルのレベルを取得します。|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split_delimiters__trimDelimiters__trimSpacing_)|区切り記号を使用して、段落を子の範囲に分割します。|
||[startNewList()](/javascript/api/word/word.paragraph#startNewList__)|この段落を含む新しいリストを開始します。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#styleBuiltIn)|段落の組み込みスタイル名を取得または設定します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getFirst__)|このコレクション内の最初の段落を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getFirstOrNullObject__)|このコレクション内の最初の段落を取得します。|
||[getLast()](/javascript/api/word/word.paragraphcollection#getLast__)|このコレクション内の最後の段落を取得します。|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getLastOrNullObject__)|このコレクション内の最後の段落を取得します。|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#compareLocationWith_range_)|この範囲の場所を別の範囲の場所と比較します。|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandTo_range_)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandToOrNullObject_range_)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[getHyperlinkRanges()](/javascript/api/word/word.range#getHyperlinkRanges__)|範囲内のハイパーリンクの子の範囲を取得します。|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRange_endingMarks__trimSpacing_)|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRangeOrNullObject_endingMarks__trimSpacing_)|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getRange_rangeLocation_)|範囲の複製を作成するか、新しい範囲として開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getTextRanges_endingMarks__trimSpacing_)|句読点や他の終了記号を使用して、範囲内のテキストの子範囲を取得します。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#insertTable_rowCount__columnCount__insertLocation__values_)|指定した数の行と列を含むテーブルを挿入します。|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectWith_range_)|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectWithOrNullObject_range_)|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[isEmpty](/javascript/api/word/word.range#isEmpty)|範囲の長さが 0 であるかどうかを確認します。|
||[lists](/javascript/api/word/word.range#lists)|範囲内のリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.range#parentBody)|範囲の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentContentControlOrNullObject)|範囲を格納するコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.range#parentTable)|範囲を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.range#parentTableCell)|範囲を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parentTableCellOrNullObject)|範囲を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parentTableOrNullObject)|範囲を含むテーブルを取得します。|
||[テーブル](/javascript/api/word/word.range#tables)|範囲内のテーブル オブジェクトのコレクションを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|区切り記号を使用して、範囲を子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.range#styleBuiltIn)|範囲の組み込みスタイル名を取得または設定します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getFirst__)|このコレクション内の最初の範囲を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getFirstOrNullObject__)|このコレクション内の最初の範囲を取得します。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api セット: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getNext__)|次のセクションを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.section#getNextOrNullObject__)|次のセクションを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getFirst__)|このコレクション内の最初のセクションを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getFirstOrNullObject__)|このコレクション内の最初のセクションを取得します。|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addColumns_insertLocation__columnCount__values_)|最初または最後の既存の列をテンプレートとして使用して、テーブルの最初または最後に列を追加します。|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addRows_insertLocation__rowCount__values_)|最初または最後の既存の行をテンプレートとして使用して、テーブルの最初または最後に行を追加します。|
||[配置](/javascript/api/word/word.table#alignment)|ページ列に対するテーブルの配置を取得または設定します。|
||[autoFitWindow()](/javascript/api/word/word.table#autoFitWindow__)|テーブルの列をウィンドウの幅に合わせて自動調整します。|
||[clear()](/javascript/api/word/word.table#clear__)|テーブルの内容をクリアします。|
||[delete()](/javascript/api/word/word.table#delete__)|テーブル全体を削除します。|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deleteColumns_columnIndex__columnCount_)|特定の列を削除します。|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleteRows_rowIndex__rowCount_)|特定の行を削除します。|
||[distributeColumns()](/javascript/api/word/word.table#distributeColumns__)|列の幅を揃えます。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getBorder_borderLocation_)|指定した罫線の罫線スタイルを取得します。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCell_rowIndex__cellIndex_)|指定された行と列のテーブル セルを取得します。|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCellOrNullObject_rowIndex__cellIndex_)|指定された行と列のテーブル セルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getCellPadding_cellPaddingLocation_)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.table#getNext__)|次のテーブルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.table#getNextOrNullObject__)|次のテーブルを取得します。|
||[getParagraphAfter()](/javascript/api/word/word.table#getParagraphAfter__)|テーブルの後の段落を取得します。|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getParagraphAfterOrNullObject__)|テーブルの後の段落を取得します。|
||[getParagraphBefore()](/javascript/api/word/word.table#getParagraphBefore__)|テーブルの前の段落を取得します。|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getParagraphBeforeOrNullObject__)|テーブルの前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getRange_rangeLocation_)|このテーブルを含む範囲、あるいはテーブルの開始または終了の範囲を取得します。|
||[headerRowCount](/javascript/api/word/word.table#headerRowCount)|ヘッダー行の数を取得および設定します。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalAlignment)|テーブル内のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.table#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignoreSpace)||
||[insertContentControl()](/javascript/api/word/word.table#insertContentControl__)|テーブルにコンテンツ コントロールを挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertParagraph_paragraphText__insertLocation_)|指定した位置に、段落を挿入します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#insertTable_rowCount__columnCount__insertLocation__values_)|指定した数の行と列を含むテーブルを挿入します。|
||[matchCase](/javascript/api/word/word.table#matchCase)||
||[matchPrefix](/javascript/api/word/word.table#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.table#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.table#matchWildcards)||
||[font](/javascript/api/word/word.table#font)|フォントを取得します。|
||[isUniform](/javascript/api/word/word.table#isUniform)|すべてのテーブル行が均一かどうかを示します。|
||[nestingLevel](/javascript/api/word/word.table#nestingLevel)|テーブルの入れ子のレベルを取得します。|
||[parentBody](/javascript/api/word/word.table#parentBody)|テーブルの親の本文を取得します。|
||[parentContentControl](/javascript/api/word/word.table#parentContentControl)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentContentControlOrNullObject)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.table#parentTable)|このテーブルを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.table#parentTableCell)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parentTableCellOrNullObject)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parentTableOrNullObject)|このテーブルを含むテーブルを取得します。|
||[rowCount](/javascript/api/word/word.table#rowCount)|表の行数を取得します。|
||[rows](/javascript/api/word/word.table#rows)|すべてのテーブルの行を取得します。|
||[テーブル](/javascript/api/word/word.table#tables)|1 レベル深く入れ子にされた子テーブルを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|テーブル オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select_selectionMode_)|テーブル、あるいはテーブルの開始位置または終了位置を選択して、Word の UI に移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setCellPadding_cellPaddingLocation__cellPadding_)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.table#shadingColor)|網かけの色を取得および設定します。|
||[style](/javascript/api/word/word.table#style)|テーブルのスタイル名を取得または設定します。|
||[styleBandedColumns](/javascript/api/word/word.table#styleBandedColumns)|テーブルの列を縞模様にするかどうかを取得および設定します。|
||[styleBandedRows](/javascript/api/word/word.table#styleBandedRows)|テーブルの行を縞模様にするかどうかを取得および設定します。|
||[styleBuiltIn](/javascript/api/word/word.table#styleBuiltIn)|テーブルの組み込みスタイル名を取得または設定します。|
||[styleFirstColumn](/javascript/api/word/word.table#styleFirstColumn)|テーブルの最初の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleLastColumn](/javascript/api/word/word.table#styleLastColumn)|テーブルの最後の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleTotalRow](/javascript/api/word/word.table#styleTotalRow)|テーブルの集計 (最後) 行に特別なスタイルを指定するかどうかを取得および設定します。|
||[values](/javascript/api/word/word.table#values)|2D の Javascript 配列として、テーブルのテキスト値を取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.table#verticalAlignment)|テーブル内のすべてのセルの垂直方向の配置を取得および設定します。|
||[width](/javascript/api/word/word.table#width)|テーブルの幅をポイント単位で取得および設定します。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|テーブルの罫線の色を取得または設定します。|
||[type](/javascript/api/word/word.tableborder#type)|テーブルの罫線の種類を取得または設定します。|
||[width](/javascript/api/word/word.tableborder#width)|テーブルの罫線の幅をポイント単位で得または設定します。|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnWidth)|セルの列の幅をポイント単位で取得または設定します。|
||[deleteColumn()](/javascript/api/word/word.tablecell#deleteColumn__)|このセルを含む列を削除します。|
||[deleteRow()](/javascript/api/word/word.tablecell#deleteRow__)|このセルを含む行を削除します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getBorder_borderLocation_)|指定した罫線の罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getCellPadding_cellPaddingLocation_)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablecell#getNext__)|次のセルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getNextOrNullObject__)|次のセルを取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalAlignment)|セルの水平方向の配置を取得および設定します。|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertColumns_insertLocation__columnCount__values_)|セルの列をテンプレートとして使用して、列をセルの左または右に追加します。|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertRows_insertLocation__rowCount__values_)|セルの行をテンプレートとして使用して、行をセルの上または下に挿入します。|
||[body](/javascript/api/word/word.tablecell#body)|セルの本文オブジェクトを取得します。|
||[cellIndex](/javascript/api/word/word.tablecell#cellIndex)|その行のセルのインデックスを取得します。|
||[parentRow](/javascript/api/word/word.tablecell#parentRow)|セルの親行を取得します。|
||[parentTable](/javascript/api/word/word.tablecell#parentTable)|セルの親テーブルを取得します。|
||[rowIndex](/javascript/api/word/word.tablecell#rowIndex)|テーブルのセル行のインデックスを取得します。|
||[width](/javascript/api/word/word.tablecell#width)|セルの幅をポイント単位で取得します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setCellPadding_cellPaddingLocation__cellPadding_)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingColor)|セルの網かけの色を取得または設定します。|
||[value](/javascript/api/word/word.tablecell#value)|セルのテキストを取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalAlignment)|セルの垂直方向の配置を取得および設定します。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getFirst__)|このコレクション内の最初のテーブル セルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getFirstOrNullObject__)|このコレクション内の最初のテーブル セルを取得します。|
||[items](/javascript/api/word/word.tablecellcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getFirst__)|このコレクション内の最初のテーブルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getFirstOrNullObject__)|このコレクション内の最初のテーブルを取得します。|
||[items](/javascript/api/word/word.tablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear__)|行の内容をクリアします。|
||[delete()](/javascript/api/word/word.tablerow#delete__)|行全体を削除します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getBorder_borderLocation_)|行のセルの罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getCellPadding_cellPaddingLocation_)|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablerow#getNext__)|次の行を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getNextOrNullObject__)|次の行を取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalAlignment)|行のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignoreSpace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertRows_insertLocation__rowCount__values_)|この行をテンプレートとして使用して、行を挿入します。|
||[matchCase](/javascript/api/word/word.tablerow#matchCase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchWildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredHeight)|適切な行の高さをポイント単位で取得および設定します。|
||[cellCount](/javascript/api/word/word.tablerow#cellCount)|行のセルの数を取得します。|
||[セル](/javascript/api/word/word.tablerow#cells)|セルを取得します。|
||[font](/javascript/api/word/word.tablerow#font)|フォントを取得します。|
||[isHeader](/javascript/api/word/word.tablerow#isHeader)|行がヘッダー行であるかどうかを確認します。|
||[parentTable](/javascript/api/word/word.tablerow#parentTable)|親テーブルを取得します。|
||[rowIndex](/javascript/api/word/word.tablerow#rowIndex)|親テーブル内の行のインデックスを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|行のスコープで指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select_selectionMode_)|行を選択し、その行に Word の UI を移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setCellPadding_cellPaddingLocation__cellPadding_)|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingColor)|網かけの色を取得および設定します。|
||[values](/javascript/api/word/word.tablerow#values)|行のテキスト値を 2D Javascript 配列として取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalAlignment)|行のセルの垂直方向の配置を取得および設定します。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getFirst__)|このコレクション内の最初の行を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getFirstOrNullObject__)|このコレクション内の最初の行を取得します。|
||[items](/javascript/api/word/word.tablerowcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
