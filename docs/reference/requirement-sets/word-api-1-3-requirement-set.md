---
title: Word JavaScript API 要件セット 1.3
description: WordApi 1.3 要件セットの詳細。
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: d9e0d450b601845d4e11e0fd74652c4e167f802c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746030"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 の新機能

WordApi 1.3 では、コンテンツ コントロールとドキュメント レベルの設定のサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Word JavaScript API 要件セット 1.3 の API の一覧を示します。 Word JavaScript API 要件セット 1.3 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、「要件セット [1.3 以前の Word API」を参照してください](/javascript/api/word?view=word-js-1.3&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|オプションの base64 エンコードファイルを使用して新しい.docxします。|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|範囲として、本文全体、あるいは本文の開始点または終了点を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|指定した数の行と列を含むテーブルを挿入します。|
||[サイト](/javascript/api/word/word.body#word-word-body-lists-member)|本文に含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|本文の親の本文を取得します。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|本文の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|本文を含むコンテンツ コントロールを取得します。|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|本文の親セクションを取得します。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|本文の親セクションを取得します。|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|本文の組み込みスタイル名を取得または設定します。|
||[テーブル](/javascript/api/word/word.body#word-word-body-tables-member)|本文に含まれるテーブル オブジェクトのコレクションを取得します。|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|本文の種類を取得します。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|範囲として、コンテンツ コントロール全体、あるいはコンテンツ コントロールの開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|句読点や他の終了記号を使用して、コンテンツ コントロール内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|指定した数の行と列を含むテーブルを、コンテンツ コントロール内またはコンテンツ コントロールの横に挿入します。|
||[サイト](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|コンテンツ コントロールに含まれるリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|コンテンツ コントロールの親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|コンテンツ コントロールを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|コンテンツ コントロールを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|コンテンツ コントロールを含むテーブル セルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|コンテンツ コントロールを含むテーブルを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|区切り記号を使用して、コンテンツ コントロールを子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|コンテンツ コントロールの組み込みスタイル名を取得または設定します。|
||[サブタイプ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|コンテンツ コントロールのサブタイプを取得します。|
||[テーブル](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|コンテンツ コントロールに含まれるテーブル オブジェクトのコレクションを取得します。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|コンテンツ コントロールの識別子によってコンテンツ コントロールを取得します。|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|指定した種類またはサブタイプを持つコンテンツ コントロールを取得します。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|このコレクション内の最初のコンテンツ コントロールを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|このコレクション内の最初のコンテンツ コントロールを取得します。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|カスタム プロパティを削除します。|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|カスタム プロパティのキーを取得します。|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|カスタム プロパティの値の型を取得します。|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|カスタム プロパティの値を取得または設定します。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|このコレクション内のすべてのカスタム プロパティを削除します。|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|カスタム プロパティの数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ドキュメント](/javascript/api/word/word.document)|[プロパティ](/javascript/api/word/word.document#word-word-document-properties-member)|ドキュメントのプロパティを取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|ドキュメントの body オブジェクトを取得します。|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|ドキュメント内のコンテンツ コントロール オブジェクトのコレクションを取得します。|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|ドキュメントを開きます。|
||[プロパティ](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|ドキュメントのプロパティを取得します。|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|ドキュメントを保存します。|
||[保存済み](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|ドキュメント内の変更が保存されているかどうかを示します。|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|ドキュメント内のセクション オブジェクトのコレクションを取得します。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|ドキュメントのアプリケーション名を取得します。|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|ドキュメントの作成者を取得または設定します。|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|ドキュメントのカテゴリを取得または設定します。|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|ドキュメントのコメントを取得または設定します。|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|ドキュメントの会社を取得または設定します。|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|ドキュメントの作成日を取得します。|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|ドキュメントのカスタム プロパティのコレクションを取得します。|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|ドキュメントの書式設定を取得または設定します。|
||[キーワード](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|ドキュメントのキーワードを取得または設定します。|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|ドキュメントの最後の作成者を取得します。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|ドキュメントを最後に印刷した日を取得します。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|ドキュメントを最後に保存した時刻を取得します。|
||[上司](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|ドキュメントのマネージャーを取得または設定します。|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|ドキュメントのリビジョン番号を取得します。|
||[セキュリティ](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|ドキュメントのセキュリティ設定を取得します。|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|ドキュメントの件名を取得または設定します。|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|ドキュメントのテンプレートを取得します。|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|ドキュメントのタイトルを取得または設定します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|次のインライン画像を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|次のインライン画像を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|範囲として、画像、あるいは画像の開始点または終了点を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|インライン画像を含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|インライン イメージを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|インライン イメージを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|インライン イメージを含むテーブルを取得します。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|このコレクション内の最初のインライン イメージを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|このコレクション内の最初のインライン イメージを取得します。|
|[リスト](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|リスト内の指定したレベルで発生する段落を取得します。|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|指定したレベルの行頭文字、数値、または図を文字列として取得します。|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|リストの ID を取得します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|9 つの各レベルがリストに存在するかどうかを確認します。|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|リスト内の 9 レベルのすべての種類を取得します。|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|リスト内の段落を取得します。|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|リスト内の指定されたレベルでの箇条書き、数字、または図の配置を設定します。|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|リスト内の指定したレベルで行頭文字の書式を設定します。|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|リスト内の指定したレベルの 2 つのインデントを設定します。|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: 配列<文字列 \|>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|リスト内の指定したレベルで番号付け書式を設定します。|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|リスト内の指定したレベルで開始番号を設定します。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|識別子を使用してリストを取得します。|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|識別子を使用してリストを取得します。|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|このコレクション内の最初のリストを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|このコレクション内の最初のリストを取得します。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|コレクション内のインデックスを使用して、リスト オブジェクトを取得します。|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|リスト アイテムのすべての子孫のリスト アイテムを取得します。|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|リスト内のアイテムのレベルを取得または設定します。|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|リスト アイテムの箇条書き、数値、または図を文字列として取得します。|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|兄弟を基準にしてリスト アイテムの注文番号を取得します。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|指定したレベルで段落を既存のリストに結合させます。|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|段落がリスト アイテムである場合は、この段落をリストから移動します。|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|次の段落を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|次の段落を取得します。|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|前の段落を取得します。|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|段落全体、あるいは段落の開始点または終了点を範囲として取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|句読点や他の終了記号を使用して、段落内のテキスト範囲を取得します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|指定した数の行と列を含むテーブルを挿入します。|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|段落がその親の本文内の最後の段落であることを示します。|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|段落がリスト アイテムであるかどうかを確認します。|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|この段落が属するリストを取得します。|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|段落の ListItem を取得します。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|段落の ListItem を取得します。|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|この段落が属するリストを取得します。|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|段落の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|段落を格納しているコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|段落を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|段落を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|段落を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|段落を含むテーブルを取得します。|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|区切り記号を使用して、段落を子の範囲に分割します。|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|この段落を含む新しいリストを開始します。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|段落の組み込みスタイル名を取得または設定します。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|段落のテーブルのレベルを取得します。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|このコレクション内の最初の段落を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|このコレクション内の最初の段落を取得します。|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|このコレクション内の最後の段落を取得します。|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|このコレクション内の最後の段落を取得します。|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|この範囲の場所を別の範囲の場所と比較します。|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|範囲内のハイパーリンクの子の範囲を取得します。|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|句読点や他の終了記号を使用して、次のテキスト範囲を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|範囲の複製を作成するか、新しい範囲として開始点または終了点を取得します。|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|句読点や他の終了記号を使用して、範囲内のテキストの子範囲を取得します。|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|指定した数の行と列を含むテーブルを挿入します。|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|別の範囲とこの範囲の交点として、新しい範囲を返します。|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|範囲の長さが 0 であるかどうかを確認します。|
||[サイト](/javascript/api/word/word.range#word-word-range-lists-member)|範囲内のリスト オブジェクトのコレクションを取得します。|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|範囲の親の本文を取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|範囲を格納するコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|範囲を含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|範囲を含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|範囲を含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|範囲を含むテーブルを取得します。|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|区切り記号を使用して、範囲を子の範囲に分割します。|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|範囲の組み込みスタイル名を取得または設定します。|
||[テーブル](/javascript/api/word/word.range#word-word-range-tables-member)|範囲内のテーブル オブジェクトのコレクションを取得します。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|このコレクション内の最初の範囲を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|このコレクション内の最初の範囲を取得します。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Api セット: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|次のセクションを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|次のセクションを取得します。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|このコレクション内の最初のセクションを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|このコレクション内の最初のセクションを取得します。|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|最初または最後の既存の列をテンプレートとして使用して、テーブルの最初または最後に列を追加します。|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|最初または最後の既存の行をテンプレートとして使用して、テーブルの最初または最後に行を追加します。|
||[配置](/javascript/api/word/word.table#word-word-table-alignment-member)|ページ列に対するテーブルの配置を取得または設定します。|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|テーブルの列をウィンドウの幅に合わせて自動調整します。|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|テーブルの内容をクリアします。|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|テーブル全体を削除します。|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|特定の列を削除します。|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|特定の行を削除します。|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|列の幅を揃えます。|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|フォントを取得します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|指定した罫線の罫線スタイルを取得します。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|指定された行と列のテーブル セルを取得します。|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|指定された行と列のテーブル セルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|次のテーブルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|次のテーブルを取得します。|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|テーブルの後の段落を取得します。|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|テーブルの後の段落を取得します。|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|テーブルの前の段落を取得します。|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|テーブルの前の段落を取得します。|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|このテーブルを含む範囲、あるいはテーブルの開始または終了の範囲を取得します。|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|ヘッダー行の数を取得および設定します。|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|テーブル内のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|テーブルにコンテンツ コントロールを挿入します。|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|指定した位置に、段落を挿入します。|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|指定した数の行と列を含むテーブルを挿入します。|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|すべてのテーブル行が均一かどうかを示します。|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|テーブルの入れ子のレベルを取得します。|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|テーブルの親の本文を取得します。|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|テーブルを含むコンテンツ コントロールを取得します。|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|このテーブルを含むテーブルを取得します。|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|このテーブルを含むテーブルのセルを取得します。|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|このテーブルを含むテーブルを取得します。|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|表の行数を取得します。|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|すべてのテーブルの行を取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.table#word-word-table-search-member(1))|テーブル オブジェクトのスコープで、指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|テーブル、あるいはテーブルの開始位置または終了位置を選択して、Word の UI に移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|網かけの色を取得および設定します。|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|テーブルのスタイル名を取得または設定します。|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|テーブルの列を縞模様にするかどうかを取得および設定します。|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|テーブルの行を縞模様にするかどうかを取得および設定します。|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|テーブルの組み込みスタイル名を取得または設定します。|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|テーブルの最初の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|テーブルの最後の列に特別なスタイルを指定するかどうかを取得および設定します。|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|テーブルの集計 (最後) 行に特別なスタイルを指定するかどうかを取得および設定します。|
||[テーブル](/javascript/api/word/word.table#word-word-table-tables-member)|1 レベル深く入れ子にされた子テーブルを取得します。|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|2D の Javascript 配列として、テーブルのテキスト値を取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|テーブル内のすべてのセルの垂直方向の配置を取得および設定します。|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|テーブルの幅をポイント単位で取得および設定します。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|テーブルの罫線の色を取得または設定します。|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|テーブルの罫線の種類を取得または設定します。|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|テーブルの罫線の幅をポイント単位で得または設定します。|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|セルの本文オブジェクトを取得します。|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|その行のセルのインデックスを取得します。|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|セルの列の幅をポイント単位で取得または設定します。|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|このセルを含む列を削除します。|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|このセルを含む行を削除します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|指定した罫線の罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|次のセルを取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|次のセルを取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|セルの水平方向の配置を取得および設定します。|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|セルの列をテンプレートとして使用して、列をセルの左または右に追加します。|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|セルの行をテンプレートとして使用して、行をセルの上または下に挿入します。|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|セルの親行を取得します。|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|セルの親テーブルを取得します。|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|テーブルのセル行のインデックスを取得します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|セルの網かけの色を取得または設定します。|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|セルのテキストを取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|セルの垂直方向の配置を取得および設定します。|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|セルの幅をポイント単位で取得します。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|このコレクション内の最初のテーブル セルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|このコレクション内の最初のテーブル セルを取得します。|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|このコレクション内の最初のテーブルを取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|このコレクション内の最初のテーブルを取得します。|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|行のセルの数を取得します。|
||[セル](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|セルを取得します。|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|行の内容をクリアします。|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|行全体を削除します。|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|フォントを取得します。|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|行のセルの罫線スタイルを取得します。|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|セル内のスペースをポイント単位で取得します。|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|次の行を取得します。|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|次の行を取得します。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|行のすべてのセルの水平方向の配置を取得および設定します。|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|この行をテンプレートとして使用して、行を挿入します。|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|行がヘッダー行であるかどうかを確認します。|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|親テーブルを取得します。|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|適切な行の高さをポイント単位で取得および設定します。|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|親テーブル内の行のインデックスを取得します。|
||[search(searchText: string, searchOptions?: Word.SearchOptions { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean matchWildcards \| ?: boolean })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|行のスコープで指定した SearchOptions を使用して検索を実行します。|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|行を選択し、その行に Word の UI を移動します。|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|セル内のスペースをポイント単位で設定します。|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|網かけの色を取得および設定します。|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|行のテキスト値を 2D Javascript 配列として取得および設定します。|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|行のセルの垂直方向の配置を取得および設定します。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|このコレクション内の最初の行を取得します。|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|このコレクション内の最初の行を取得します。|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Word JavaScript API の要件セット](word-api-requirement-sets.md)
