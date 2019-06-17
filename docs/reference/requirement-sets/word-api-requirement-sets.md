---
title: Word JavaScript API の要件セット
description: ''
ms.date: 06/11/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: be2c9834fbf3ceabcbbca6f2378b4356095ab387
ms.sourcegitcommit: e112a9b29376b1f574ee13b01c818131b2c7889d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2019
ms.locfileid: "34997394"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Word アドインは、Windows での Office 2016 以降、Office for iPad、Office for Mac、Office Online など、複数のバージョンの Office で機能します。 次の表は、Word の要件セット、その要件セットをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド番号またはバージョン番号の一覧です。

> [!NOTE]
> ベータ版としてマークされている要件セットについては、指定されたバージョン以降の Office ソフトウェアを使用し、CDN のベータ版のライブラリを使用します: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
>
> ベータ版として表示されていないエントリは一般公開されており、引き続き Production CDN ライブラリを使用できます: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  要件セット  |   Windows での Office\*<br>(Office 365 に接続)  |  Office for iPad<br>(Office 365 に接続)  |  Office for Mac<br>(Office 365 に接続)  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| [プレビュー](/javascript/api/word)  | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要があります) |
| WordApi 1.3 | バージョン 1612 (ビルド 7668.1000) 以降| 2017 年 3 月、2.22 以降 | 2017 年 3 月、15.32 以降| 2017 年 3 月 ||
| WordApi 1.2  | 2015年 12 月更新プログラム、バージョン 1601 (ビルド 6568.1000) 以降 | 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 | |
| WordApi 1.1  | バージョン 1509 (ビルド 4266.1001) 以降| 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 | |

> [!NOTE]
> MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。 このバージョンには、WordApi 1.1 の要件セットのみが含まれています。

バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="word-javascript-preview-apis"></a>Word JavaScript プレビュー API

新しい Word JavaScript API は最初に「プレビュー」で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには、CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)) で**ベータ** ライブラリを参照する必要があります。場合によっては、Office Insider プログラムに参加し、新しい Office ビルドを入手する必要があります。

以下は、プレビュー中の API の完全な一覧です。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|コンテンツ コントロール内のデータが変更された場合に発生します。 新しいテキストを取得するには、このコンテンツ コントロールをハンドラーに読み込みます。 古いテキストを取得するには、読み込まないでください。|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|コンテンツ コントロールが変更された場合に発生します。 このコンテンツ コントロールはハンドラーに読み込まないでください。これ以外の場合、元のプロパティを取得することはできません。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|コンテンツ コントロール内の選択が変更された場合に発生します。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|イベントを発生させたオブジェクト。 このオブジェクトを読み込み、プロパティを取得します。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|イベントの種類。 詳細については、「Word.EventType」を参照してください。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|カスタム XML パーツを削除します。|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|xpath で識別された要素から、指定された名前を持つ属性を削除します。|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|xpath で識別された要素を削除します。|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|カスタム XML パーツのすべての XML コンテンツを取得します。|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|xpath で識別された要素に、指定された名前および値を持つ属性を挿入します。|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|指定された XML を、子ポジション インデックスの xpath で識別された親要素の下に、挿入します。|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|カスタム XML パーツの XML コンテンツをクエリします。|
||[id](/javascript/api/word/word.customxmlpart#id)|カスタム XML パーツの ID を取得します。 読み取り専用です。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|カスタム XML パーツの名前空間 URI を取得します。 読み取り専用です。|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|カスタム XML パーツのすべての XML コンテンツを設定します。|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|xpath で識別された要素の指定された名前を持つ属性の値を更新します。|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|xpath で識別された要素の XML を更新します。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|ドキュメントに新しいカスタム XML パーツを追加します。|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。 読み取り専用です。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。 CustomXmlPart が存在しない場合は、null オブジェクトを返します。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|ID に基づいて、カスタム XML パーツを取得します。 読み取り専用です。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|ID に基づいて、カスタム XML パーツを取得します。 コレクション内に CustomXmlPart が存在しない場合は、null オブジェクトを返します。|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。 これ以外の場合、このメソッドはエラーになります。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。 これ以外の場合、このメソッドは null オブジェクトを返します。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deletebookmark-name-)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|ブックマークの範囲を取得します。 ブックマークが存在しない場合は、スローします。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|ブックマークの範囲を取得します。 ブックマークが存在しない場合は、null オブジェクトを返します。|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|ドキュメントのカスタム XML パーツを取得します。 読み取り専用です。|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|コンテンツ コントロールが追加された場合に発生します。 ハンドラーで context.sync() を実行して、新しいコンテンツ コントロールのプロパティを取得します。|
||[settings](/javascript/api/word/word.document#settings)|ドキュメント内のアドインの設定を取得します。 読み取り専用です。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|ブックマークの範囲を取得します。 ブックマークが存在しない場合は、スローします。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|ブックマークの範囲を取得します。 ブックマークが存在しない場合は、null オブジェクトを返します。|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|ドキュメントのカスタム XML パーツを取得します。 読み取り専用です。|
||[settings](/javascript/api/word/word.documentcreated#settings)|ドキュメント内のアドインの設定を取得します。 読み取り専用です。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|インライン画像の形式を取得します。 読み取り専用です。|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getlevelfont-level-)|リスト内の指定したレベルで行頭文字のフォント、番号、画像のいずれかを取得します。|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|リスト内の指定したレベルで画像の base64 エンコード文字列表記を取得します。|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|リスト内の指定したレベルで行頭文字のフォント、番号、画像のいずれかを再設定します。|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|リスト内の指定したレベルで画像を設定します。|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|範囲内または重なる範囲のすべてのブックマークの名前を取得します。 名前がアンダースコア文字で始まるブックマークは、非表示になります。|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertbookmark-name-)|範囲にブックマークを挿入します。 同じ名前のブックマークがどこかに存在する場合は、最初に削除されます。|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|設定を削除します。|
||[key](/javascript/api/word/word.setting#key)|設定のキーを取得します。 読み取り専用です。|
||[value](/javascript/api/word/word.setting#value)|設定の値を取得または設定します。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|新しい設定を作成するか、既存の設定を設定します。|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteall--)|このアドインのすべての設定を削除します。|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|設定の数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|setting オブジェクトをその大文字と小文字が区別されるキーによって取得します。 setting が存在しない場合にスローされます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|setting オブジェクトをその大文字と小文字が区別されるキーによって取得します。 setting が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/word/word.settingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|最初と最後のセルによって包括的に囲まれたセルを結合します。|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|セルを指定した数の行と列に分割します。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|行にコンテンツ コントロールを挿入します。|
||[merge()](/javascript/api/word/word.tablerow#merge--)|行を 1 つのセルに結合します。|

## <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 の新機能

要件セット 1.3 の Word JavaScript API に新たに追加された機能は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:-----|-----|:----|:----|
|[アプリケーション](/javascript/api/word/word.application)|_メソッド_ > createDocument(base64File: string) | Base64 でエンコードされた .docx ファイルを使用して、新しい文書を作成します。 読み取り専用です。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > lists|本文に含まれるリスト オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > parentBody|本文の親の本文を取得します。たとえば、テーブル セル本文の親本文にはヘッダーを指定できます。読み取り専用。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > parentSection|本文の親セクションを取得します。読み取り専用。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > styleBuiltIn|本文の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > tables|本文に含まれるテーブル オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[body](/javascript/api/word/word.body)|_リレーションシップ_ > type|本文の種類を取得します。種類は、'MainDoc'、'Section'、'Header'、'Footer'、または 'TableCell' にできます。読み取り専用。|1.3|
|[本文](/javascript/api/word/word.body)|_メソッド_ > getRange(rangeLocation: RangeLocation)|範囲として、本文全体、あるいは本文の開始点または終了点を取得します。|1.3|
|[本文](/javascript/api/word/word.body)|_メソッド_ > insertTable(rowCount: number, columnCount: number, insertLocation:InsertLocation, values: string)|指定した数の行と列を含むテーブルを挿入します。insertLocation の値には、'Start' または 'End' を指定できます。|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_リレーションシップ_ > breaks|改行の形式を指定します。行、ページ、またはセクション タイプです。 読み取り専用です。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > lists|コンテンツ コントロールに含まれるリスト オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > parentBody|コンテンツ コントロールの親の本文を取得します。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > parentTable|コンテンツ コントロールを含むテーブルを取得します。テーブルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > parentTableCell|コンテンツ コントロールを含むテーブル セルを取得します。テーブル セルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > styleBuiltIn|コンテンツ コントロールの組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > subtype|コンテンツ コントロールのサブタイプを取得します。リッチ テキスト コンテンツ コントロールの場合、サブタイプは、'RichTextInline'、'RichTextParagraphs'、'RichTextTableCell'、'RichTextTableRow' および 'RichTextTable' にできます。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_リレーションシップ_ > tables|コンテンツ コントロールに含まれるテーブル オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_メソッド_ > getRange(rangeLocation: RangeLocation)|範囲として、コンテンツ コントロール全体、あるいはコンテンツ コントロールの開始点または終了点を取得します。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_メソッド_ > getTextRanges(endingMarks: string, trimSpacing: bool)|句読点と他の終了記号、またはそのいずれかを使用して、コンテンツ コントロール内のテキスト範囲を取得します。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_メソッド_ > insertTable(rowCount: number, columnCount: number, insertLocation:InsertLocation, values: string)|指定した数の行と列を含むテーブルを、コンテンツ コントロール内またはコンテンツ コントロールの横に挿入します。insertLocation の値には、'Start'、'End'、'Before' または 'After' を指定できます。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_メソッド_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|区切り記号を使用して、コンテンツ コントロールを子の範囲に分割します。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_メソッド_ > getByTypes(types: ContentControlType)|指定した種類とサブタイプ、またはそのいずれかを含むコンテンツ コントロールを取得します。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_メソッド_ > getFirst()|このコレクション内の最初のコンテンツ コントロールを取得します。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_プロパティ_ > key|カスタム プロパティのキーを取得します。 読み取り専用です。 |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_プロパティ_ > value|カスタム プロパティの値を取得または設定します。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_リレーションシップ_ > type|カスタム プロパティの値の型を取得します。 読み取り専用です。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_メソッド_ > delete()|カスタム プロパティを削除します。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_プロパティ_ > items|customProperty オブジェクトのコレクション。読み取り専用。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_メソッド_ > deleteAll()|このコレクション内のすべてのカスタム プロパティを削除します。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_メソッド_ > getCount()|カスタム プロパティの数を取得します。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_メソッド_ > getItem(key: string)|キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_メソッド_ > set(key: string, value: object)|カスタム プロパティを作成または設定します。|1.3|
|[document](/javascript/api/word/word.document)|_リレーションシップ_ > properties|現在のドキュメントのプロパティを取得します。読み取り専用。|1.3|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|_メソッド_ > open()|ドキュメントを開きます。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > applicationName|ドキュメントのアプリケーション名を取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > author|ドキュメントの作成者を取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > category|ドキュメントのカテゴリを取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > comments|ドキュメントのコメントを取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > company|ドキュメントの会社を取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > format|ドキュメントの書式設定を取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > keywords|ドキュメントのキーワードを取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > lastAuthor|ドキュメントの最後の作成者を取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > manager|ドキュメントのマネージャーを取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > revisionNumber|ドキュメントのリビジョン番号を取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > security|ドキュメントのセキュリティを取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > subject|ドキュメントの件名を取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > template|ドキュメントのテンプレートを取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_プロパティ_ > title|ドキュメントのタイトルを取得または設定します。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_リレーションシップ_ > creationDate|ドキュメントの作成日を取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_リレーションシップ_ > customProperties|ドキュメントのカスタム プロパティのコレクションを取得します。読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_リレーションシップ_ > lastPrintDate|ドキュメントを最後に印刷した日を取得します。 読み取り専用です。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_リレーションシップ_ > lastSaveTime|ドキュメントを最後に保存した時刻を取得します。 読み取り専用です。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_リレーションシップ_ > parentTable|インライン イメージを含むテーブルを取得します。テーブルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_リレーションシップ_ > parentTableCell|インライン イメージを含むテーブルのセルを取得します。テーブル セルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > getNext()|次のインライン画像を取得します。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > getRange(rangeLocation: RangeLocation)|範囲として、画像、あるいは画像の開始点または終了点を取得します。|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_メソッド_ > getFirst()|このコレクション内の最初のインライン イメージを取得します。|1.3|
|[list](/javascript/api/word/word.list)|_プロパティ_ > id|リストの ID を取得します。読み取り専用。|1.3|
|[list](/javascript/api/word/word.list)|_プロパティ_ > levelExistences|リスト内に 9 つの各レベルが存在するかどうかを確認します。値が true の場合は、レベルが存在することを示します。つまり、そのレベルに少なくとも 1 つのリスト アイテムがあることを意味します。読み取り専用。|1.3|
|[list](/javascript/api/word/word.list)|_リレーションシップ_ > levelTypes|リスト内の 9 レベルのすべての種類を取得します。各種類は、'Bullet', 'Number' または 'Picture' にできます。読み取り専用。|1.3|
|[list](/javascript/api/word/word.list)|_リレーションシップ_ > paragraphs|リスト内の段落を取得します。読み取り専用。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > getLevelParagraphs(level: number)|リスト内の指定したレベルで発生する段落を取得します。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > getLevelString(level: number)|指定したレベルで行頭文字、番号、または画像を文字列として取得します。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|指定した位置に段落を挿入します。insertLocation の値には、'Start'、'End'、'Before'、'After' のいずれかを指定できます。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > setLevelAlignment(level: number, alignment: Alignment)|リスト内の指定したレベルで行頭文字の配置、番号、画像のいずれかを設定します。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|リスト内の指定したレベルで行頭文字の書式を設定します。行頭文字が 'Custom' の場合は、charCode が必要です。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|リスト内の指定したレベルの 2 つのインデントを設定します。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object)|リスト内の指定したレベルで番号付け書式を設定します。|1.3|
|[リスト](/javascript/api/word/word.list)|_メソッド_ > setLevelStartingNumber(level: number, startingNumber: number)|リスト内の指定したレベルで開始番号を設定します。既定値は 1 です。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_プロパティ_ > items|リスト オブジェクトのコレクション。読み取り専用。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_メソッド_ > getById(id: number)|識別子を使用してリストを取得します。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_メソッド_ > getFirst()|このコレクション内の最初のリストを取得します。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_メソッド_ > getItem(index: number)|コレクション内のインデックスを使用して、リスト オブジェクトを取得します。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_プロパティ_ > level|リスト内のアイテムのレベルを取得または設定します。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_プロパティ_ > listString|リスト アイテムの行頭文字、番号、または画像を文字列として取得します。読み取り専用。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_プロパティ_ > siblingIndex|兄弟を基準にしてリスト アイテムの注文番号を取得します。読み取り専用。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_メソッド_ > getAncestor(parentOnly: bool)|親が存在しない場合は、リスト アイテムの親または最も近い先祖を取得します。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_メソッド_ > getDescendants(directChildrenOnly: bool)|リスト アイテムのすべての子孫のリスト アイテムを取得します。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_プロパティ_ > isLastParagraph|段落がその親の本文内の最後の段落であることを示します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_プロパティ_ > isListItem|段落がリスト アイテムであるかどうかを確認します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_プロパティ_ > tableNestingLevel|段落のテーブルのレベルを取得します。段落がテーブル内にない場合は、0 を返します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > list|この段落が属するリストを取得します。段落がリスト内にない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > listItem|段落の ListItem を取得します。段落がリストの一部でない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > parentBody|段落の親の本文を取得します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > parentTable|段落を含むテーブルを取得します。テーブルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > parentTableCell|段落を含むテーブルのセルを取得します。テーブル セルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_リレーションシップ_ > styleBuiltIn|段落の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > attachToList(listId: number, level: number)|指定したレベルで段落を既存のリストに結合させます。段落をリストに結合できない場合、または段落が既にリスト アイテムである場合は、失敗します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > detachFromList()|段落がリスト アイテムである場合は、この段落をリストから移動します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > getNext()|次の段落を取得します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > getPrevious()|前の段落を取得します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > getRange(rangeLocation: RangeLocation)|段落全体、あるいは段落の開始点または終了点を範囲として取得します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > getTextRanges(endingMarks: string, trimSpacing: bool)|句読点と他の終了記号、またはそのいずれかを使用して、段落内のテキスト範囲を取得します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > insertTable(rowCount: number, columnCount: number, insertLocation:InsertLocation, values: string)|指定した数の行と列を含むテーブルを挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > split(delimiters: string, trimDelimiters: bool, trimSpacing: bool)|区切り記号を使用して、段落を子の範囲に分割します。|1.3|
|[段落](/javascript/api/word/word.paragraph)|_メソッド_ > startNewList()|この段落を含む新しいリストを開始します。段落が既にリスト アイテムである場合は失敗します。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_メソッド_ > getFirst()|このコレクション内の最初の段落を取得します。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_メソッド_ > getLast()|このコレクション内の最後の段落を取得します。|1.3|
|[range](/javascript/api/word/word.range)|_プロパティ_ > hyperlink|範囲内の最初のハイパーリンクを取得するか、または範囲にハイパーリンクを設定します。範囲に新しいハイパーリンクを設定すると、範囲内のすべてのハイパーリンクが削除されます。改行文字 ('\n') を使用して、アドレスの部分とオプションの場所の部分を区切ります。|1.3|
|[range](/javascript/api/word/word.range)|_プロパティ_ > isEmpty|範囲の長さが 0 であるかどうかを確認します。読み取り専用。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > lists|範囲内のリスト オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > parentBody|範囲の親の本文を取得します。読み取り専用。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > parentTable|範囲を含むテーブルを取得します。テーブルに含まれていない場合は、null を返します。読み取り専用。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > parentTableCell|範囲を含むテーブルのセルを取得します。テーブル セルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > styleBuiltIn|範囲の組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|1.3|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > tables|範囲内のテーブル オブジェクトのコレクションを取得します。読み取り専用。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > compareLocationWith(range: Range)|この範囲の場所を別の範囲の場所と比較します。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > expandTo(range: Range)|別の範囲を対象にするために、いずれかの方向でこの範囲から拡張する新しい範囲を返します。この範囲は変更されません。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > getHyperlinkRanges()|範囲内のハイパーリンクの子の範囲を取得します。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > getNextTextRange(endingMarks: string, trimSpacing: bool)|句読点と他の終了記号、またはそのいずれかを使用して、次のテキスト範囲を取得します。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > getRange(rangeLocation: RangeLocation)|範囲の複製を作成するか、新しい範囲として開始点または終了点を取得します。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > getTextRanges(endingMarks: string, trimSpacing: bool)|句読点と他の終了記号、またはそのいずれかを使用して、範囲内にあるテキストの子の範囲を取得します。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > insertTable(rowCount: number, columnCount: number, insertLocation:InsertLocation, values: string)|指定した数の行と列を含むテーブルを挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > intersectWith(range: Range)|別の範囲とこの範囲の交点として、新しい範囲を返します。この範囲は変更されません。|1.3|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|区切り記号を使用して、範囲を子の範囲に分割します。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_プロパティ_ > items|範囲オブジェクトのコレクション。読み取り専用。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_メソッド_ > getFirst()|このコレクション内の最初の範囲を取得します。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_メソッド_ > getItem(index: number)|コレクション内のインデックスを使用して、範囲オブジェクトを取得します。|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_メソッド_ > load(object: object, option: object)|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオプションを設定します。 |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_メソッド_ > sync()|要求キューを Word に送信し、さらに多くの操作を連続的に繋ぐために使用できる約束オブジェクトを返します。|1.3|
|[セクション](/javascript/api/word/word.section)|_メソッド_ > getNext()|次のセクションを取得します。|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_メソッド_ > getFirst()|このコレクション内の最初のセクションを取得します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > headerRowCount|ヘッダー行の数を取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > height|テーブルの高さをポイント単位で取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > isUniform|すべてのテーブル行が均一かどうかを示します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > nestingLevel|テーブルの入れ子のレベルを取得します。最上位のテーブルのレベルは、レベル 1 です。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > rowCount|テーブルの行数を取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > shadingColor|網かけの色を取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > style|テーブルのスタイル名を取得または設定します。カスタム スタイルとローカライズされたスタイルの名前には、このプロパティを使用します。ロケール間で移植可能な組み込みスタイルを使用するには、"styleBuiltIn" プロパティを参照してください。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > styleBandedColumns|テーブルの列を縞模様にするかどうかを取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > styleBandedRows|テーブルの行を縞模様にするかどうかを取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > styleFirstColumn|テーブルの最初の列に特別なスタイルを指定するかどうかを取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > styleLastColumn|テーブルの最後の列に特別なスタイルを指定するかどうかを取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > styleTotalRow|テーブルの集計 (最後) 行に特別なスタイルを指定するかどうかを取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > values|2D の Javascript 配列として、テーブルのテキスト値を取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_プロパティ_ > width|テーブルの幅をポイント単位で取得および設定します。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > font|フォントを取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > horizontalAlignment|テーブル内のすべてのセルの水平方向の配置を取得および設定します。値には、"left"、"centered"、"right"、または "justified" を指定できます。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > paragraphAfter|テーブルの後の段落を取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > paragraphBefore|テーブルの前の段落を取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > parentBody|テーブルの親の本文を取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > parentContentControl|テーブルを含むコンテンツ コントロールを取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > parentTable|このテーブルを含むテーブルを取得します。テーブルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > parentTableCell|このテーブルを含むテーブルのセルを取得します。テーブル セルに含まれていない場合は、null オブジェクトを返します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > rows|すべてのテーブルの行を取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > styleBuiltIn|テーブルの組み込みスタイル名を取得または設定します。ロケール間で移植可能な組み込みスタイルの場合は、このプロパティを使用します。カスタム スタイルまたはローカライズされたスタイルの名前を使用するには、"style" プロパティを参照してください。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > tables|1 レベル深く入れ子にされた子テーブルを取得します。読み取り専用。|1.3|
|[table](/javascript/api/word/word.table)|_リレーションシップ_ > verticalAlignment|テーブル内のすべてのセルの垂直方向の配置を取得および設定します。値には、'top'、'center' または 'bottom' を指定できます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|最初または最後の既存の列をテンプレートとして使用して、テーブルの最初または最後に列を追加します。これは、統一されたテーブルに適用可能です。指定すると、文字列値は新しく挿入された行に設定されます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|最初または最後の既存の行をテンプレートとして使用して、テーブルの最初または最後に行を追加します。指定すると、文字列値は新しく挿入された行に設定されます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > autoFitContents()|テーブルの列をコンテンツの幅に合わせて自動調整します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > autoFitWindow()|テーブルの列をウィンドウの幅に合わせて自動調整します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > clear()|テーブルの内容をクリアします。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > delete()|テーブル全体を削除します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > deleteColumns(columnIndex: number, columnCount: number)|特定の列を削除します。これは、統一されたテーブルに適用可能です。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > deleteRows(rowIndex: number, rowCount: number)|特定の行を削除します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > distributeColumns()|列の幅を揃えます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > distributeRows()|行の高さを揃えます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > getBorder(borderLocation: BorderLocation)|指定した罫線の罫線スタイルを取得します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > getCell(rowIndex: number, cellIndex: number)|指定された行と列のテーブル セルを取得します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|セル内のスペースをポイント単位で取得します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > getNext()|次のテーブルを取得します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > getRange(rangeLocation: RangeLocation)|このテーブルを含む範囲、あるいはテーブルの開始または終了の範囲を取得します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > insertContentControl()|テーブルにコンテンツ コントロールを挿入します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > insertTable(rowCount: number, columnCount: number, insertLocation:InsertLocation, values: string)|指定した数の行と列を含むテーブルを挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|テーブル オブジェクトの範囲で、searchOptions を指定した検索を実行します。検索結果は、範囲オブジェクトのコレクションになります。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > select(selectionMode: SelectionMode)|テーブル、あるいはテーブルの開始位置または終了位置を選択して、Word の UI に移動します。|1.3|
|[テーブル](/javascript/api/word/word.table)|_メソッド_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|セル内のスペースをポイント単位で設定します。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_プロパティ_ > color|16 進数値または名前として、テーブルの罫線の色を取得または設定します。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_プロパティ_ > width|テーブルの罫線の幅をポイント単位で得または設定します。幅が固定されているテーブルの罫線の種類には適用できません。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_リレーションシップ_ > type|テーブルの罫線の種類を取得または設定します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > cellIndex|その行のセルのインデックスを取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > columnWidth|セルの列の幅をポイント単位で取得または設定します。これは、統一されたテーブルに適用可能です。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > rowIndex|テーブルのセル行のインデックスを取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > shadingColor|セルの網かけの色を取得または設定します。色は、"#RRGGBB" 形式で指定するか、色の名前を使用して指定します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > value|セルのテキストを取得および設定します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_プロパティ_ > width|セルの幅をポイント単位で取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_リレーションシップ_ > body|セルの本文オブジェクトを取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_リレーションシップ_ > horizontalAlignment|セルの水平方向の配置を取得および設定します。値には、"left"、"centered"、"right"、または "justified" を指定できます。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_リレーションシップ_ > parentRow|セルの親行を取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_リレーションシップ_ > parentTable|セルの親テーブルを取得します。読み取り専用。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_リレーションシップ_ > verticalAlignment|セルの垂直方向の配置を取得および設定します。値には、'top'、'center' または 'bottom' を指定できます。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > deleteColumn()|このセルを含む列を削除します。これは、統一されたテーブルに適用可能です。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > deleteRow()|このセルを含む行を削除します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > getBorder(borderLocation: BorderLocation)|指定した罫線の罫線スタイルを取得します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|セル内のスペースをポイント単位で取得します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > getNext()|次のセルを取得します。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > insertColumns(insertLocation: InsertLocation, columnCount: number, values: string)|セルの列をテンプレートとして使用して、列をセルの左または右に追加します。これは、統一されたテーブルに適用可能です。指定すると、文字列値は新しく挿入された行に設定されます。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|セルの行をテンプレートとして使用して、行をセルの上または下に挿入します。指定すると、文字列値は新しく挿入された行に設定されます。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_メソッド_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|セル内のスペースをポイント単位で設定します。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_プロパティ_ > items|tableCell オブジェクトのコレクション。読み取り専用。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_メソッド_ > getFirst()|このコレクション内の最初のテーブル セルを取得します。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_メソッド_ > getItem(index: number)|コレクション内のインデックスを使用して、テーブル セル オブジェクトを取得します。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_プロパティ_ > items|Table オブジェクトのコレクション。読み取り専用です。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_メソッド_ > getFirst()|このコレクション内の最初のテーブルを取得します。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_メソッド_ > getItem(index: number)|コレクション内のインデックスを使用して、テーブル オブジェクトを取得します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > cellCount|行のセルの数を取得します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > isHeader|行がヘッダー行であるかどうかを確認します。読み取り専用。ヘッダー行の数を設定するには、テーブル オブジェクトの HeaderRowCount を使用します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > preferredHeight|適切な行の高さをポイント単位で取得および設定します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > rowIndex|親テーブル内の行のインデックスを取得します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > shadingColor|網かけの色を取得および設定します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_プロパティ_ > values|1D の Javascript 配列として、行のテキスト値を取得および設定します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_リレーションシップ_ > cells|セルを取得します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_リレーションシップ_ > font|フォントを取得します。これを使用して、フォント名、サイズ、色、およびその他のプロパティを取得および設定します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_リレーションシップ_ > horizontalAlignment|行のすべてのセルの水平方向の配置を取得および設定します。値には、"left"、"centered"、"right"、または "justified" を指定できます。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_リレーションシップ_ > parentTable|親テーブルを取得します。読み取り専用。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_リレーションシップ_ > verticalAlignment|行のセルの垂直方向の配置を取得および設定します。値には、'top'、'center' または 'bottom' を指定できます。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > clear()|行の内容をクリアします。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > delete()|行全体を削除します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > getBorder(borderLocation: BorderLocation)|行のセルの罫線スタイルを取得します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|セル内のスペースをポイント単位で取得します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > getNext()|次の行を取得します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|この行をテンプレートとして使用して、行を挿入します。値を指定すると、新しい行に値を挿入します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|行の範囲で、指定した searchOptions を使って検索を実行します。検索結果は、範囲オブジェクトのコレクションになります。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > select(selectionMode: SelectionMode)|行を選択し、その行に Word の UI を移動します。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_メソッド_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|セル内のスペースをポイント単位で設定します。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_プロパティ_ > items|tableRow オブジェクトのコレクション。読み取り専用。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_メソッド_ > getFirst()|このコレクション内の最初の行を取得します。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_メソッド_ > getItem(index: number)|コレクション内のインデックスを使用して、テーブル行オブジェクトを取得します。|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 の新機能

要件セット 1.2 の Word JavaScript API に新たに追加された機能は次のとおりです。 

|オブジェクト| 新機能| 説明|要件セット|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_メソッド_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|コンテンツ コントロール内の指定された位置にインライン画像を挿入します。insertLocation の値は、'Replace'、'Start'、'End' のいずれかになります。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_リレーションシップ_ > paragraph|インライン イメージを含む親段落を取得します。読み取り専用。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > delete()|ドキュメントからインライン画像を削除します。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertBreak(breakType: BreakType, insertLocation: InsertLocation)|メイン文書の指定した位置に、区切りを挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|指定した位置に文書を挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertHtml(html: string, insertLocation: InsertLocation)|指定した位置に HTML を挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation|指定された位置にインライン画像を挿入します。insertLocation の値には、'Replace'、'Before'、'After' のいずれかを指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertOoxml(ooxml: string, insertLocation: InsertLocation)|指定した位置に、OOXML を挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|指定した位置に、段落を挿入します。有効な insertLocation の値は、'Before' または 'After' です。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > insertText(text: string, insertLocation: InsertLocation)|指定した位置にテキストを挿入します。insertLocation の値には、'Before' または 'After' を指定できます。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_メソッド_ > select(selectionMode: SelectionMode)|インライン画像を選択します。その結果、Word は選択範囲にスクロールされます。|1.2|
|[range](/javascript/api/word/word.range)|_リレーションシップ_ > inlinePictures|範囲に含まれるインライン画像オブジェクトのコレクションを取得します。読み取り専用。|1.2|
|[範囲](/javascript/api/word/word.range)|_メソッド_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|指定された位置に画像を挿入します。insertLocation の値には、'Replace'、'Start'、'End'、'Before'、'After' のいずれかを指定できます。|1.2|

## <a name="word-javascript-api-11"></a>Word JavaScript API 1.1

Word JavaScript API 1.1 は、API の最初のバージョンです。API の詳細については、[Word JavaScript API](/javascript/api/word) リファレンスのトピックを参照してください。 

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
