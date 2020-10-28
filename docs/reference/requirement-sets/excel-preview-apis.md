---
title: Excel JavaScript プレビュー API
description: 今後の Excel JavaScript Api についての詳細。
ms.date: 10/26/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1cb3afb28f69ff5b0c0bd03bfae9877dda91906
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774741"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript プレビュー API

新しい Excel JavaScript API は最初に "プレビュー" で導入され、その後、十分なテストが行われ、ユーザー フィードバックが得られてから、番号付きの特定の要件セットの一部になります。

最初の表には API が簡潔にまとめられています。その後の表は詳しい一覧になっています。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| リンクされたデータ型 | 外部ソースから Excel に接続されたデータ型のサポートを追加します。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 指定したシートビュー | ユーザー単位のワークシートビューをプログラムによって制御します。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="api-list"></a>API リスト

次の表に、現在プレビュー中の Excel JavaScript Api を示します。 すべての Excel JavaScript Api (プレビュー Api および以前リリースされた Api を含む) の完全なリストについては、「 [すべての Excel Javascript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[プロバイダー](/javascript/api/excel/excel.linkeddatatype#dataprovider)|リンクされたデータ型のデータプロバイダーの名前を指定します。 これは、情報がサービスから取得されたときに変わる可能性があります。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|リンクされたデータ型が最後に更新されたときに、ブックが開かれてからのローカルタイムゾーンの日付と時刻。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|リンクされたデータ型の名前を指定します。 これは、情報がサービスから取得されたときに変わる可能性があります。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|リンクされたデータ型が `refreshMode` "定期的" に設定されている場合に更新される頻度 (秒単位)。|
||[示し](/javascript/api/excel/excel.linkeddatatype#refreshmode)|リンクされたデータ型のデータを取得するメカニズムを指定します。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|リンクされたデータ型の一意の id。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|リンクされたデータ型によってサポートされるすべての更新モードを含む配列を返します。 サービスから情報を取得すると、配列の内容が変わる可能性があります。|
||[requestRefresh ()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|リンクされたデータ型を更新する要求を行います。 サービスがビジーである場合、または一時的にアクセスできない場合、要求は満たされません。|
||[requestSetRefreshMode (refreshMode: LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|このリンクされたデータ型の更新モードを変更する要求を行います。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新しいリンクされたデータ型の一意の id。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|コレクション内のリンクされたデータ型の数を取得します。|
||[getItem (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|リンクされたデータ型をサービス id で取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|コレクション内のインデックスによって、リンクされたデータ型を取得します。|
||[getItemOrNullObject (key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|ID でリンクされたデータ型を取得します。 リンクされたデータ型が存在しない場合は、そのプロパティがに設定されたオブジェクト `isNullObject` `true` 。 詳細については、「{@link」を参照してください。 https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | * OrNullObject メソッドとプロパティ}。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|コレクション内のすべてのリンクされたデータ型を更新する要求を行います。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシートビューをアクティブにします。 これは、Excel UI の [切り替え先] を使用するのと同じです。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシートビューを削除します。|
||[重複 (名前?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシートビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シートビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定した名前の新しいシートビューを作成します。|
||[enterTemporary ()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シートビューを作成してアクティブにします。|
||[exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシートビューを終了します。|
||[getActive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシートビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシートビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシートビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスによってシートビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|ピボットテーブルの代替テキストタイトル。|
||[各アイテムを表示する (display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|各アイテムの後に空白行を表示するかどうかを設定します。 これは、ピボットテーブルのグローバルレベルで設定され、個々のピボットフィールドに適用されます。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|ピボットテーブル内の空のセルに自動的に入力されるテキスト `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|ピボットテーブルの空のセルにを設定するかどうかを指定し `emptyCellText` ます。 既定では False。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|データ階層と、それぞれの階層の行および列の項目に基づいて、ピボットテーブル内の一意のセルを取得します。  返されるセルは、指定した階層のデータが含まれる、指定された行と列の交差部分です。  このメソッドは、特定のセルでの getPivotItems および getDataHierarchy の呼び出しを逆にしたものです。|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|ピボットテーブルのすべてのフィールドで [すべてのアイテムのラベルを繰り返す] 設定を設定します。|
||[setStyle (style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|ピボットテーブルに適用されるスタイルを設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|ピボットテーブルにフィールドヘッダーを表示するかどうかを指定します (フィールドのタイトルとフィルターのドロップダウン)。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|ブックを開くときにピボットテーブルを更新するかどうかを指定します。 UI の [読み込み時に更新] 設定に相当します。|
|[Range](/javascript/api/excel/excel.range)|[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|`RangeAreas`この範囲内の結合された領域を表すオブジェクトを返します。 この範囲内のマージされた領域の数が512を超える場合、API は結果を返すことに失敗します。|
||[getPrecedents 元 ()](/javascript/api/excel/excel.range#getprecedents--)|`WorkbookRangeAreas`同じワークシートまたは複数のワークシート内のセルのすべての参照元を含む範囲を表すオブジェクト型 (object) の値を取得します。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[示し](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|リンクされたデータ型の更新モード。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|更新モードが変更されたオブジェクトの一意の id です。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[更新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|更新要求が正常に終了したかどうかを示します。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|更新要求が完了したオブジェクトの一意の id。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|イベントのソースを取得します。 詳細については、Excel.EventSource をご覧ください。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[注意](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|更新要求によって生成された警告を含む配列。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|XML 文字列からスケーラブルなベクター グラフィックス (SVG) を作成し、それをワークシートに追加します。 新しい画像を表す Shape オブジェクトを返します。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|数式で使用するスライサーの名前を表します。|
||[setStyle (style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|スライサーに適用されるスタイルを設定します。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|既定のテーブル スタイルを使用するようにテーブルを変更します。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|フィルターが特定のテーブルに適用されたときに発生します。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|表に適用されるスタイルです。|
||[setStyle (style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|表に適用するスタイルを設定します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|ブックまたはワークシートのテーブルにフィルターが適用されたときに発生します。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|フィルターが適用されているテーブルの id を取得します。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|テーブルを含むワークシートの id を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes 型](/javascript/api/excel/excel.workbook#linkeddatatypes)|ブックの一部である、リンクされたデータ型のコレクションを返します。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|ピボットテーブルのフィールドリストウィンドウをブックレベルで表示するかどうかを指定します。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|ブックの日付を 1904 年から計算する場合、true となります。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートにあるシートビューのコレクションを返します。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|フィルターが特定のワークシートに適用されたときに発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|あるブックの指定されたワークシートを現在のブックに挿入します。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|ブック内でワークシートのフィルターが適用されたときに発生します。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|イベントの種類を取得します。 詳細については、Excel.EventType をご覧ください。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|フィルターが適用されているワークシートの id を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
