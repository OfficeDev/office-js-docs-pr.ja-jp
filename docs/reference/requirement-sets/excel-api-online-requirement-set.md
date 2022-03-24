---
title: Excel JavaScript API オンライン専用要件セット
description: ExcelApiOnline 要件セットの詳細。
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f3ec510e889ecfe565767352c59cd349e0701830
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746607"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API オンライン専用要件セット

要件`ExcelApiOnline`セットは、ユーザーが使用できる機能のみを含む特別な要件セットExcel on the web。 この要件セットの API は、アプリケーションの実稼働 API と見なされます (文書化されていない動作や構造上の変更のExcel on the webされます。 `ExcelApiOnline`API は、他のプラットフォーム (Windows、Mac、iOS) の "プレビュー" API と見なされ、これらのプラットフォームではサポートされない場合があります。

要件セット内の `ExcelApiOnline` API がすべてのプラットフォームでサポートされている場合は、次にリリースされた要件セット () に追加されます`ExcelApi 1.[NEXT]`。 その新しい要件が公開されると、これらの API はから削除されます `ExcelApiOnline`。 これは、プレビューからリリースに移行する API と同様のプロモーション プロセスと考えて下さい。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 API の唯一のバージョンです。 これは、最新Excel on the webユーザーが常に 1 つのバージョンを使用できるためです。

次の表に、API の簡潔な概要を示しますが、後続の [API](#api-list) リスト テーブルでは、現在の API の詳細な一覧を `ExcelApiOnline` 示します。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| リンクされたブック | ブック間のリンクを管理します。ブックリンクの更新と破損のサポートを含む。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| ワークシートの移動イベント | コレクション内でワークシートを移動する場合、ワークシートの位置、変更元を検出します。 | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>推奨される使用法

API は `ExcelApiOnline` Excel on the webによってのみサポートされるので、アドインは、これらの API を呼び出す前に、要件セットがサポートされているのか確認する必要があります。 これにより、別のプラットフォームでオンライン専用 API を呼び出すのを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API がクロスプラットフォーム要件セットに入った後は、チェックを削除または編集する必要 `isSetSupported` があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行う場合は、必ずこれらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証 `ExcelApiOnline 1.1` 要件を指定することはできません。 Set 要素で使用する有効な値 [ではありません](../manifest/set.md)。

## <a name="api-list"></a>API リスト

次の表に、要件Excel含まれている JavaScript API の一覧を`ExcelApiOnline`示します。 すべての JavaScript API (`ExcelApiOnline`API Excel以前にリリースされた API を含む) の完全な一覧については、[JavaScript API Excel参照してください](/javascript/api/excel?view=excel-js-online&preserve-view=true)。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|リンクされたブックを指すリンクを壊す要求を行います。|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|リンクされたブックを指す元の URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|リンクされたブックから取得したデータを更新する要求を行います。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|リンクされたブックへのすべてのリンクを壊します。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|リンクされたブックに関する情報を URL で取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|リンクされたブックに関する情報を URL で取得します。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|すべてのブック リンクを更新する要求を行います。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|ブック リンクの更新モードを表します。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|指定した名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|コレクション内のインデックスによってシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|テーブルから複数の行を削除します。|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|指定したインデックスから、指定した数の行をテーブルから削除します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|リンクされたブックのコレクションを返します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|ワークシートに存在するシート ビューのコレクションを返します。|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|ワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|ワークシートの表示設定が変更された場合に発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|ワークシートがブック内のユーザーによって移動された場合に発生します。|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|ワークシートコレクションでワークシート名が変更された場合に発生します。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|ワークシート コレクションでワークシートの表示設定が変更された場合に発生します。|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|移動後のワークシートの新しい位置を取得します。|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|移動の前に、ワークシートの前の位置を取得します。|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|移動したワークシートの ID を取得します。|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|名前の変更後に、ワークシートの新しい名前を取得します。|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|名前が変更される前に、ワークシートの前の名前を取得します。|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|新しい名前を持つワークシートの ID を取得します。|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|イベントの種類を取得します。|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|表示設定の変更後に、ワークシートの新しい表示設定を取得します。|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|表示設定が変更される前に、ワークシートの以前の表示設定を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|表示が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](excel-preview-apis.md)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
