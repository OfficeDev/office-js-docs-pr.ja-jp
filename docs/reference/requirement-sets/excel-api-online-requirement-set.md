---
title: ExcelJavaScript API のオンライン専用要件セット
description: ExcelApiOnline 要件セットの詳細。
ms.date: 07/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ef4831cf6a6f9be1a5413c89ae0f971bef51a9b1
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290804"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>ExcelJavaScript API のオンライン専用要件セット

要件セットは、ユーザーが使用できる機能のみを含む特別な要件 `ExcelApiOnline` セットExcel on the web。 この要件セットの API は、アプリケーションの実稼働 API (文書化されていない動作や構造上の変更の対象ではない) とExcel on the webされます。 `ExcelApiOnline`API は、他のプラットフォーム (Windows、Mac、iOS) の 「プレビュー」 API と見なされ、これらのプラットフォームではサポートされない場合があります。

要件セット内の API がすべてのプラットフォームでサポートされている場合は、次にリリースされた要件セット ( ) に `ExcelApiOnline` 追加されます `ExcelApi 1.[NEXT]` 。 その新しい要件が公開されると、これらの API はから削除されます `ExcelApiOnline` 。 これは、プレビューからリリースに移行する API と同様のプロモーション プロセスと考えて下さい。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 API の唯一のバージョンです。 これは、最新Excel on the webユーザーが常に 1 つのバージョンを使用できるためです。

次の表に、API の簡潔な概要を示しますが、後続の API リスト テーブルでは、現在の [API](#api-list) の詳細な一覧を `ExcelApiOnline` 示します。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>推奨される使用法

API はユーザーによってのみサポートExcel on the web、アドインは、これらの API を呼び出す前に要件セットがサポートされていない `ExcelApiOnline` か確認する必要があります。 これにより、別のプラットフォームでオンライン専用 API を呼び出すのを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API がクロスプラットフォーム要件セットに入った後は、チェックを削除または編集する必要 `isSetSupported` があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行う場合は、必ずこれらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証 `ExcelApiOnline 1.1` 要件を指定することはできません。 Set 要素で使用する有効な値 [ではありません](../manifest/set.md)。

## <a name="api-list"></a>API リスト

次の表に、要件Excel含まれている JavaScript API の一覧を `ExcelApiOnline` 示します。 すべての JavaScript API (API Excel以前にリリースされた API を含む) の完全な一覧については `ExcelApiOnline` [、JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true)API Excel参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|このシート ビューをアクティブ化します。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|ワークシートからシート ビューを削除します。|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|このシート ビューのコピーを作成します。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|シート ビューの名前を取得または設定します。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|指定した名前の新しいシート ビューを作成します。|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|新しい一時シート ビューを作成してアクティブ化します。|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|現在アクティブなシート ビューを終了します。|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|ワークシートの現在アクティブなシート ビューを取得します。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|このワークシートのシート ビューの数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|名前を使用してシート ビューを取得します。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|コレクション内のインデックスによってシート ビューを取得します。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートに存在するシート ビューのコレクションを返します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](excel-preview-apis.md)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
