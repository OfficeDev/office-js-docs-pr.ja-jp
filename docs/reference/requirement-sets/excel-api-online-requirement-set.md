---
title: Excel JavaScript API オンライン専用要件セット
description: ExcelApiOnline 要件セットの詳細。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 282e11e415d51a6724715091d894df64ebaabfae
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604681"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API オンライン専用要件セット

要件セットは、Web 上の Excel でのみ使用できる機能を含む `ExcelApiOnline` 特別な要件セットです。 この要件セットの API は、Web アプリケーション上の Excel の実稼働 API (文書化されていない動作や構造上の変更の対象ではない) と見なされます。 `ExcelApiOnline` API は、他のプラットフォーム (Windows、Mac、iOS) の "プレビュー" API と見なされ、これらのプラットフォームではサポートされない場合があります。

要件セット内の API がすべてのプラットフォームでサポートされている場合は、次にリリースされた要件セット ( ) に `ExcelApiOnline` 追加されます `ExcelApi 1.[NEXT]` 。 その新しい要件が公開されると、これらの API はから削除されます `ExcelApiOnline` 。 これは、プレビューからリリースに移行する API と同様のプロモーション プロセスと考えて下さい。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 API の唯一のバージョンです。 これは、Web 上の Excel には常に、最新バージョンのユーザーが使用できる 1 つのバージョンが含まれるためです。

次の表に、API の簡潔な概要を示しますが、後続の API リスト テーブルでは、現在の [API](#api-list) の詳細な一覧を `ExcelApiOnline` 示します。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| 名前付きシート ビュー | ユーザーごとのワークシート ビューをプログラムで制御できます。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>推奨される使用法

API は Web 上の Excel でのみサポートされますので、アドインは、これらの API を呼び出す前に要件セットがサポートされるのか `ExcelApiOnline` 確認する必要があります。 これにより、別のプラットフォームでオンライン専用 API を呼び出すのを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API がクロスプラットフォーム要件セットに入った後は、チェックを削除または編集する必要 `isSetSupported` があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行う場合は、必ずこれらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証 `ExcelApiOnline 1.1` 要件を指定することはできません。 Set 要素で使用する有効な値 [ではありません](../manifest/set.md)。

## <a name="api-list"></a>API リスト

次の表に、要件セットに現在含まれている Excel JavaScript API の一 `ExcelApiOnline` 覧を示します。 すべての Excel JavaScript API (API と以前にリリースされた API を含む) の完全なリストについては、 `ExcelApiOnline` [すべての Excel JavaScript API を参照してください](/javascript/api/excel?view=excel-js-online&preserve-view=true)。

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
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|指定された方向に基づいて、現在の範囲と範囲の端までの範囲オブジェクトを返します。|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|この範囲内の `RangeAreas` 結合領域を表すオブジェクトを返します。|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|指定された方向に対応するデータ領域のエッジ セルである範囲オブジェクトを返します。|
|[表](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|テーブルのサイズを新しい範囲に変更します。|
|[ワークシート](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|ワークシートに存在するシート ビューのコレクションを返します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](excel-preview-apis.md)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
