---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットの詳細
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: aa497ff97533ff3a414905547a949fa8430c3efe
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430815"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API のオンラインのみの要件セット

`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。 この要件セットの Api は、web アプリケーション上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。 `ExcelApiOnline` 他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。

要件セットの Api `ExcelApiOnline` がすべてのプラットフォームでサポートされている場合は、次にリリースされる要件セット () に追加され `ExcelApi 1.[NEXT]` ます。 新しい要件が公開されると、これらの Api はから削除され `ExcelApiOnline` ます。 この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 Api の唯一のバージョンです。 これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。

## <a name="recommended-usage"></a>推奨される使用法

`ExcelApiOnline`Api は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。 これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

クロスプラットフォームの要件セットに含まれる API は、チェックを削除または編集する必要があり `isSetSupported` ます。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行うときは、これらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェスト `ExcelApiOnline 1.1` でライセンス認証の要件として指定することはできません。 [Set 要素](../manifest/set.md)で使用する有効な値ではありません。

## <a name="api-list"></a>API リスト

次の Api は、現在、要件セットの一部として web 上の Excel で使用でき `ExcelApiOnline 1.1` ます。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|グラフ軸のタイトルに対して、テキストの方向を指定する角度を指定します。 この値は、-90 ~ 90 の整数、または垂直方向のテキストの整数の180のいずれかである必要があります。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|コレクション内のピボットテーブルの数を取得します。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|コレクション内の最初のピボットテーブルを取得します。 コレクション内のピボットテーブルは、上から下、左から右に並べ替えられます。この場合、左上のテーブルはコレクションの最初のピボットテーブルになります。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|名前に基づいてピボットテーブルを取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|名前を使用してピボットテーブルを取得します。 PivotTable が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables テーブル (fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|範囲に重なっているピボットテーブルのスコープ設定されたコレクションを取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](./excel-preview-apis.md)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)