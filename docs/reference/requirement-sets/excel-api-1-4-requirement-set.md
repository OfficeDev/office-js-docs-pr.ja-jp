---
title: Excel JavaScript API 要件セット1.4
description: ExcelApi 1.4 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1bd6d913bf069e4b8774b8eb65ea147992f98b9b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611429"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 の新機能

要件セット 1.4 の Excel JavaScript API に新たに追加された機能は次のとおりです。

## <a name="named-item-add-and-new-properties"></a>名前付きアイテムの追加と新しいプロパティ

新しいプロパティ:

* `comment`
* `scope`-ワークシートまたはブックを対象範囲とするアイテム。
* `worksheet`-名前付きアイテムのスコープが設定されているワークシートを返します。

新しいメソッド:

* `add(name: string, reference: Range or string, comment: string)`-指定したスコープのコレクションに新しい名前を追加します。
* `addFormulaLocal(name: string, formula: string, comment: string)`-式のユーザーのロケールを使用して、指定したスコープのコレクションに新しい名前を追加します。

## <a name="settings-api-in-the-excel-namespace"></a>Excel 名前空間内の Setting API

[Setting](/javascript/api/excel/excel.setting) オブジェクトは、ドキュメントに永続的に適用される設定のキーと値のペアを表します。 `Excel.Setting` の機能は `Office.Settings` と同等ですが、共通 API のコールバック モデルではなくバッチ API 構文を使用します。

Api は `getItem()` 、キーを使用して設定エントリを取得し、 `add()` 指定されたキー: 値の設定のペアをブックに追加するために含まれます。

## <a name="others"></a>Others

* 表の列名を設定します。
* 表の列を表の末尾に追加します。
* 一度に複数の行をテーブルに追加します。
* `range.getColumnsAfter(count: number)` および `range.getColumnsBefore(count: number)` を使用して、現在の Range オブジェクトの左右にある特定の数の列を取得します。
* [Get item または null オブジェクト関数](../../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods): この機能は、キーを使用してオブジェクトを取得することを可能にします。 オブジェクトが存在しない場合、返されるオブジェクトの `isNullObject` プロパティは true になります。 これにより、開発者は、オブジェクトが存在するかどうかを確認することができます。ただし、例外処理によって処理する必要はありません。 この `*OrNullObject` メソッドは、ほとんどのコレクションオブジェクトで使用できます。

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.4 の Api を示します。 Excel JavaScript API 要件セット1.4 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「[要件セット1.4 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.4)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|コレクションに含まれるバインドの数を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|ID によってバインド オブジェクトを取得します。 バインディング オブジェクトが存在しない場合は null オブジェクトを返します。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|ワークシート上のグラフの数を返します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|グラフ名を使用してグラフを取得します。 同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|系列に含まれるグラフのポイントの数を返します。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|コレクションに含まれるデータ系列の数を返します。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|この名前に関連付けられているコメントを表します。|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|指定された名前を削除します。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|名前に関連付けられている範囲オブジェクトを返します。 名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。|
||[scope](/javascript/api/excel/excel.nameditem#scope)|ブックまたは特定のワークシートに対して名前のスコープを設定するかどうかを示します。 可能な値は次のとおりです。ワークシート、ブック。 読み取り専用です。|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|名前付きのアイテムの対象になるワークシートを返します。 アイテムのスコープがブックに設定されている場合は、エラーをスローします。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|名前付きのアイテムの対象になるワークシートを返します。 アイテムがブックを対象にしている場合は、null オブジェクトを返します。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add (name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|指定のスコープのコレクションに新しい名前を追加します。|
||[addFormulaLocal (name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|ユーザーのロケールを数式に使用して、指定のスコープのコレクションに新しい名前を追加します。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|コレクションに含まれる名前付きアイテムの数を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|名前を使用して、NamedItem オブジェクトを取得します。 nameditem オブジェクトが存在しない場合は null オブジェクトを返します。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|コレクションに含まれるピボット テーブルの数を取得します。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|名前を使用してピボットテーブルを取得します。 PivotTable が存在しない場合は null オブジェクトを返します。|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Range \| 文字列)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|指定した範囲の長方形の交差を表す範囲オブジェクトを取得します。 交差部分が見つからない場合は、null オブジェクトを返します。|
||[getUsedRangeOrNullObject (パラメーターの設定のみ?: boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|コレクションに含まれる RangeView オブジェクトの数を取得します。|
|[設定](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|設定を削除します。|
||[key](/javascript/api/excel/excel.setting#key)|Setting の ID を表すキーを返します。 読み取り専用です。|
||[value](/javascript/api/excel/excel.setting#value)|この設定に格納されている値を表します。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add (key: string, value: string \| number \| boolean \| Date \| Array <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|指定した設定をブックに設定または追加します。|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|コレクションに含まれる設定の数を取得します。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|キーに基づいて設定エントリを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|キーから Setting エントリを取得します。 Setting が存在しない場合は null オブジェクトを返します。|
||[items](/javascript/api/excel/excel.settingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|ドキュメント内の設定が変更されるときに発生します。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|コレクションに含まれるテーブルの数を取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|名前または ID でテーブルを取得します。 テーブルが存在しない場合は null オブジェクトを返します。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|表の列数を取得します。|
||[getItemOrNullObject (key: number \| 文字列)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|名前または ID によって、列オブジェクトを取得します。 列が存在しない場合は null オブジェクトを返します。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|表の行数を取得します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|ブックに関連付けられている Setting のコレクションを表します。 読み取り専用です。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (パラメーターの設定のみ?: boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。|
||[姓名](/javascript/api/excel/excel.worksheet#names)|現在のワークシートにスコープされている名前のコレクション。 読み取り専用です。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|コレクションに含まれるワークシートの数を取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|名前または ID を使用して、ワークシート オブジェクトを取得します。 ワークシートが存在しない場合は null オブジェクトを返します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.4)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
