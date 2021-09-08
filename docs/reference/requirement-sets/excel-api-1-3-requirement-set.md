---
title: ExcelJavaScript API 要件セット 1.3
description: ExcelApi 1.3 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3606b74e8a1099cd58631cc047a783f27a09a19
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937373"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

ExcelApi 1.3 では、データ バインドと基本的なピボットテーブル アクセスのサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.3 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.3 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット[1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|バインドを削除します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|特定の範囲に新しいバインドを追加します。|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromSelection(bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|現在の選択範囲に基づいて新しいバインドを追加します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|現在のピボットテーブルを含んでいるワークシート。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|ピボットテーブルを更新します。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|コレクション内のすべてのピボットテーブルを更新します。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getVisibleView__)|現在の範囲の表示されている行を表します。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|R1C1 スタイル表記の数式を表します。|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|現在の値に関連付けられている親範囲を取得します `RangeView` 。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|指定したセルの Excel の数値書式コードを表します。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|のセル アドレスを表します `RangeView` 。|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|表示される列の数。|
||[index](/javascript/api/excel/excel.rangeview#index)|のインデックスを表す値を返します `RangeView` 。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|表示される行の数。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|範囲に関連付けられている範囲ビューのコレクションを表します。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|各セルのデータの種類を表します。|
||[values](/javascript/api/excel/excel.rangeview#values)|指定した範囲ビューの Raw 値を表します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|インデックスを `RangeView` 使用して行を取得します。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|最初の列に特別な書式が含まれている場合に指定します。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|最後の列に特別な書式が含まれている場合に指定します。|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|テーブルの読み取りを容易にするために、奇数列が偶数列とは異なる方法で強調表示されるバンド書式を列に表示する場合に指定します。|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|テーブルの読み取りを容易にするために、奇数行が偶数行とは異なる方法で強調表示されるバンド書式を行に表示する場合に指定します。|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|フィルター ボタンが各列ヘッダーの上部に表示される場合に指定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[ピボットテーブル](/javascript/api/excel/excel.workbook#pivotTables)|ブックに関連付けられているピボットテーブルのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[ピボットテーブル](/javascript/api/excel/excel.worksheet#pivotTables)|ワークシートの一部になっているピボットテーブルのコレクション。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
