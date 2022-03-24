---
title: Excel JavaScript API 要件セット 1.3
description: ExcelApi 1.3 要件セットの詳細。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1bf8bc604c2c770f517878193994c1ed32640da1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745340"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

ExcelApi 1.3 では、データ バインドと基本的なピボットテーブル アクセスのサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.3 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.3 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#excel-excel-binding-delete-member(1))|バインドを削除します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-add-member(1))|特定の範囲に新しいバインドを追加します。|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromnameditem-member(1))|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromSelection(bindingType: Excel.BindingType、 id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromselection-member(1))|現在の選択範囲に基づいて新しいバインドを追加します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-name-member)|ピボットテーブルの名前。|
||[refresh()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refresh-member(1))|ピボットテーブルを更新します。|
||[worksheet](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-worksheet-member)|現在のピボットテーブルを含んでいるワークシート。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitem-member(1))|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-refreshall-member(1))|コレクション内のすべてのピボットテーブルを更新します。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#excel-excel-range-getvisibleview-member(1))|現在の範囲の表示されている行を表します。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-celladdresses-member)|のセル アドレスを表します `RangeView`。|
||[columnCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-columncount-member)|表示される列の数。|
||[formulas](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulas-member)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulaslocal-member)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulasr1c1-member)|R1C1 スタイル表記の数式を表します。|
||[getRange()](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-getrange-member(1))|現在の値に関連付けられている親範囲を取得します `RangeView`。|
||[index](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-index-member)|のインデックスを表す値を返します `RangeView`。|
||[numberFormat](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-numberformat-member)|指定したセルの Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rowcount-member)|表示される行の数。|
||[rows](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rows-member)|範囲に関連付けられている範囲ビューのコレクションを表します。|
||[text](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-text-member)|指定した範囲のテキスト値。|
||[valueTypes](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuetypes-member)|各セルのデータの種類を表します。|
||[values](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-values-member)|指定した範囲ビューの Raw 値を表します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getitemat-member(1))|インデックスを使用 `RangeView` して行を取得します。|
||[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightfirstcolumn-member)|最初の列に特別な書式が含まれている場合に指定します。|
||[highlightLastColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightlastcolumn-member)|最後の列に特別な書式が含まれている場合に指定します。|
||[showBandedColumns](/javascript/api/excel/excel.table#excel-excel-table-showbandedcolumns-member)|テーブルの読み取りを容易にするために、奇数列が偶数列とは異なる方法で強調表示されるバンド書式を列に表示する場合に指定します。|
||[showBandedRows](/javascript/api/excel/excel.table#excel-excel-table-showbandedrows-member)|テーブルの読み取りを容易にするために、奇数行が偶数行とは異なる方法で強調表示されるバンド書式を行に表示する場合に指定します。|
||[showFilterButton](/javascript/api/excel/excel.table#excel-excel-table-showfilterbutton-member)|フィルター ボタンが各列ヘッダーの上部に表示される場合に指定します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[ピボットテーブル](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottables-member)|ブックに関連付けられているピボットテーブルのコレクションを表します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[ピボットテーブル](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pivottables-member)|ワークシートの一部になっているピボットテーブルのコレクション。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
