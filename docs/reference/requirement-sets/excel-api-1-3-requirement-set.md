---
title: Excel JavaScript API 要件セット1.3
description: ExcelApi 1.3 の要件セットの詳細
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 71bad9fae630a11688458e4cb76ded2fa523a563
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430871"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

ExcelApi 1.3 には、データバインドと基本的なピボットテーブルアクセスのサポートが追加されました。

## <a name="api-list"></a>API リスト

次の表に、Excel JavaScript API 要件セット1.3 の Api を示します。 Excel JavaScript API 要件セット1.3 またはそれ以前でサポートされているすべての Api の API リファレンスドキュメントを表示するには、「 [要件セット1.3 またはそれ以前の Excel api](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)」を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|バインドを削除します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range: Range \| string, bindingtype: Excel. bindingtype, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|特定の範囲に新しいバインドを追加します。|
||[addFromNamedItem (name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromSelection (bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|現在の選択範囲に基づいて新しいバインドを追加します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|現在のピボットテーブルを含んでいるワークシート。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|ピボットテーブルを更新します。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|コレクション内のすべてのピボットテーブルを更新します。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|現在の範囲の表示されている行を表します。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[数式](/javascript/api/excel/excel.rangeview#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|R1C1 スタイル表記の数式を表します。|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|現在の RangeView に関連付けられている親の範囲を取得します。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|指定したセルの Excel の数値書式コードを表します。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|RangeView のセル アドレスを表します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|表示されている列の数を返します。 読み取り専用です。|
||[index](/javascript/api/excel/excel.rangeview#index)|RangeView のインデックスを表す値を返します。 読み取り専用です。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|表示されている行の数を返します。 読み取り専用です。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|範囲に関連付けられている範囲ビューのコレクションを表します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangeview#values)|指定した範囲ビューの Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|インデックスを使用して、RangeView 行を取得します。 0 を起点とする番号になります。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[表](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[ブック](/javascript/api/excel/excel.workbook)|[テーブル](/javascript/api/excel/excel.workbook#pivottables)|ブックに関連付けられているピボットテーブルのコレクションを表します。 読み取り専用です。|
|[ワークシート](/javascript/api/excel/excel.worksheet)|[テーブル](/javascript/api/excel/excel.worksheet#pivottables)|ワークシートの一部になっているピボットテーブルのコレクション。 読み取り専用です。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
