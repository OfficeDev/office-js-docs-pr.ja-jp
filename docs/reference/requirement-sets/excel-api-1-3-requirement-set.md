---
title: Excel JavaScript API 要件セット1.3
description: ExcelApi 1.3 の要件セットの詳細
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4698b0fad3122c8ecf52117c35d4928305d812fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771996"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能

ExcelApi 1.3 には、データバインドと基本的なピボットテーブルアクセスのサポートが追加されました。

## <a name="api-list"></a>API リスト

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|バインドを削除します。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range: Range \| String, bindingtype: "range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|特定の範囲に新しいバインドを追加します。|
||[add (range: Range \| String, bindingtype: Excel. bindingtype, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|特定の範囲に新しいバインドを追加します。|
||[addFromNamedItem (name: string, bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromNamedItem (name: string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|
||[addFromSelection (bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|現在の選択範囲に基づいて新しいバインドを追加します。|
||[addFromSelection (bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|現在の選択範囲に基づいて新しいバインドを追加します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|現在のピボットテーブルを含んでいるワークシート。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|ピボットテーブルを更新します。|
||[set (properties: Excel. PivotTable)](/javascript/api/excel/excel.pivottable#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: PivotTableUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.pivottable#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|名前に基づいてピボットテーブルを取得します。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|コレクション内のすべてのピボットテーブルを更新します。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|コレクション内の各アイテムの場合: ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|コレクション内の各アイテムについて: 現在のピボットテーブルを含むワークシート。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|ピボットテーブルの名前。|
|[ピボットのオプション](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|ピボットテーブルの名前。|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|現在のピボットテーブルを含んでいるワークシート。|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|ピボットテーブルの名前。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|現在の範囲の表示されている行を表します。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|A1 スタイル表記の数式を表します。|
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
||[set (properties: Excel. RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|既存の読み込まれたオブジェクトに基づいて、オブジェクトに複数のプロパティを設定します。|
||[set (properties: RangeViewUpdateData, options?: Officeextension.error)](/javascript/api/excel/excel.rangeview#set-properties--options-)|一度に1つのオブジェクトの複数のプロパティを設定します。 適切なプロパティを持つプレーンオブジェクト、または同じ種類の別の API オブジェクトのいずれかを渡すことができます。|
||[values](/javascript/api/excel/excel.rangeview#values)|指定した範囲ビューの Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|インデックスを使用して、RangeView 行を取得します。 0 を起点とする番号になります。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[RangeViewCollectionLoadOptions](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|コレクション内の各アイテムについて: RangeView のセルのアドレスを表します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|コレクション内の各アイテムについて: 表示されている列の数を返します。 読み取り専用です。|
||[formulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|コレクション内の各アイテムについて: A1 形式の表記で数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|コレクション内の各項目について、: ユーザーの言語と書式設定ロケールで、A1 形式の表記法の数式を表します。  たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|コレクション内の各項目について、: R1C1 形式の表記法で数式を表します。|
||[index](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|コレクション内の各アイテムについて: RangeView のインデックスを表す値を返します。 読み取り専用です。|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|コレクション内の各アイテムについて: 指定されたセルの Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|コレクション内の各アイテムについて: 表示されている行の数を返します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|コレクション内の各項目について: 指定された範囲のテキスト値。 テキスト値は、セルの幅には依存しません。 Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。 読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|コレクション内の各アイテムについて: 各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|コレクション内の各アイテムについて: 指定された範囲ビューの生の値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeViewData](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|RangeView のセル アドレスを表します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangeviewdata#columncount)|表示されている列の数を返します。 読み取り専用です。|
||[formulas](/javascript/api/excel/excel.rangeviewdata#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|R1C1 スタイル表記の数式を表します。|
||[index](/javascript/api/excel/excel.rangeviewdata#index)|RangeView のインデックスを表す値を返します。 読み取り専用です。|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|指定したセルの Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|表示されている行の数を返します。 読み取り専用です。|
||[rows](/javascript/api/excel/excel.rangeviewdata#rows)|範囲に関連付けられている範囲ビューのコレクションを表します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|指定した範囲ビューの Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeViewLoadOptions](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|RangeView のセル アドレスを表します。 読み取り専用です。|
||[columnCount](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|表示されている列の数を返します。 読み取り専用です。|
||[formulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|R1C1 スタイル表記の数式を表します。|
||[index](/javascript/api/excel/excel.rangeviewloadoptions#index)|RangeView のインデックスを表す値を返します。 読み取り専用です。|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|指定したセルの Excel の数値書式コードを表します。|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|表示されている行の数を返します。 読み取り専用です。|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|各セルのデータの種類を表します。 読み取り専用です。|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|指定した範囲ビューの Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[RangeViewUpdateData](/javascript/api/excel/excel.rangeviewupdatedata)|[formulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|A1 スタイル表記の数式を表します。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, 1.5)" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|R1C1 スタイル表記の数式を表します。|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|指定したセルの Excel の数値書式コードを表します。|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|指定した範囲ビューの Raw 値を表します。 返されるデータの型は、文字列、数値、ブール値のいずれかになります。 エラーが含まれているセルは、エラー文字列を返します。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|コレクション内の各アイテムについて: 最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|コレクション内の各アイテムについて: 最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|コレクション内の各アイテムについて: 列に縞模様の書式が表示されているかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|コレクション内の各アイテムについて: 行に縞模様の書式が設定されているかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|コレクション内の各アイテムについて: フィルターボタンを各列ヘッダーの上部に表示するかどうかを示します。 これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|最初の列に特別な書式設定が含まれているかどうかを示します。|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|最後の列に特別な書式設定が含まれているかどうかを示します。|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|
|[Workbook](/javascript/api/excel/excel.workbook)|[テーブル](/javascript/api/excel/excel.workbook#pivottables)|ブックに関連付けられているピボットテーブルのコレクションを表します。 読み取り専用。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[テーブル](/javascript/api/excel/excel.workbookdata#pivottables)|ブックに関連付けられているピボットテーブルのコレクションを表します。 読み取り専用。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[テーブル](/javascript/api/excel/excel.worksheet#pivottables)|ワークシートの一部になっているピボットテーブルのコレクション。 読み取り専用です。|
|[ワークシートデータ](/javascript/api/excel/excel.worksheetdata)|[テーブル](/javascript/api/excel/excel.worksheetdata#pivottables)|ワークシートの一部になっているピボットテーブルのコレクション。 読み取り専用です。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンスドキュメント](/javascript/api/excel)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)
