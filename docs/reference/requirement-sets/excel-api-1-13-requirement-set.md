---
title: Excel JavaScript API 要件セット 1.13
description: ExcelApi 1.13 要件セットの詳細。
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 5d7358c35dc4560bf5478bb9ad9970fc364a1b6a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747043"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>JavaScript API 1.13 Excel新機能

ExcelApi 1.13 では、Base64 エンコード文字列からブックにワークシートを挿入するメソッドと、ブックのアクティブ化を検出するイベントが追加されました。 また、API を追加して数式の変更を追跡し、数式の直接依存セルを特定することで、範囲内の数式のサポートを増やしました。 さらに、代替テキスト、スタイル、空のセル管理用の PivotLayout API を追加することで、ピボットテーブルのサポートを拡張しました。

| 機能領域 | 説明 | 関連オブジェクト |
|:--- |:--- |:--- |
| [数式の変更イベント](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | 変更の原因となるイベントのソースと種類を含む、数式の変更を追跡します。 | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [数式の依存](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | 数式の直接依存セルを見つける。 | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [ワークシートの挿入](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | 別のブックのワークシートを現在のブックに Base64 エンコード文字列として挿入します。 | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [ピボットテーブル ピボットレイアウト](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Alt テキストと空のセル管理の新しいサポートを含む、PivotLayout クラスの拡張。 | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>API リスト

次の表に、JavaScript API 要件セット 1.13 Excel API の一覧を示します。 Excel JavaScript API 要件セット 1.13 以前でサポートされているすべての API の API リファレンス ドキュメントを表示するには、要件セット [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true) 以前の Excel API を参照してください。

| クラス | フィールド | 説明 |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-celladdress-member)|変更された数式を含むセルのアドレス。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-previousformula-member)|変更前の数式を表します。|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-positiontype-member)|新しいワークシートの現在のブック内の挿入位置。|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-relativeto-member)|パラメーターに対して参照されている現在のブック内のワークシート `WorksheetPositionType` 。|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member)|挿入する個々のワークシートの名前。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttextdescription-member)|ピボットテーブルの代替テキストの説明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttexttitle-member)|ピボットテーブルの代替テキスト タイトル。|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-displayblanklineaftereachitem-member(1))|各項目の後に空白行を表示するかどうかを設定します。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-emptycelltext-member)|ピボットテーブル内の空のセルに自動的に入力されるテキスト 。`fillEmptyCells == true`|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-fillemptycells-member)|ピボットテーブルの空のセルに、 を設定するかどうかを指定します `emptyCellText`。|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-repeatallitemlabels-member(1))|ピボットテーブルのすべてのフィールドで[すべてのアイテム ラベルを繰り返す] 設定を設定します。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showfieldheaders-member)|ピボットテーブルにフィールド ヘッダー (フィールド キャプションとフィルター ドロップダウン) を表示するかどうかを指定します。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refreshonopen-member)|ブックが開くとピボットテーブルが更新されるかどうかを指定します。|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))|同じワークシート `WorkbookRangeAreas` または複数のワークシート内のセルのすべての直接依存を含む範囲を表すオブジェクトを返します。|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1))|指定された方向に基づいて、現在の範囲と範囲の端までの範囲オブジェクトを返します。|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getmergedareasornullobject-member(1))|この範囲内の結合領域を表す RangeAreas オブジェクトを返します。|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1))|指定された方向に対応するデータ領域のエッジ セルである範囲オブジェクトを返します。|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1))|テーブルのサイズを新しい範囲に変更します。|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1))|指定したワークシートをソース ブックから現在のブックに挿入します。|
||[onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member)|ブックがアクティブ化されると発生します。|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#excel-excel-workbookactivatedeventargs-type-member)|イベントの種類を取得します。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|このワークシートで 1 つ以上の数式が変更された場合に発生します。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member)|このコレクションのワークシートで 1 つ以上の数式が変更された場合に発生します。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member)|変更された数式の `FormulaChangedEventDetail` 詳細を含むオブジェクトの配列を取得します。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-source-member)|イベントのソース。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-type-member)|イベントの種類を取得します。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-worksheetid-member)|数式が変更されたワークシートの ID を取得します。|

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
