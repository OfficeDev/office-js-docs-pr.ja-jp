---
title: JavaScript API を使用して数式の前例と依存Excel処理する
description: JavaScript API の Excelを使用して、数式の前例と依存を取得する方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9d686b416b271dce81ee072a98f8cb9e1dac65b2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744089"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>JavaScript API を使用して数式の前例と依存Excel取得する

Excelは、多くの場合、他のセルを参照します。 これらのクロスセル参照は、"前例" および "依存" と呼ばれる。 前例は、数式にデータを提供するセルです。 従属とは、他のセルを参照する数式を含むセルです。 セル間のリレーションシップにExcelする機能の詳細については、「数式とセル間のリレーションシップを表示する[」を参照してください](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)。

前例のセルには、独自の前例セルがあります。 前例のこのチェーン内のすべての先行セルは、元のセルの前例です。 被扶養者に対して同じ関係が存在します。 別のセルの影響を受けるセルは、そのセルに依存します。 "直接の前例" は、親子関係の親の概念と同様に、このシーケンス内のセルの最初の前のグループです。 "直接依存" は、親子関係の子と同様に、シーケンス内のセルの最初の依存グループです。

この記事では、JavaScript API を使用して数式の前例と依存を取得するコード Excel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「[Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。

## <a name="get-the-precedents-of-a-formula"></a>数式の前例を取得する

[Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)) を使用して数式の先行セルを検索します。 `Range.getPrecedents` オブジェクトを返 `WorkbookRangeAreas` します。 このオブジェクトには、ブック内のすべての前例のアドレスが含まれます。 このオブジェクトには、少 `RangeAreas` なくとも 1 つの数式の前例を含むワークシートごとに個別のオブジェクトがあります。 オブジェクトの詳細については、「`RangeAreas`複数の範囲を同時に処理する」を参照Excel[アドインを参照してください](excel-add-ins-multiple-ranges.md)。

数式の直接の先行セルのみを検索するには、 [Range.getDirectPrecedents を使用します](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))。 `Range.getDirectPrecedents` 同様に動作 `Range.getPrecedents` し、直接の `WorkbookRangeAreas` 前例のアドレスを含むオブジェクトを返します。

次のスクリーンショットは、UI の [前例のトレース] ボタンを選択した結果Excel示しています。 このボタンは、前のセルから選択したセルに矢印を描画します。 選択したセル **E3** には数式 "=C3 * **D3**" が含まれているので、**C3 と D3** の両方が先行セルです。 UI ボタンExcel、メソッド`getPrecedents``getDirectPrecedents`は矢印を描画しない。

![UI の矢印トレースの先行セルExcelします。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> and `getPrecedents` メソッド `getDirectPrecedents` は、ブック全体で先行セルを取得しない。

次のコード サンプルは、and メソッドを使用する方法 `Range.getPrecedents` を `Range.getDirectPrecedents` 示しています。 サンプルは、アクティブな範囲の前例を取得し、それらの前のセルの背景色を変更します。 直接の前例セルの背景色は黄色に設定され、他の前のセルの背景色はオレンジ色に設定されます。

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
await Excel.run(async (context) => {
  let range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  let precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  let directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  await context.sync();

  console.log(`All precedent cells of ${range.address}:`);
  
  // Use the precedents API to loop through all precedents of the active cell.
  for (let i = 0; i < precedents.areas.items.length; i++) {
    // Highlight and print out the address of all precedent cells.
    precedents.areas.items[i].format.fill.color = "Orange";
    console.log(`  ${precedents.areas.items[i].address}`);
  }

  console.log(`Direct precedent cells of ${range.address}:`);

  // Use the direct precedents API to loop through direct precedents of the active cell.
  for (let i = 0; i < directPrecedents.areas.items.length; i++) {
    // Highlight and print out the address of each direct precedent cell.
    directPrecedents.areas.items[i].format.fill.color = "Yellow";
    console.log(`  ${directPrecedents.areas.items[i].address}`);
  }
});
```

## <a name="get-the-direct-dependents-of-a-formula"></a>数式の直接依存を取得する

[Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) を使用して数式の直接依存セルを検索します。 同様 `Range.getDirectPrecedents`に、 `Range.getDirectDependents` オブジェクトも返 `WorkbookRangeAreas` します。 このオブジェクトには、ブック内のすべての直接依存のアドレスが含まれます。 このオブジェクトには、少 `RangeAreas` なくとも 1 つの数式に依存するワークシートごとに個別のオブジェクトがあります。 オブジェクトの操作の詳細については、「`RangeAreas`複数の範囲を同時に操作する」を参照Excel[してください](excel-add-ins-multiple-ranges.md)。

次のスクリーンショットは、UI の [トレース依存] ボタンを選択した結果Excel示しています。 このボタンは、依存セルから選択したセルに矢印を描画します。 選択したセル **D3** には、セル **E3** が従属セルとして含されます。 **E3 には** 、"=C3 * D3" という数式が含まれる。 UI ボタンExcel異なり、メソッド`getDirectDependents`は矢印を描画しない。

![UI 内の依存セルをExcelします。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> この `getDirectDependents` メソッドは、ブック全体の依存セルを取得しない。

次のコード サンプルは、アクティブな範囲の直接の依存を取得し、それらの依存セルの背景色を黄色に変更します。

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
await Excel.run(async (context) => {
    // Direct dependents are cells that contain formulas that refer to other cells.
    let range = context.workbook.getActiveCell();
    let directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    await context.sync();
    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
