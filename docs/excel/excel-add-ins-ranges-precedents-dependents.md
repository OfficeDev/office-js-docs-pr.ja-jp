---
title: Excel JavaScript API を使用して数式の前例と依存を操作する
description: Excel JavaScript API を使用して、数式の前例と依存を取得する方法について説明します。
ms.date: 05/19/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ca432b7eb6825781960e995af2ed2193c7caa5e2
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628097"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Excel JavaScript API を使用して数式の前例と依存を取得する

Excel数式は、多くの場合、他のセルを参照します。 これらのセル間参照は、"前例" と "依存" と呼ばれます。 前例は、数式にデータを提供するセルです。 依存するセルは、他のセルを参照する数式を含むセルです。 セル間のリレーションシップに関連するExcel機能の詳細については、「[数式とセル間のリレーションシップを表示する](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)」を参照してください。

先行セルには、独自の先行セルを含む場合があります。 この前例のチェーン内のすべての先行セルは、元のセルの前例のままです。 依存関係にも同じリレーションシップが存在します。 別のセルの影響を受けるセルは、そのセルに依存します。 "直接の前例" は、親と子の関係における親の概念と同様に、このシーケンス内の最初のセルグループです。 "直接依存" は、親子関係の子と同様に、シーケンス内のセルの最初の依存グループです。

この記事では、Excel JavaScript API を使用して、数式の前例と依存を取得するコード サンプルを提供します。 オブジェクトがサポートする`Range`プロパティとメソッドの完全な一覧については、「[Range オブジェクト (javaScript API for Excel)」](/javascript/api/excel/excel.range)を参照してください。

## <a name="get-the-precedents-of-a-formula"></a>数式の前例を取得する

[Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)) を使用して、数式の前のセルを見つけます。 `Range.getPrecedents` はオブジェクトを `WorkbookRangeAreas` 返します。 このオブジェクトには、ブック内のすべての前例のアドレスが含まれます。 ワークシートごとに、少なくとも 1 つの数式の前例を含む個別 `RangeAreas` のオブジェクトがあります。 オブジェクトの詳細`RangeAreas`については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」を参照してください。

数式の直接の先行セルのみを検索するには、 [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)) を使用します。 `Range.getDirectPrecedents`は、直接の前例の`WorkbookRangeAreas`アドレスを含むオブジェクトと同様`Range.getPrecedents`に動作し、返します。

次のスクリーンショットは、Excel UI で [**Trace Precedents**] ボタンを選択した結果を示しています。 このボタンは、前のセルから選択したセルに矢印を描画します。 選択したセル **E3** には数式 "=C3 * D3" が含まれているため、**C3 と D3** の両方が先行セルです。 Excel UI ボタンとは異なり、`getPrecedents`メソッドと`getDirectPrecedents`メソッドは矢印を描画しません。

![Excel UI の前のセルをトレースする矢印。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> および`getDirectPrecedents`メソッドは`getPrecedents`、ブック間で前例のセルを取得しません。

次のコード サンプルは、および`Range.getDirectPrecedents`メソッドを操作する方法を`Range.getPrecedents`示しています。 サンプルは、アクティブな範囲の前例を取得し、その前のセルの背景色を変更します。 直接先行セルの背景色は黄色に設定され、他の先行セルの背景色はオレンジ色に設定されます。

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

## <a name="get-the-dependents-of-a-formula"></a>数式の依存を取得する

[Range.getDependents](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1)) を使用して数式の依存セルを見つけます。 同様に `Range.getPrecedents`、 `Range.getDependents` オブジェクトも返します `WorkbookRangeAreas` 。 このオブジェクトには、ブック内のすべての依存要素のアドレスが含まれます。 これには、少なくとも 1 つの数式に依存するワークシートごとに個別 `RangeAreas` のオブジェクトがあります。 オブジェクトの`RangeAreas`操作の詳細については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」を参照してください。

数式の直接依存セルのみを検索するには、 [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) を使用します。 `Range.getDirectDependents`は、直接依存のアドレスを`WorkbookRangeAreas`含むオブジェクトと同様`Range.getDependents`に動作し、返します。

次のスクリーンショットは、Excel UI で **[トレース依存**] ボタンを選択した結果を示しています。 このボタンは、選択したセルから依存セルへの矢印を描画します。 選択したセル **D3** には、依存セルとして **セル E3** があります。 **E3** には、"=C3 * D3" という数式が含まれています。 Excel UI ボタンとは異なり、`getDependents`メソッドと`getDirectDependents`メソッドは矢印を描画しません。

![Excel UI の方向トレース依存セル。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> および`getDirectDependents`メソッドは`getDependents`、ブック間で依存セルを取得しません。

次のコード サンプルでは、アクティブな範囲の直接の依存を取得し、それらの依存セルの背景色を黄色に変更します。

次のコード サンプルは、および`Range.getDirectDependents`メソッドを操作する方法を`Range.getDependents`示しています。 サンプルは、アクティブな範囲の依存を取得し、それらの依存セルの背景色を変更します。 直接依存セルの背景色は黄色に設定され、他の依存セルの背景色はオレンジ色に設定されます。

```js
// This code sample shows how to find and highlight the dependents 
// and direct dependents of the currently selected cell.
await Excel.run(async (context) => {
    let range = context.workbook.getActiveCell();
    // Dependents are all cells that contain formulas that refer to other cells.
    let dependents = range.getDependents();  
    // Direct dependents are the child cells, or the first succeeding group of cells in a sequence of cells that refer to other cells.
    let directDependents = range.getDirectDependents();

    range.load("address");
    dependents.areas.load("address");    
    directDependents.areas.load("address");
    
    await context.sync();

    console.log(`All dependent cells of ${range.address}:`);
    
    // Use the dependents API to loop through all dependents of the active cell.
    for (let i = 0; i < dependents.areas.items.length; i++) {
      // Highlight and print out the addresses of all dependent cells.
      dependents.areas.items[i].format.fill.color = "Orange";
      console.log(`  ${dependents.areas.items[i].address}`);
    }

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
- [Excel JavaScript API を使用してセルを操作する](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
