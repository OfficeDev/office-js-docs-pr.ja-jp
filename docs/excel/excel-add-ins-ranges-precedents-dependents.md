---
title: JavaScript API を使用して数式の前例と依存Excel処理する
description: JavaScript API の Excelを使用して、数式の前例と依存を取得する方法について説明します。
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: defab38c2096fa00051d5246d734e0bae592f46b
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868709"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>JavaScript API を使用して数式の前例と依存Excel取得する

Excelは、多くの場合、他のセルを参照します。 これらのクロスセル参照は、"前例" および "依存" と呼ばれる。 前例は、数式にデータを提供するセルです。 従属とは、他のセルを参照する数式を含むセルです。 セル間のリレーションシップにExcelする機能の詳細については、「数式とセル間のリレーションシップを表示する[」を参照してください](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)。

セルには前例のセルを含め、その前のセルには独自の前例セルを含めできます。 "直接の前例" は、親子関係の親の概念と同様に、このシーケンス内のセルの最初の前のグループです。 "直接依存" は、親子関係の子と同様に、シーケンス内のセルの最初の依存グループです。 ブック内の他のセルを参照するが、リレーションシップが親子関係ではないセルは、直接依存または直接の前例ではありません。

この記事では、JavaScript API を使用して数式の直接の前例と直接依存を取得するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Range Object (JavaScript API for Excel)」を参照してください](/javascript/api/excel/excel.range)。

## <a name="get-the-direct-precedents-of-a-formula"></a>数式の直接の前例を取得する

[Range.getDirectPrecedents](/javascript/api/excel/excel.range#getDirectPrecedents__)を使用して数式の直接の先行セルを検索します。 `Range.getDirectPrecedents` オブジェクトを返 `WorkbookRangeAreas` します。 このオブジェクトには、ブック内のすべての直接の前例のアドレスが含まれます。 このオブジェクトには、少なくとも 1 つの数式の前例を含 `RangeAreas` むワークシートごとに個別のオブジェクトがあります。 オブジェクトの操作の詳細については、「複数の範囲を同時に操作する」を参照Excel `RangeAreas` [アドインを参照してください](excel-add-ins-multiple-ranges.md)。

次のスクリーンショットは、UI の [前例のトレース] ボタンを選択した結果Excel示しています。 このボタンは、前のセルから選択したセルに矢印を描画します。 選択したセル **E3** には数式 "=C3 * D3" が含まれているので **、C3** と **D3** の両方が先行セルです。 UI ボタンExcel異なり、 `getDirectPrecedents` メソッドは矢印を描画しない。

![UI の矢印トレースの先行セルExcelします。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> メソッド `getDirectPrecedents` は、ブック間で先行セルを取得できない。

次のコード サンプルでは、アクティブな範囲の直接の前例を取得し、それらの前のセルの背景色を黄色に変更します。

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a>数式の直接依存を取得する

[Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__)を使用して数式の直接依存セルを検索します。 同様 `Range.getDirectPrecedents` に `Range.getDirectDependents` 、オブジェクトも返 `WorkbookRangeAreas` します。 このオブジェクトには、ブック内のすべての直接依存のアドレスが含まれます。 このオブジェクトには、少なくとも 1 つの数式に依存 `RangeAreas` するワークシートごとに個別のオブジェクトがあります。 オブジェクトの操作の詳細については、「複数の範囲を同時に操作する」を参照Excel `RangeAreas` [アドインを参照してください](excel-add-ins-multiple-ranges.md)。

次のスクリーンショットは、UI の [トレース依存] ボタンを選択した結果Excel示しています。 このボタンは、依存セルから選択したセルに矢印を描画します。 選択したセル **D3** には、セル **E3** が従属セルとして含されます。 **E3 には** 、"=C3 * D3" という数式が含まれる。 UI ボタンExcel異なり、 `getDirectDependents` メソッドは矢印を描画しない。

![UI 内の依存セルをExcelします。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> メソッド `getDirectDependents` は、ブック間で依存セルを取得できない。

次のコード サンプルは、アクティブな範囲の直接の依存を取得し、それらの依存セルの背景色を黄色に変更します。

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
