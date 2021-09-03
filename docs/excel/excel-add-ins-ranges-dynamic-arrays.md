---
title: JavaScript API を使用して動的配列と範囲のスピルExcel処理する
description: JavaScript API を使用して動的配列と範囲のスピルを処理するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d95546b4cff3f0ba7410d9ceaa73e19b7e684985
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868688"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>JavaScript API を使用して動的配列とスピルExcel処理する

この記事では、JavaScript API を使用して動的配列と範囲のスピルを処理するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。

## <a name="dynamic-arrays"></a>動的配列

一部Excelは動的配列[を返します](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。 数式の元のセルの外側にある複数のセルの値を入力します。 この値のオーバーフローは"スピル" と呼ばれます。 アドインは [、Range.getSpillingToRange](/javascript/api/excel/excel.range#getSpillingToRange__) メソッドを使用して流出に使用される範囲を検索できます。 [*OrNullObject バージョンも用意されています](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。

次のサンプルは、セルに範囲の内容をコピーする基本的な数式を示しています。これは隣接するセルに流出します。 その後、アドインは流出を含む範囲をログに記録します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a>範囲の流出

[Range.getSpillParent](/javascript/api/excel/excel.range#getSpillParent__)メソッドを使用して、特定のセルにこぼれるセルを検索します。 range オブジェクト `getSpillParent` が 1 つのセルの場合にのみ機能します。 複数 `getSpillParent` のセルを含む範囲を呼び出す場合、エラーがスローされます (または null の範囲が返されます `Range.getSpillParentOrNullObject` )。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
