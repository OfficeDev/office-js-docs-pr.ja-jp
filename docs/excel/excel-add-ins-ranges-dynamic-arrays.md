---
title: JavaScript API を使用して動的配列と範囲のスピルExcel処理する
description: JavaScript API を使用して動的配列と範囲のスピルを処理するExcel説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4ba4ab2bbce04465bc7db0a75e8ce39a6584a5a8
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745070"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>JavaScript API を使用して動的配列とスピルExcel処理する

この記事では、JavaScript API を使用して動的配列と範囲のスピルを処理するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

## <a name="dynamic-arrays"></a>動的配列

一部Excelは動的配列[を返します](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。 数式の元のセルの外側にある複数のセルの値を入力します。 この値のオーバーフローは"スピル" と呼ばれます。 アドインは、 [Range.getSpillingToRange](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1)) メソッドを使用して流出に使用される範囲を検索できます。 *[OrNullObject バージョンがあります](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)`Range.getSpillingToRangeOrNullObject`。

次のサンプルは、セルに範囲の内容をコピーする基本的な数式を示しています。これは隣接するセルに流出します。 その後、アドインは流出を含む範囲をログに記録します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    let targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    let spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    await context.sync();

    // This will log the range as "G4:J4".
    console.log(`Copying the table headers spilled into ${spillRange.address}.`);
});
```

## <a name="range-spilling"></a>範囲の流出

[Range.getSpillParent](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1)) メソッドを使用して、特定のセルにこぼれるセルを検索します。 range オブジェクトが `getSpillParent` 1 つのセルの場合にのみ機能します。 複数 `getSpillParent` のセルを含む範囲を呼び出す場合、エラーがスローされます (または null の範囲が返されます `Range.getSpillParentOrNullObject`)。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
