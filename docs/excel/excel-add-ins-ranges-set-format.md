---
title: JavaScript API を使用して範囲のExcel設定する
description: JavaScript API の Excelを使用して範囲の形式を設定する方法について説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: b3cd37aa4415627c2d549b1d11bdb6276388868b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745355"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲Excel設定する

この記事では、JavaScript API を使用して範囲のセルのフォントの色、塗りつぶしの色、数値の形式を設定するExcelします。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>フォントの色と塗りつぶしの色を設定する

次のコード サンプルは、範囲 **B2：E2** のセルのフォントの色と塗りつぶしの色を設定します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    await context.sync();
});
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>フォントの色と塗りつぶしの色を設定する前の範囲内のデータ

![書式が設定Excel前のデータ。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>フォントの色と塗りつぶしの色を設定した後の範囲内のデータ

![書式が設定Excel後のデータ。](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>数値の書式を設定する

次のコード サンプルは、範囲 **D3：E5** のセルの数値を書式を設定します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    await context.sync();
});
```

### <a name="data-in-range-before-number-format-is-set"></a>数値の書式を設定する前の範囲内のデータ

![数値の形式Excel前のデータ。](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>数値の書式を設定した後の範囲内のデータ

![数値の形式Excel後のデータ。](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を設定Excel取得する](excel-add-ins-ranges-set-get.md)
- [JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する](excel-add-ins-ranges-set-get-values.md)
