---
title: JavaScript API を使用して範囲Excel挿入する
description: JavaScript API を使用してセル範囲を挿入するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e14aeb030e01dbf170d3acc1edd4952b4989a557
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149839"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>JavaScript API を使用してセル範囲をExcelする

この記事では、JavaScript API を使用してセル範囲を挿入するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については、 `Range` 次の[Excel。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>セルの範囲を挿入する

次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>範囲を挿入する前のデータ

![範囲が挿入Excel前のデータ。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>範囲を挿入した後のデータ

![範囲が挿入Excel後のデータ。](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲をクリアまたはExcelする](excel-add-ins-ranges-clear-delete.md)
