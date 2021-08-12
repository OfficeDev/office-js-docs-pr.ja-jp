---
title: JavaScript API を使用して範囲Excel挿入する
description: JavaScript API を使用してセル範囲を挿入するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad75b8c107005777047418ff9a1824665552cb5cca06c1e858f3645172f12e7c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084649"
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
