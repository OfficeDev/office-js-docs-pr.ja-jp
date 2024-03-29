---
title: JavaScript API を使用して範囲Excel挿入する
description: JavaScript API を使用してセル範囲を挿入するExcel説明します。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ce559d0726b7d69c5f4c8c6d00a4e714c04df735
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745213"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>JavaScript API を使用してセル範囲をExcelする

この記事では、JavaScript API を使用してセル範囲を挿入するコード サンプルExcel示します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、次のExcel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>セルの範囲を挿入する

次のコードサンプルは、場所 **B4:E4** にセルの範囲を挿入し、他のセルを下にシフトして、新しいセルのためのスペースを提供します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### <a name="data-before-range-is-inserted"></a>範囲を挿入する前のデータ

![範囲が挿入Excel前のデータ。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>範囲を挿入した後のデータ

![範囲が挿入Excel後のデータ。](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲をクリアまたはExcelする](excel-add-ins-ranges-clear-delete.md)
