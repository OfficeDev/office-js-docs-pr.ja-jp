---
title: JavaScript API を使用して範囲をクリアまたはExcelする
description: JavaScript API を使用して範囲をクリアまたは削除するExcel説明します。
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7336a0e6485ce502216818b4a8cd077fed0069c3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340709"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲をクリアまたはExcelする

この記事では、JavaScript API を使用して範囲をクリアおよび削除するExcelを提供します。 オブジェクトでサポートされるプロパティと`Range`メソッドの完全な一覧については、「Excel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>セルの範囲をクリアする

次のコード サンプルは、範囲 **E2：E5** のセルの内容と書式をすべてクリアします。  

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("E2:E5");

    range.clear();

    await context.sync();
});
```

### <a name="data-before-range-is-cleared"></a>範囲をクリアする前のデータ

![範囲がクリアExcel前のデータ。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>範囲をクリアした後のデータ

![範囲がExcel後のデータ。](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>セルの範囲を削除する

次のコード サンプルでは、 **範囲 B4:E4** のセルを削除し、他のセルを上に移動して、削除されたセルで空いた領域を埋める。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
});
```

### <a name="data-before-range-is-deleted"></a>範囲を削除する前のデータ

![範囲がExcel前のデータ。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>範囲を削除した後のデータ

![範囲がExcelされた後のデータ。](../images/excel-ranges-after-delete.png)

## <a name="see-also"></a>関連項目

- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲を設定Excel取得する](excel-add-ins-ranges-set-get.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
