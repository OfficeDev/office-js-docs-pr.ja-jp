---
title: Excel JavaScript API を使用して範囲を挿入する
description: Excel JavaScript API を使用してセル範囲を挿入する方法について説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652922"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してセル範囲を挿入する

この記事では、Excel JavaScript API を使用してセル範囲を挿入するコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、Excel.Range クラスを参照してください](/javascript/api/excel/excel.range)。

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

![範囲を挿入する前の Excel のデータ](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>範囲を挿入した後のデータ

![範囲を挿入した後の Excel のデータ](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Excel JavaScript API を使用して範囲をクリアまたは削除する](excel-add-ins-ranges-clear-delete.md)
