---
title: JavaScript API を使用して選択した範囲を設定Excel取得する
description: JavaScript API の Excelを使用して、JavaScript API を使用して範囲を設定および取得するExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075776"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>JavaScript API を使用して範囲を設定Excel取得する

この記事では、JavaScript API を使用して範囲を設定および取得するExcel説明します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、「Excel。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>選択範囲を設定する

次のコード サンプルは、作業中のワークシートの範囲 **B2:E6** を選択します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>選択範囲 B2:E6

![[選択した範囲] Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>選択範囲を取得する

次のコード サンプルでは、選択した範囲を取得し、そのプロパティを読み込み、コンソール `address` にメッセージを書き込みます。

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して範囲の値、テキスト、または数式を設定Excel取得する](excel-add-ins-ranges-set-get-values.md)
- [JavaScript API を使用して範囲Excel設定する](excel-add-ins-ranges-set-format.md)
